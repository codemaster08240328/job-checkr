import fs from "node:fs";
import path from "node:path";
import process from "node:process";
import pLimit from "p-limit";

import {
  createSheetsClientFromServiceAccountJson,
  readSheetValues,
  updateSheetValues,
} from "./googleSheets.js";
import { checkUrl, fetchPageText } from "./urlCheck.js";

function parseArgs(argv) {
  const args = new Set(argv.slice(2));
  return {
    dryRun: args.has("--dry-run"),
    help: args.has("--help") || args.has("-h"),
  };
}

function requireEnv(name) {
  const v = process.env[name];
  if (!v) throw new Error(`Missing env var: ${name}`);
  return v;
}

function loadJsonFile(filePath) {
  const abs = path.isAbsolute(filePath)
    ? filePath
    : path.join(process.cwd(), filePath);
  return JSON.parse(fs.readFileSync(abs, "utf8"));
}

function loadServiceAccountJsonFromEnv() {
  // Prefer raw JSON in env to avoid file path issues on Windows.
  const raw = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (raw) return JSON.parse(raw);

  const filePath = process.env.GOOGLE_SERVICE_ACCOUNT_FILE;
  if (!filePath) {
    throw new Error(
      "Set GOOGLE_SERVICE_ACCOUNT_JSON (preferred) or GOOGLE_SERVICE_ACCOUNT_FILE."
    );
  }
  return loadJsonFile(filePath);
}

function ensureHeaderMap(headerRow) {
  const map = new Map();
  for (let i = 0; i < headerRow.length; i++) {
    const key = String(headerRow[i] ?? "").trim();
    if (!key) continue;
    map.set(key, i);
  }
  return map;
}

function toA1Column(n1) {
  // 1-based column number to letters, e.g. 1->A, 27->AA
  let n = n1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function fromA1Column(letters) {
  // Letters to 1-based column number, e.g. A->1, AA->27
  const s = String(letters || "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(s)) throw new Error(`Invalid column letters: ${letters}`);
  let n = 0;
  for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

function buildUpdateRange(sheetName, startRow1, startCol1, numRows, numCols) {
  const startCol = toA1Column(startCol1);
  const endCol = toA1Column(startCol1 + numCols - 1);
  const endRow = startRow1 + numRows - 1;
  return `${sheetName}!${startCol}${startRow1}:${endCol}${endRow}`;
}

function parseA1ColSpan(a1) {
  // Supports:
  // - Sheet!A1:Z
  // - Sheet!A1:Z99
  // - Sheet!A:Z          (rowless column span)
  // - Sheet!A78:Z        (no end row)
  const m =
    /^([^!]+)!([A-Za-z]+)(\d+)?:([A-Za-z]+)(\d+)?$/.exec(String(a1));
  if (!m) {
    throw new Error(
      `Unsupported A1 range "${a1}". Expected format like "Jobs!A1:Z", "Jobs!A1:Z200", or "Jobs!A:Z".`
    );
  }
  const sheetName = m[1];
  const startCol1 = fromA1Column(m[2]);
  const startRow1 = m[3] ? Number(m[3]) : null;
  const endCol1 = fromA1Column(m[4]);
  const endRow1 = m[5] ? Number(m[5]) : null;

  if (startRow1 != null && (!Number.isFinite(startRow1) || startRow1 <= 0)) {
    throw new Error(`Invalid start row in range "${a1}"`);
  }
  if (endRow1 != null && (!Number.isFinite(endRow1) || endRow1 <= 0)) {
    throw new Error(`Invalid end row in range "${a1}"`);
  }
  if (endCol1 < startCol1) throw new Error(`Invalid column span in "${a1}"`);
  if (startRow1 != null && endRow1 != null && endRow1 < startRow1) {
    throw new Error(`Invalid row span in "${a1}"`);
  }

  return { sheetName, startCol1, startRow1, endCol1, endRow1 };
}

function parseSimpleA1Range(a1) {
  // Supports: Sheet!A1:Z, Sheet!A1:Z99, Sheet!A78:Z
  // Does NOT support: Sheet!A:Z (rowless) because callers need a concrete startRow1.
  const meta = parseA1ColSpan(a1);
  if (meta.startRow1 == null) {
    throw new Error(
      `Unsupported A1 range "${a1}". Expected a start row like "Jobs!A1:Z" or "Jobs!A78:Z".`
    );
  }
  return meta;
}

function formatNowIso() {
  return new Date().toISOString();
}

function normalizeKeywords(list) {
  if (!Array.isArray(list)) return [];
  return list
    .map((s) => String(s ?? "").trim())
    .filter(Boolean)
    .map((s) => s.toLowerCase());
}

function findBadKeyword(text, badKeywordsLower) {
  const hay = String(text || "").toLowerCase();
  for (const kw of badKeywordsLower) {
    if (kw && hay.includes(kw)) return kw;
  }
  return "";
}

async function processOneSheet({
  sheets,
  spreadsheetId,
  rangeA1,
  headerRangeA1,
  dataRangeA1,
  headerRow1,
  dataStartRow1,
  dataEndRow1,
  urlColumnName,
  titleColumnName,
  concurrency,
  timeoutMs,
  badKeywordsLower,
  badTitleKeywordsLower,
  dryRun,
}) {
  const hasRowOnlyConfig =
    Number.isFinite(headerRow1) ||
    Number.isFinite(dataStartRow1) ||
    Number.isFinite(dataEndRow1);

  const effectiveRange = headerRangeA1 || dataRangeA1 || rangeA1;
  const primarySheetName = effectiveRange?.split("!")[0];
  if (!primarySheetName) {
    throw new Error(
      `Range must include a sheet name like "Jobs!A1:Z". Got: ${effectiveRange}`
    );
  }

  let header;
  let rows;
  let dataRangeMeta = null;

  // If headerRow1/dataStartRow1 are provided, build headerRangeA1/dataRangeA1 from rangeA1's sheet + column span.
  if (hasRowOnlyConfig) {
    if (!rangeA1) {
      throw new Error(
        "Row-only config requires rangeA1 to define the sheet name and column span, e.g. \"Jobs!A:Z\"."
      );
    }
    const span = parseA1ColSpan(rangeA1);

    const hdrRow = Number.isFinite(headerRow1)
      ? headerRow1
      : span.startRow1 ?? 1;

    // For row-only config, require an explicit dataStart + dataEnd (inclusive).
    const hasDataStart = Number.isFinite(dataStartRow1);
    const hasDataEnd = Number.isFinite(dataEndRow1);
    if (hasDataStart !== hasDataEnd) {
      throw new Error(
        "Row-only config requires BOTH dataStartRow1 and dataEndRow1 (inclusive)."
      );
    }
    if (!hasDataStart || !hasDataEnd) {
      throw new Error(
        "Row-only config requires dataStartRow1 and dataEndRow1."
      );
    }
    const dataStart = dataStartRow1;
    const dataEnd = dataEndRow1;

    if (!Number.isFinite(hdrRow) || hdrRow <= 0) {
      throw new Error(`Invalid headerRow1: ${headerRow1}`);
    }
    if (!Number.isFinite(dataStart) || dataStart <= 0) {
      throw new Error(`Invalid dataStartRow1: ${dataStartRow1}`);
    }
    if (!Number.isFinite(dataEnd) || dataEnd <= 0 || dataEnd < dataStart) {
      throw new Error(
        `Invalid dataEndRow1 (${dataEnd}); must be >= dataStartRow1 (${dataStart}).`
      );
    }

    const startCol = toA1Column(span.startCol1);
    const endCol = toA1Column(span.endCol1);
    const builtHeaderRangeA1 = `${span.sheetName}!${startCol}${hdrRow}:${endCol}${hdrRow}`;
    const builtDataRangeA1 = `${span.sheetName}!${startCol}${dataStart}:${endCol}${dataEnd}`;

    const [headerValues, dataValues] = await Promise.all([
      readSheetValues({
        sheets,
        spreadsheetId,
        rangeA1: builtHeaderRangeA1,
      }),
      readSheetValues({
        sheets,
        spreadsheetId,
        rangeA1: builtDataRangeA1,
      }),
    ]);
    if (headerValues.length === 0) {
      throw new Error("No rows returned from headerRow1/header range.");
    }
    header = headerValues[0];
    rows = dataValues;
    dataRangeMeta = parseSimpleA1Range(builtDataRangeA1);
  } else if (headerRangeA1 && dataRangeA1) {
    const [headerValues, dataValues] = await Promise.all([
      readSheetValues({ sheets, spreadsheetId, rangeA1: headerRangeA1 }),
      readSheetValues({ sheets, spreadsheetId, rangeA1: dataRangeA1 }),
    ]);
    if (headerValues.length === 0) {
      throw new Error("No rows returned from headerRangeA1.");
    }
    header = headerValues[0];
    rows = dataValues;
    dataRangeMeta = parseSimpleA1Range(dataRangeA1);
  } else {
    const values = await readSheetValues({ sheets, spreadsheetId, rangeA1 });
    if (values.length === 0) throw new Error("No rows returned from the range.");
    header = values[0];
    rows = values.slice(1);
    dataRangeMeta = parseSimpleA1Range(rangeA1);
    // If rangeA1 includes the header row, data starts at row+1
    dataRangeMeta = {
      ...dataRangeMeta,
      startRow1: dataRangeMeta.startRow1 + 1,
    };
  }

  const headerMap = ensureHeaderMap(header);
  const urlIdx = headerMap.get(urlColumnName);
  if (urlIdx == null) {
    throw new Error(
      `Could not find URL column "${urlColumnName}" in header row. Found: ${header
        .map((h) => String(h ?? ""))
        .join(", ")}`
    );
  }

  const titleIdx = headerMap.get(titleColumnName) ?? null;

  // Optional: allow skipping checks when Applied is not "Yes" (case-insensitive).
  const appliedIdx =
    headerMap.get("Applied") ??
    headerMap.get("applied") ??
    headerMap.get("APPLIED") ??
    null;

  // Ensure output columns exist (append if missing)
  const outCols = [
    "Approved",
    "Reason",
  ];
  for (const name of outCols) {
    if (!headerMap.has(name)) {
      headerMap.set(name, header.length);
      header.push(name);
    }
  }

  const limit = pLimit(
    Number.isFinite(concurrency) && concurrency > 0 ? concurrency : 8
  );

  const results = await Promise.all(
    rows.map((row, i) =>
      limit(async () => {
        // Title pre-filter: if title contains bad keyword => skip URL checks entirely
        const title = titleIdx != null ? String(row?.[titleIdx] ?? "") : "";
        const badTitleKeyword =
          badTitleKeywordsLower.length > 0
            ? findBadKeyword(title, badTitleKeywordsLower)
            : "";
        if (badTitleKeyword) {
          return {
            rowIndex0: i,
            skipped: true,
            res: {
              inputUrl: row[urlIdx],
              finalUrl: "",
              ok: false,
              status: "",
              error: "",
              elapsedMs: 0,
            },
            badKeyword: "",
            contentError: "",
            titleBadKeyword: badTitleKeyword,
          };
        }

        const appliedVal =
          appliedIdx != null ? String(row?.[appliedIdx] ?? "").trim() : "";
        const isAppliedNo = appliedVal.toLowerCase() !== "yes";

        if (isAppliedNo) {
          return {
            rowIndex0: i,
            skipped: true,
            res: {
              inputUrl: row[urlIdx],
              finalUrl: "",
              ok: false,
              status: "",
              error: "SKIPPED_APPLIED_NO",
              elapsedMs: 0,
            },
            badKeyword: "",
            contentError: "",
            titleBadKeyword: "",
          };
        }

        const url = row[urlIdx];
        const res = await checkUrl(url, { timeoutMs });
        let badKeyword = "";
        let contentError = "";
        if (res.status === 200 && badKeywordsLower.length > 0) {
          const page = await fetchPageText(res.finalUrl || url, { timeoutMs });
          contentError = page.error || "";
          badKeyword = findBadKeyword(page.text, badKeywordsLower);
        }
        return {
          rowIndex0: i,
          skipped: false,
          res,
          badKeyword,
          contentError,
          titleBadKeyword: "",
        };
      })
    )
  );

  // Build updates for the output columns only (keeps sheet stable)
  const outStartCol1 = Math.min(...outCols.map((c) => headerMap.get(c))) + 1;
  const outEndCol1 = Math.max(...outCols.map((c) => headerMap.get(c))) + 1;
  const outNumCols = outEndCol1 - outStartCol1 + 1;

  const outValues = rows.map(() => Array(outNumCols).fill(""));
  for (const { rowIndex0, res, skipped, badKeyword, contentError, titleBadKeyword } of results) {
    const rowOut = outValues[rowIndex0];
    const setCol = (colName, v) => {
      const colIdx0 = headerMap.get(colName);
      const offset = colIdx0 + 1 - outStartCol1;
      rowOut[offset] = v ?? "";
    };

    // Approval rule:
    // - Applied != Yes => Approved No (skip)
    // - HTTP status != 200 => Approved No
    // - If content contains any bad keyword => Approved No
    // - Else => Approved Yes
    const approvedValue =
      skipped || res.status !== 200 || Boolean(badKeyword) || Boolean(titleBadKeyword) ? "No" : "Yes";
    setCol("Approved", approvedValue);

    // Reason:
    // - If skipped (not applied) => leave blank
    // - If status is 4xx => "expired"
    // - If bad keyword hit => the keyword
    // - If status != 200 => status code (e.g. 500)
    // - Else blank (or fetch error if we couldn't fetch page text)
    const is4xx =
      typeof res.status === "number" && res.status >= 400 && res.status < 500;
    const reasonValue = skipped
      ? ""
      : titleBadKeyword
        ? "non matched position title"
        : is4xx
        ? "expired"
        : badKeyword
          ? badKeyword
          : res.status !== 200 && res.status !== ""
            ? String(res.status)
            : contentError
              ? contentError
              : "";
    setCol("Reason", reasonValue);
  }

  // Write header back if we added columns
  const sheetNameForWrites = dataRangeMeta.sheetName;
  const headerRow1 = headerRangeA1
    ? parseSimpleA1Range(headerRangeA1).startRow1
    : dataRangeMeta.startRow1 - 1;

  const headerRange = buildUpdateRange(
    sheetNameForWrites,
    headerRow1,
    dataRangeMeta.startCol1,
    1,
    header.length
  );

  const outRange = buildUpdateRange(
    sheetNameForWrites,
    dataRangeMeta.startRow1,
    dataRangeMeta.startCol1 + (outStartCol1 - 1),
    outValues.length,
    outNumCols
  );

  if (dryRun) {
    return {
      sheetName: sheetNameForWrites,
      headerRange,
      outRange,
      checkedRows: rows.length,
      sample: results.slice(0, 5),
    };
  }

  await updateSheetValues({
    sheets,
    spreadsheetId,
    rangeA1: headerRange,
    values: [header],
  });

  await updateSheetValues({
    sheets,
    spreadsheetId,
    rangeA1: outRange,
    values: outValues,
  });

  return {
    sheetName: sheetNameForWrites,
    headerRange,
    outRange,
    checkedRows: rows.length,
  };
}

async function main() {
  const { dryRun, help } = parseArgs(process.argv);
  if (help) {
    // eslint-disable-next-line no-console
    console.log(`
joblink-checkr

Required env:
  (Option A) CONFIG_FILE=path/to/config.json
  (Option B) GOOGLE_SHEETS_SPREADSHEET_ID, GOOGLE_SHEETS_RANGE, and service account env vars

Optional env:
  URL_COLUMN_NAME       (default: URL)
  CONCURRENCY           (default: 8)
  TIMEOUT_MS            (default: 15000)

This tool expects the first row of the range to be headers.
It will write results into columns (created if missing):
  Approved, Reason

Run:
  npm run check:sheet
  npm run check:sheet:dry
`);
    return;
  }

  const configFile = process.env.CONFIG_FILE;
  const cfg = configFile ? loadJsonFile(configFile) : null;

  const urlColumnName = (
    cfg?.urlColumnName ??
    process.env.URL_COLUMN_NAME ??
    "URL"
  ).trim();
  const titleColumnName = (
    cfg?.titleColumnName ??
    process.env.TITLE_COLUMN_NAME ??
    "Position Title"
  ).trim();
  const concurrency = Number(
    cfg?.concurrency ?? process.env.CONCURRENCY ?? "8"
  );
  const timeoutMs = Number(cfg?.timeoutMs ?? process.env.TIMEOUT_MS ?? "15000");
  const badKeywordsLower = normalizeKeywords(cfg?.badKeywords ?? []);
  const badTitleKeywordsLower = normalizeKeywords(cfg?.badTitleKeywords ?? []);

  const serviceAccountJson = cfg?.serviceAccount ?? loadServiceAccountJsonFromEnv();
  const sheets = createSheetsClientFromServiceAccountJson(serviceAccountJson);

  // Backward-compatible config modes:
  // - New: cfg.spreadsheets: [{ spreadsheetId, sheets: [...] }]
  // - Old: cfg.spreadsheetId + cfg.sheets (multiple tabs within one spreadsheet)
  // - Oldest: env vars for one sheet range

  const spreadsheets = Array.isArray(cfg?.spreadsheets) && cfg.spreadsheets.length > 0
    ? cfg.spreadsheets
    : [
        {
          name: cfg?.name || "default",
          spreadsheetId:
            cfg?.spreadsheetId ?? requireEnv("GOOGLE_SHEETS_SPREADSHEET_ID"),
          sheets: Array.isArray(cfg?.sheets) && cfg.sheets.length > 0
            ? cfg.sheets
            : [
                {
                  name: "default",
                  rangeA1: cfg?.rangeA1 ?? requireEnv("GOOGLE_SHEETS_RANGE"),
                  headerRangeA1: cfg?.headerRangeA1 ?? null,
                  dataRangeA1: cfg?.dataRangeA1 ?? null,
                },
              ],
        },
      ];

  const allSummaries = [];
  for (const book of spreadsheets) {
    const spreadsheetId =
      book?.spreadsheetId ?? requireEnv("GOOGLE_SHEETS_SPREADSHEET_ID");
    const sheetJobs = Array.isArray(book?.sheets) ? book.sheets : [];
    const perSheet = [];
    for (const job of sheetJobs) {
      const summary = await processOneSheet({
        sheets,
        spreadsheetId,
        rangeA1: job.rangeA1 ?? cfg?.rangeA1 ?? requireEnv("GOOGLE_SHEETS_RANGE"),
        headerRangeA1: job.headerRangeA1 ?? cfg?.headerRangeA1 ?? null,
        dataRangeA1: job.dataRangeA1 ?? cfg?.dataRangeA1 ?? null,
        headerRow1: Number(job.headerRow1 ?? cfg?.headerRow1 ?? NaN),
        dataStartRow1: Number(job.dataStartRow1 ?? cfg?.dataStartRow1 ?? NaN),
        dataEndRow1: Number(job.dataEndRow1 ?? cfg?.dataEndRow1 ?? NaN),
        urlColumnName,
        titleColumnName,
        concurrency,
        timeoutMs,
        badKeywordsLower,
        badTitleKeywordsLower,
        dryRun,
      });
      perSheet.push({ name: job.name || summary.sheetName, ...summary });
    }
    allSummaries.push({
      name: book.name || spreadsheetId,
      spreadsheetId,
      sheets: perSheet,
    });
  }

  // eslint-disable-next-line no-console
  console.log(JSON.stringify({ spreadsheets: allSummaries }, null, 2));
}

main().catch((err) => {
  // eslint-disable-next-line no-console
  console.error(err?.stack || err);
  process.exitCode = 1;
});

