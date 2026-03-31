import { fetch } from "undici";

const DEFAULT_UA =
  "joblink-checkr/1.0 (+https://example.invalid; link-checker)";

function normalizeUrl(raw) {
  if (raw == null) return null;
  const s = String(raw).trim();
  if (!s) return null;
  // If someone pasted "www.example.com" without scheme, assume https.
  if (/^www\./i.test(s)) return `https://${s}`;
  return s;
}

function stripHtmlToText(html) {
  return String(html || "")
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, " ")
    .replace(/<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Fetches page HTML and returns a best-effort plain text string.
 * Uses a byte limit to avoid downloading huge pages.
 */
export async function fetchPageText(rawUrl, options = {}) {
  const url = normalizeUrl(rawUrl);
  if (!url) return { finalUrl: "", status: "", text: "", error: "EMPTY_URL" };

  const {
    timeoutMs = 15000,
    maxRedirects = 5,
    userAgent = DEFAULT_UA,
    maxBytes = 1_000_000, // 1MB
  } = options;

  try {
    const res = await fetch(url, {
      method: "GET",
      headers: { "user-agent": userAgent, accept: "text/html,*/*" },
      redirect: "follow",
      maxRedirections: maxRedirects,
      headersTimeout: timeoutMs,
      bodyTimeout: timeoutMs,
    });

    let bytes = 0;
    const chunks = [];
    for await (const chunk of res.body) {
      bytes += chunk.length;
      if (bytes > maxBytes) break;
      chunks.push(chunk);
    }
    try {
      res.body?.destroy?.();
    } catch {
      // ignore
    }

    const html = Buffer.concat(chunks).toString("utf8");
    return {
      finalUrl: res.url ?? url,
      status: res.status,
      text: stripHtmlToText(html),
      error: bytes > maxBytes ? "MAX_BYTES_EXCEEDED" : "",
    };
  } catch (err) {
    const message =
      (err && typeof err === "object" && "message" in err && err.message) ||
      String(err);
    return { finalUrl: "", status: "", text: "", error: message };
  }
}

/**
 * Checks a URL with GET (preferred; catches more real failures), with a HEAD fallback.
 * Returns a compact result object safe to write back to a sheet.
 */
export async function checkUrl(rawUrl, options = {}) {
  const url = normalizeUrl(rawUrl);
  if (!url) {
    return {
      inputUrl: rawUrl ?? "",
      finalUrl: "",
      ok: false,
      status: "",
      error: "EMPTY_URL",
      elapsedMs: 0,
    };
  }

  const {
    timeoutMs = 15000,
    maxRedirects = 5,
    userAgent = DEFAULT_UA,
  } = options;

  const started = Date.now();
  const common = {
    headers: { "user-agent": userAgent, accept: "*/*" },
    redirect: "follow",
    maxRedirections: maxRedirects,
    // Undici uses `bodyTimeout` / `headersTimeout`:
    headersTimeout: timeoutMs,
    bodyTimeout: timeoutMs,
  };

  async function attempt(method) {
    const res = await fetch(url, { method, ...common });
    // Ensure the body is consumed/terminated to avoid resource warnings.
    // For HEAD there is no body; for GET we don't need it.
    try {
      res.body?.cancel?.();
    } catch {
      // ignore
    }

    return {
      inputUrl: rawUrl,
      finalUrl: res.url ?? url,
      ok: res.status >= 200 && res.status < 400,
      status: res.status,
      error: "",
      elapsedMs: Date.now() - started,
    };
  }

  try {
    // Many job sites block HEAD; try GET first.
    return await attempt("GET");
  } catch (err) {
    // Try a HEAD fallback in case GET was blocked but HEAD is allowed (rare).
    try {
      return await attempt("HEAD");
    } catch (err2) {
      const message =
        (err2 && typeof err2 === "object" && "message" in err2 && err2.message) ||
        (err && typeof err === "object" && "message" in err && err.message) ||
        String(err2 || err);

      return {
        inputUrl: rawUrl,
        finalUrl: "",
        ok: false,
        status: "",
        error: message,
        elapsedMs: Date.now() - started,
      };
    }
  }
}

