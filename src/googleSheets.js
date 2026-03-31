import { google } from "googleapis";

/**
 * Creates an authenticated Google Sheets client using a Service Account JSON key.
 * Share your spreadsheet with the service account's client_email.
 */
export function createSheetsClientFromServiceAccountJson(serviceAccountJson) {
  if (!serviceAccountJson?.client_email || !serviceAccountJson?.private_key) {
    throw new Error(
      "Invalid service account JSON. Expected client_email and private_key."
    );
  }

  const auth = new google.auth.JWT({
    email: serviceAccountJson.client_email,
    key: serviceAccountJson.private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  return google.sheets({ version: "v4", auth });
}

export async function readSheetValues({ sheets, spreadsheetId, rangeA1 }) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: rangeA1,
    valueRenderOption: "UNFORMATTED_VALUE",
  });
  return res.data.values ?? [];
}

export async function updateSheetValues({
  sheets,
  spreadsheetId,
  rangeA1,
  values,
}) {
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: rangeA1,
    valueInputOption: "RAW",
    requestBody: { values },
  });
}

