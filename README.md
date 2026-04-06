# joblink-checkr

Node.js CLI that reads job links from a Google Sheet, checks each URL, and writes the results back to the sheet.

## What it does

- Reads a sheet range (first row must be headers)
- Finds the URL column (default header name: `URL`)
- Checks each URL (follows redirects, timeout configurable)
- Fetches page text and scans for configured "bad keywords" (optional)
- Writes these columns (created if missing):
  - `Approved`, `Reason`

## Prereqs

- Node.js **18+**
- A Google Sheet with a header row that includes a column named `URL` (or set a different name in config)

## Google Sheets setup (Service Account)

1. In Google Cloud Console, create a project (or use an existing one).
2. Enable **Google Sheets API** for the project.
3. Create a **Service Account** and download its JSON key.
4. Open your Google Sheet and **share it** with the service account email (`client_email` in the JSON), giving **Editor** permission.

## Configure

This repo avoids committing `.env` files. Use a JSON config instead:

1. Copy `config.example.json` to `config.json`
2. Fill in:
   - `spreadsheets: [...]` if you have multiple **separate spreadsheets** (recommended), and list them in the order you want them processed
   - `badKeywords` (optional): if any of these phrases appear in the job description page text, `Approved` becomes `No`
   - `badTitleKeywords` (optional): if any of these phrases appear in the position title, the URL check is skipped and `Approved` becomes `No`
   - `serviceAccount` (paste the full service account JSON object)
   - For each sheet entry configure:
     - `sheetName`, `startCol`, `endCol`
     - `headerRow1`, `dataStartRow1`, `dataEndRow1`

### Row/column config (no A1 ranges)

Set:

- `sheetName`: the tab name in Google Sheets (e.g. `Jobs` or `Sheet1`)
- `startCol` / `endCol`: column letters (e.g. `A` and `Z`)
- `headerRow1`: header row number (1-based)
- `dataStartRow1`: first data row number (1-based, inclusive)
- `dataEndRow1`: last data row number (1-based, inclusive)

## Install & run

```bash
npm install
```

Dry run (prints a sample of what would be written):

```bash
# Windows PowerShell
$env:CONFIG_FILE="E:\joblink-checkr\config.json"
npm run check:sheet:dry
```

Real update (writes back to the sheet):

```bash
# Windows PowerShell
$env:CONFIG_FILE="E:\joblink-checkr\config.json"
npm run check:sheet
```

## Expected sheet layout

Example headers:

- `URL`
- (any other columns you want)

The script will append result columns if they don’t exist yet.

