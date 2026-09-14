# academics-lo-monitor

Google Apps Script tool for loading ZipGrade quiz exports into
grade-level template spreadsheets.

## Files

- `Code.gs` — the complete script: menu setup, the file picker (finds
  ZipGrade exports in Google Sheets or `.xlsx` format in the same Drive
  folder as the template), grade-level validation, and matching/color-
  coding student responses into the template.
- `appsscript.json` — the project manifest. Enables the Drive API
  (advanced service), which is needed to convert an uploaded `.xlsx`
  file to a Google Sheet before it can be read.

## Setup

Each grade-level template spreadsheet gets its own copy of this script,
bound directly to it:

1. Open the template spreadsheet → **Extensions → Apps Script**.
2. Paste the contents of `Code.gs` into the editor (replacing the default
   `Code.gs` there).
3. In **Project Settings** (⚙️), enable "Show `appsscript.json` manifest
   file in editor", then paste the contents of this repo's
   `appsscript.json` into it — this turns on the Drive API service the
   script needs for `.xlsx` support.
4. Save, then reload the spreadsheet. A "📊 ZipGrade Loader" menu should
   appear.

Repeat for each grade-level template. There's no shared library — update
each template's `Code.gs` directly if the script changes.
