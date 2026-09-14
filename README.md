# academics-lo-monitor

Google Apps Script tooling for loading ZipGrade quiz exports into
per-grade-level template spreadsheets.

## Layout

- `library/` — the `LO_library_code_academics` Apps Script **library**
  (project title: `LO-library-code-academics`): all the real logic (file
  picker, grade-level validation, `.xlsx` conversion, matching ZipGrade
  rows into the template, color-coding responses).
- `template-bootstrap/` — the small script bound to *each* grade-level
  template spreadsheet. It just wires up the `onOpen()` menu and forwards
  to the library, so every template stays in sync when the library is
  updated instead of carrying its own copy of the logic. Its
  `appsscript.json` auto-loads the library — no manual "Add a library"
  step needed in the Apps Script UI once the Script ID below is filled
  in.

This split exists because Apps Script requires `onOpen()` (a simple
trigger) and anything called by `google.script.run` from the picker
dialog to live in the container-bound script itself — a library can't
provide those directly.

## One-time setup: publish the library

1. Create a standalone Apps Script project (via `clasp` or
   script.google.com), push `library/Code.gs` and `library/appsscript.json`
   into it.
2. **Deploy → New deployment → type "Library"** → Deploy.
3. Note the project's **Script ID** (Project Settings ⚙️) and the
   **version number** of the deployment.
4. In `template-bootstrap/appsscript.json`, replace
   `REPLACE_WITH_LIBRARY_SCRIPT_ID` with that Script ID (and bump
   `"version"` here whenever you deploy a new library version).

## Adding a grade-level template

For each template spreadsheet, in **Extensions → Apps Script**:

1. Paste `template-bootstrap/Code.gs` and `template-bootstrap/appsscript.json`
   (with the real library Script ID filled in) into that template's bound
   script project.
2. Turn on "Show `appsscript.json` manifest file in editor" under Project
   Settings ⚙️ if the manifest isn't visible yet, so you can paste the
   JSON in directly.
3. Reload the spreadsheet — the "📊 ZipGrade Loader" menu should appear,
   backed by the library. No manual Libraries-dialog step is needed since
   the manifest already declares the dependency.

## Updating the logic later

1. Edit `library/Code.gs`, push, then create a **new** library deployment
   (Deploy → Manage deployments → Edit → New version).
2. Bump `"version"` in every template's `appsscript.json` to the new
   number. There's no "always latest" option for published libraries —
   each template pins a specific version.
