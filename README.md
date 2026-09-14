# academics-lo-monitor

Google Apps Script tooling for loading ZipGrade quiz exports into
per-grade-level template spreadsheets.

## Layout

- `library/` — the `ZipGradeLoader` Apps Script **library**: all the real
  logic (file picker, grade-level validation, xlsx conversion, matching
  ZipGrade rows into the template, color-coding responses).
- `template-bootstrap/` — the small script that gets bound to *each*
  grade-level template spreadsheet. It just wires up the `onOpen()` menu
  and forwards to the library, so every template stays in sync when the
  library is updated instead of carrying its own copy of the logic.

This split exists because Apps Script requires `onOpen()` (a simple
trigger) and anything called by `google.script.run` from the picker
dialog to live in the container-bound script itself — a library can't
provide those directly.

## One-time setup: publish the library

Requires [`clasp`](https://github.com/google/clasp) installed and logged
in (`clasp login`) with a Google account that has access to your Drive.

```bash
cd library
clasp create --type standalone --title "ZipGradeLoader Library" --rootDir .
clasp push
clasp open   # opens the new project in the Apps Script editor
```

`clasp create` creates the project at the root of My Drive — drag the
resulting file into your target folder
(<https://drive.google.com/drive/folders/1NmiYLxRmH4ZA5uMThjX8IigCM0XKv3O4>)
so it lives alongside the templates it serves.

In the Apps Script editor:

1. **Deploy → New deployment → select type "Library"** → Deploy.
2. Note the **Script ID** (Project Settings, ⚙️ icon) and the **version
   number** of the deployment you just created — both are needed below.

## Adding the library to a grade-level template

For each template spreadsheet (in **Extensions → Apps Script**):

1. In the editor, click **Libraries** (the `+` next to it) and paste the
   library's Script ID. Select the version you deployed, and set the
   **Identifier** to `ZipGradeLoader` (must match the name used in
   `template-bootstrap/Code.gs`).
2. Copy `template-bootstrap/Code.gs` and `template-bootstrap/appsscript.json`
   into that template's bound script project (paste manually, or `clasp
   clone <boundScriptId>` and copy the files in). If pasting the manifest
   directly, replace `REPLACE_WITH_LIBRARY_SCRIPT_ID` with the real
   library Script ID from above — adding the library through the editor
   UI in step 1 fills this in automatically, so this only matters if
   you're pushing the manifest via `clasp` instead.

Repeat for every grade-level template that should use the loader.

## Updating the logic later

1. Edit `library/Code.gs`, `clasp push`, then in the Apps Script editor
   create a **new** library deployment (Deploy → Manage deployments →
   Edit → New version, or Deploy → New deployment).
2. In each template's script project, open **Libraries** and bump the
   selected version to the new one. There's no "always latest" option for
   published libraries — each template pins a specific version and has to
   be updated explicitly.
