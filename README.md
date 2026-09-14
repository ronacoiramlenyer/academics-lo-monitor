# academics-lo-monitor

Google Apps Script tooling for loading ZipGrade quiz exports into
grade-level template spreadsheets, tagging results against Learning
Outcomes/Competencies, and pushing the result to a shared Central
Database.

## Files

- `Code.gs` — the complete script: menu setup, the file picker (finds
  ZipGrade exports in Google Sheets or `.xlsx` format in the same Drive
  folder as the template), grade-level validation, matching/color-
  coding student responses into the template, LO/Competency tagging,
  and the Central Database push.
- `appsscript.json` — the project manifest. Enables the Drive API
  (advanced service), needed to convert an uploaded `.xlsx` file to a
  Google Sheet before it can be read.

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
4. Set `MASTER_SHEET_ID` near the top of `Code.gs` to the Central
   Database spreadsheet's ID (same value across every department's
   copy) — department itself is read straight from `DEPARTMENT_CELL`
   (default `A4`) on the active sheet, not set in code.
5. Save, then reload the spreadsheet. A "📊 ZipGrade Loader" menu and a
   "Learning Outcomes" menu should appear.

Repeat for each grade-level template. There's no shared library —
update each template's `Code.gs` directly if the script changes.

## LO/Competency tagging (optional)

Two sheets are involved:

- **`LOs-Competency`** — a pure reference sheet, declared once by
  whoever owns LO tracking. One row per LO, with up to five
  `Competency N` columns holding each competency's description:

  | LO Code | LO Description | Competency 1 | Competency 2 | ... |
  |---|---|---|---|---|
  | Comp. G10.1 | ... | Clean raw data in... | Organize data structures by... | |

  The loader never writes to this sheet — it only reads it.

- **`GRADE #`** (e.g. `GRADE 7`, `GRADE 10`) — synced by the
  **"Refresh Competency Summary"** menu action (separate from "Load
  ZipGrade Data", so updating `Item Placement` never requires
  re-loading the quiz file). It unrolls every non-blank `Competency N`
  cell from `LOs-Competency` into its own row (`LO Code`,
  `LO Description`, `Competency`), so nothing needs retyping there. The
  **only** manual step is filling in `Item Placement` per row — the
  ZipGrade question numbers that competency covers, separated by commas
  (e.g. `1, 3, 4, 5`) — spaces alone also work. Re-syncing preserves
  whatever's already typed there.

  **Rows above row 12 on this sheet are never read from or written to**
  — that's reserved for whatever title/header formatting you've already
  built there. The script locates its columns (`LO Code`, `Competency`,
  section letters, `TOTAL`, ...) by scanning that header area for
  matching label text, then only ever writes into row 12 and below —
  updating an existing row's computed cells in place, or appending a
  new row at the bottom for a (LO, Competency) pair it hasn't seen
  before. If a label or section can't be found in the header, that
  column is just skipped rather than guessed.

  Each section's cell is the sum, across every student matched into
  that section, of how many of that row's items they answered
  correctly.

If any question number isn't covered by any row's `Item Placement`,
the refresh's completion popup (and the log) calls it out so it can be
fixed.

**Workflow**: a teacher runs "Load ZipGrade Data" as usual — nothing
changes for them. Whoever owns LO tracking sets up `LOs-Competency`
once, fills in `Item Placement` on `GRADE #` as needed, and runs
"Refresh Competency Summary" whenever they want the totals recomputed
from whatever's currently in the section sheets. If `LOs-Competency`
doesn't exist yet, "Refresh Competency Summary" just says so instead of
doing anything.

## Central Database push (optional)

The "Learning Outcomes" menu pushes the active `GRADE #` sheet's rows
up to one shared Master spreadsheet across departments:

- **Push to Central Database** — extracts, shows an update/new count,
  confirms, then writes.
- **Preview extraction (no push)** — runs the same extraction and logs
  it, without touching the Master sheet.

Both read every field straight from the `GRADE #` sheet's own columns
(`Item`, `MaxScore`, `Competency`, `AssessmentType`, `ItemPerformance`)
rather than deriving them, since those can already hold whatever the
sheet owner put there (a rubric label, a formula result, a slot name
like `Competency1`, a ZipGrade item number) — the columns are located
dynamically by header text, same as the rest of this script, never
assumed to be at fixed positions. `StudentCount` comes from a footer
row matching `FOOTER_LABEL_PATTERN` (e.g. "TOTAL NO. OF STUDENTS").

Department is read straight from `DEPARTMENT_CELL` (default `A4`) on
the active sheet — a merged cell reading e.g. "FILIPINO Department";
the trailing " Department" is stripped off automatically. Nothing
needs setting in code for this.

Constants to set near the top of `Code.gs`:
- `MASTER_SHEET_ID` — the Central Database spreadsheet's ID (ships as
  `'CHANGE_ME'`; same value across every department's copy).
- `HEADER_INFO_CELL` (default `'A5'`) — where "Third Trimester,
  SY 2025-2026"-style text lives in the `GRADE #` sheet's own title
  rows — adjust if it's elsewhere.
- `DEPARTMENT_CELL` (default `'A4'`) — where the department name lives
  — adjust if it's elsewhere.
