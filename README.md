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

If `LO Mapping` doesn't exist, this step is skipped entirely and the
rest of the load behaves exactly as before.
