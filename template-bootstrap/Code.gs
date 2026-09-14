/**
 * ZIPGRADE DATA LOADER - TEMPLATE BOOTSTRAP
 *
 * This is the small script bound to each grade-level template
 * spreadsheet. All of the real logic lives in the shared
 * "LO_library_code_academics" Apps Script library (see /library in
 * this repo), auto-loaded via appsscript.json, so every template
 * stays in sync when the library is updated. This file only needs
 * the pieces Apps Script requires to live in the container-bound
 * script itself:
 *   - onOpen(), a simple trigger, which libraries cannot provide
 *   - functions referenced by menu items and by google.script.run,
 *     which must be top-level functions in the bound script
 *
 * See README.md for how to wire up the library Script ID.
 */

// Set once per copy of this file - the department this spreadsheet
// pushes to the Central Database as.
const DEPARTMENT = "CHANGE_ME";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 ZipGrade Loader")
    .addItem("Load ZipGrade Data", "showFileList")
    .addItem("Refresh Competency Summary", "refreshCompetencySummary")
    .addItem("Show Instructions", "showInstructions")
    .addSeparator()
    .addItem("View Logs", "viewLogs")
    .addToUi();

  ui.createMenu("Learning Outcomes")
    .addItem("Push to Central Database", "pushToCentralDatabase")
    .addItem("Preview extraction (no push)", "previewExtraction")
    .addToUi();
}

function pushToCentralDatabase() {
  LO_library_code_academics.push(DEPARTMENT);
}

function previewExtraction() {
  LO_library_code_academics.preview();
}

function showFileList() {
  LO_library_code_academics.showFileList();
}

function processSelectedFile(fileId) {
  LO_library_code_academics.processSelectedFile(fileId);
}

function refreshCompetencySummary() {
  LO_library_code_academics.refreshCompetencySummary();
}

function showInstructions() {
  LO_library_code_academics.showInstructions();
}

function viewLogs() {
  LO_library_code_academics.viewLogs();
}
