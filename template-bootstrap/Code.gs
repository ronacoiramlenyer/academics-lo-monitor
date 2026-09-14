/**
 * ZIPGRADE DATA LOADER - TEMPLATE BOOTSTRAP
 *
 * This is the small script bound to each grade-level template
 * spreadsheet. All of the real logic lives in the shared
 * "ZipGradeLoader" Apps Script library (see /library in this repo),
 * so every template stays in sync when the library is updated —
 * this file only needs the pieces that Apps Script requires to live
 * in the container-bound script itself:
 *   - onOpen(), a simple trigger, which libraries cannot provide
 *   - functions referenced by menu items and by google.script.run,
 *     which must be top-level functions in the bound script
 *
 * See README.md for how to add the ZipGradeLoader library to a
 * template's bound script.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 ZipGrade Loader")
    .addItem("Load ZipGrade Data", "showFileList")
    .addItem("Show Instructions", "showInstructions")
    .addSeparator()
    .addItem("View Logs", "viewLogs")
    .addToUi();
}

function showFileList() {
  ZipGradeLoader.showFileList();
}

function processSelectedFile(fileId) {
  ZipGradeLoader.processSelectedFile(fileId);
}

function showInstructions() {
  ZipGradeLoader.showInstructions();
}

function viewLogs() {
  ZipGradeLoader.viewLogs();
}
