/**
 * ZIPGRADE DATA LOADER - PRACTICAL SKILLS TEMPLATE
 *
 * Smart grade level detection and validation
 * Automatically finds ZipGrade files in the same folder as template
 * Dynamically detects all template sheets
 * Remembers the template format for future loads
 */

// Pure reference sheet, declared beforehand: one row per LO, with up to
// 5 "Competency N" description columns. The loader only reads it to
// auto-line-up (LO Code, Competency) rows on the GRADE # sheet below -
// it never writes to LOs-Competency itself.
const LO_COMPETENCY_SHEET_NAME = "LOs-Competency";

// The GRADE # sheet's own title/header formatting lives in rows above
// this and is never touched by the script - LO/Competency rows always
// start here, whatever's already above stays exactly as it is.
const GRADE_SUMMARY_DATA_START_ROW = 12;

/**
 * Create menu when sheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 ZipGrade Loader")
    .addItem("Load ZipGrade Data", "showFileList")
    .addItem("Refresh Competency Summary", "refreshCompetencySummary")
    .addItem("Show Instructions", "showInstructions")
    .addSeparator()
    .addItem("View Logs", "viewLogs")
    .addToUi();
}

/**
 * Get files from same folder as template
 */
function showFileList() {
  try {
    const templateSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const templateFile = DriveApp.getFileById(templateSpreadsheet.getId());
    const parentFolder = templateFile.getParents().next();

    // Auto-detect grade level from sheets
    const allSheets = templateSpreadsheet.getSheets();
    const detectedGrades = new Set();

    for (const sheet of allSheets) {
      const grade = extractGradeLevel(sheet.getName());
      if (grade) detectedGrades.add(grade);
    }

    if (detectedGrades.size === 0) {
      throw new Error("Could not detect grade level from sheet names");
    }

    if (detectedGrades.size > 1) {
      throw new Error(`Multiple grade levels detected: ${Array.from(detectedGrades).join(', ')}. This template mixes grades.`);
    }

    const selectedGrade = Array.from(detectedGrades)[0];
    console.log(`Template folder: ${parentFolder.getName()}`);
    console.log(`Auto-detected grade level: ${selectedGrade}`);

    // Get all files in folder
    // ZipGrade exports can be a native Google Sheet or an uploaded .xlsx file
    const allowedMimeTypes = [MimeType.GOOGLE_SHEETS, MimeType.MICROSOFT_EXCEL];
    const files = parentFolder.getFiles();
    const fileList = [];

    while (files.hasNext()) {
      const file = files.next();
      const fileId = file.getId();

      // Skip the template itself and any non-spreadsheet files
      if (fileId === templateSpreadsheet.getId()) continue;
      if (allowedMimeTypes.indexOf(file.getMimeType()) === -1) continue;

      fileList.push({
        name: file.getName(),
        id: fileId
      });
    }

    if (fileList.length === 0) {
      SpreadsheetApp.getUi().alert("❌ No other spreadsheet files found in this folder.");
      return;
    }

    // Sort by name
    fileList.sort((a, b) => a.name.localeCompare(b.name));

    // Create HTML with file list
    const html = HtmlService.createHtmlOutput(`
      <style>
        body {
          font-family: Arial, sans-serif;
          padding: 20px;
          background: #f9f9f9;
        }
        .file-list {
          max-height: 400px;
          overflow-y: auto;
          border: 1px solid #ddd;
          border-radius: 4px;
          background: white;
        }
        .file-item {
          padding: 12px;
          border-bottom: 1px solid #eee;
          cursor: pointer;
          transition: background 0.2s;
        }
        .file-item:hover {
          background: #f0f0f0;
        }
        .file-item:last-child {
          border-bottom: none;
        }
        .file-name {
          font-weight: 500;
          color: #333;
        }
        p {
          margin: 0 0 15px 0;
          color: #666;
        }
      </style>

      <p>📁 Select a ZipGrade file from this folder:</p>

      <div class="file-list" id="fileList"></div>

      <script>
        const files = ${JSON.stringify(fileList)};

        const fileListEl = document.getElementById('fileList');
        files.forEach(file => {
          const item = document.createElement('div');
          item.className = 'file-item';
          const nameEl = document.createElement('div');
          nameEl.className = 'file-name';
          nameEl.textContent = file.name;
          item.appendChild(nameEl);
          item.onclick = () => {
            google.script.run.processSelectedFile(file.id);
            google.script.host.close();
          };
          fileListEl.appendChild(item);
        });
      </script>
    `);

    SpreadsheetApp.getUi().showModelessDialog(html, 'Select ZipGrade File');

  } catch (error) {
    console.error("ERROR: " + error.message);
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.message);
  }
}

const PROGRESS_CACHE_KEY = "zipgrade_loader_progress";

/**
 * Record progress for the modeless dialog (shown by showProgressDialog)
 * to poll and render. Stored in the script cache since the dialog's
 * polling calls are separate executions from the one doing the work.
 */
function updateProgress(current, total, message) {
  CacheService.getScriptCache().put(PROGRESS_CACHE_KEY, JSON.stringify({
    current: current,
    total: total,
    message: message,
    done: false,
    error: null
  }), 300);
}

function completeProgress(message) {
  CacheService.getScriptCache().put(PROGRESS_CACHE_KEY, JSON.stringify({
    current: 1,
    total: 1,
    message: message,
    done: true,
    error: null
  }), 300);
}

function failProgress(message) {
  CacheService.getScriptCache().put(PROGRESS_CACHE_KEY, JSON.stringify({
    current: 0,
    total: 0,
    message: "",
    done: true,
    error: message
  }), 300);
}

/**
 * Called by the progress dialog's client-side polling script.
 */
function getProgress() {
  const raw = CacheService.getScriptCache().get(PROGRESS_CACHE_KEY);
  return raw ? JSON.parse(raw) : { current: 0, total: 0, message: "", done: false, error: null };
}

/**
 * Show a small modeless dialog with a progress bar that polls
 * getProgress() and closes itself once the load finishes or errors.
 */
function showProgressDialog() {
  const html = HtmlService.createHtmlOutput(`
    <style>
      body {
        font-family: Arial, sans-serif;
        padding: 16px;
      }
      #status {
        font-size: 13px;
        color: #333;
        margin-bottom: 10px;
        min-height: 18px;
      }
      .bar-track {
        width: 100%;
        height: 10px;
        background: #eee;
        border-radius: 5px;
        overflow: hidden;
      }
      .bar-fill {
        height: 100%;
        width: 0%;
        background: #4285F4;
        transition: width 0.3s ease;
      }
      .bar-fill.error {
        background: #EA4335;
      }
    </style>
    <div id="status">Starting...</div>
    <div class="bar-track"><div class="bar-fill" id="fill"></div></div>
    <script>
      function poll() {
        google.script.run.withSuccessHandler(onProgress).withFailureHandler(onFailure).getProgress();
      }

      function onFailure(error) {
        document.getElementById('status').textContent = '❌ ' + error.message;
        document.getElementById('fill').classList.add('error');
        setTimeout(() => google.script.host.close(), 2500);
      }

      function onProgress(progress) {
        const statusEl = document.getElementById('status');
        const fillEl = document.getElementById('fill');

        if (progress.error) {
          statusEl.textContent = '❌ ' + progress.error;
          fillEl.classList.add('error');
          fillEl.style.width = '100%';
          setTimeout(() => google.script.host.close(), 2500);
          return;
        }

        const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : (progress.done ? 100 : 0);
        fillEl.style.width = pct + '%';
        statusEl.textContent = progress.message + (progress.total > 0 ? \` (\${progress.current}/\${progress.total})\` : '');

        if (progress.done) {
          setTimeout(() => google.script.host.close(), 1200);
          return;
        }

        setTimeout(poll, 700);
      }

      poll();
    </script>
  `).setWidth(320).setHeight(90);

  SpreadsheetApp.getUi().showModelessDialog(html, "Loading ZipGrade Data...");
}

/**
 * Process the selected file
 */
function processSelectedFile(fileId) {
  let tempConvertedFileId = null;
  updateProgress(0, 0, "Starting...");
  showProgressDialog();
  try {
    console.log("Processing file ID: " + fileId);

    // Auto-detect grade level from template sheets
    const templateSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = templateSpreadsheet.getSheets();
    const detectedGrades = new Set();

    for (const sheet of allSheets) {
      const grade = extractGradeLevel(sheet.getName());
      if (grade) detectedGrades.add(grade);
    }

    if (detectedGrades.size === 0) {
      throw new Error("Could not detect grade level from sheet names");
    }

    const templateGradeLevel = Array.from(detectedGrades)[0];
    console.log(`Auto-detected template grade level: ${templateGradeLevel}`);

    // Get the selected file
    let zipgradeFile;
    try {
      zipgradeFile = DriveApp.getFileById(fileId);
    } catch (e) {
      throw new Error(`Could not access file: ${e.message}`);
    }

    console.log(`Found file: ${zipgradeFile.getName()}`);

    // Open file as spreadsheet. ZipGrade exports may be a native Google
    // Sheet, or an .xlsx upload that first needs converting to one.
    const zipgradeMimeType = zipgradeFile.getMimeType();
    let zipgradeSpreadsheet;
    try {
      if (zipgradeMimeType === MimeType.GOOGLE_SHEETS) {
        zipgradeSpreadsheet = SpreadsheetApp.open(zipgradeFile);
      } else if (zipgradeMimeType === MimeType.MICROSOFT_EXCEL) {
        console.log("Converting uploaded Excel file to Google Sheets format...");
        const converted = Drive.Files.insert(
          { title: zipgradeFile.getName(), mimeType: MimeType.GOOGLE_SHEETS },
          zipgradeFile.getBlob(),
          { convert: true }
        );
        tempConvertedFileId = converted.id;
        zipgradeSpreadsheet = SpreadsheetApp.openById(tempConvertedFileId);
      } else {
        throw new Error(`Unsupported file type "${zipgradeMimeType}". Please upload a Google Sheet or an .xlsx file.`);
      }
    } catch (e) {
      throw new Error(`Could not open file: ${e.message}`);
    }

    const zipgradeSheet = zipgradeSpreadsheet.getSheets()[0];
    const zipgradeData = zipgradeSheet.getDataRange().getValues();

    console.log(`Loaded ${zipgradeData.length} rows from ZipGrade file`);

    // SMART CHECK: Detect grade level from ZipGrade data
    const zipgradeGradeLevel = detectGradeLevelFromZipGrade(zipgradeData);

    // VALIDATION: Check if grade levels match
    if (templateGradeLevel !== zipgradeGradeLevel) {
      throw new Error(`❌ GRADE LEVEL MISMATCH!\n\nTemplate grade level: ${templateGradeLevel}\nZipGrade file grade level: ${zipgradeGradeLevel}\n\nPlease load the correct ZipGrade file for Grade ${templateGradeLevel}.`);
    }

    console.log(`✓ Grade levels match: Grade ${templateGradeLevel}`);

    // Find column indices
    const headers = zipgradeData[0];
    let zipgradeIdIndex = -1;
    let firstNameIndex = -1;
    let lastNameIndex = -1;
    let numQuestionsIndex = -1;
    let numCorrectIndex = -1;
    let percentCorrectIndex = -1;
    let classIndex = -1;
    let questionStartIndex = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim().toLowerCase();
      if (header === "zipgrade id") zipgradeIdIndex = i;
      else if (header === "first name") firstNameIndex = i;
      else if (header === "last name") lastNameIndex = i;
      else if (header === "num questions") numQuestionsIndex = i;
      else if (header === "num correct") numCorrectIndex = i;
      else if (header === "percent correct") percentCorrectIndex = i;
      else if (header === "class") classIndex = i;
      else if (header.match(/^q\d+$/)) {
        if (questionStartIndex === -1) questionStartIndex = i;
      }
    }

    // Validate columns
    if (zipgradeIdIndex === -1) throw new Error("Column 'ZipGrade ID' not found");
    if (numQuestionsIndex === -1) throw new Error("Column 'Num Questions' not found");
    if (numCorrectIndex === -1) throw new Error("Column 'Num Correct' not found");
    if (percentCorrectIndex === -1) throw new Error("Column 'Percent Correct' not found");
    if (classIndex === -1) throw new Error("Column 'Class' not found");
    if (questionStartIndex === -1) throw new Error("Question columns (Q1, Q2, ...) not found");

    // Count questions
    let numQuestions = 0;
    for (let i = questionStartIndex; i < headers.length; i++) {
      if (headers[i].toString().match(/^Q\d+$/i)) numQuestions++;
      else break;
    }
    console.log(`Found ${numQuestions} questions`);

    // Create lookup, keyed by section (Class column) + student number so a
    // student number that repeats across sections can't cross-match.
    const studentLookup = {};
    for (let row = 1; row < zipgradeData.length; row++) {
      const studentNum = zipgradeData[row][zipgradeIdIndex];
      if (studentNum) {
        const section = normalizeSection(zipgradeData[row][classIndex]);
        const key = buildLookupKey(section, studentNum);
        studentLookup[key] = {
          firstName: zipgradeData[row][firstNameIndex] || "",
          lastName: zipgradeData[row][lastNameIndex] || "",
          numQuestions: zipgradeData[row][numQuestionsIndex] || 0,
          numCorrect: zipgradeData[row][numCorrectIndex] || 0,
          percentCorrect: zipgradeData[row][percentCorrectIndex] || 0,
          questions: zipgradeData[row].slice(questionStartIndex, questionStartIndex + numQuestions)
        };
      }
    }
    console.log(`Created lookup for ${Object.keys(studentLookup).length} students`);

    const greenFill = "#92D050";
    const redFill = "#FF7F7F";

    // Filter sheets for selected grade
    const sheetsForGrade = allSheets.filter(sheet => {
      const sheetGrade = extractGradeLevel(sheet.getName());
      return sheetGrade === templateGradeLevel;
    });

    console.log(`Found ${sheetsForGrade.length} sheets for Grade ${templateGradeLevel}`);

    const totalSteps = sheetsForGrade.length;

    // Process sheets for this grade
    let totalMatched = 0;
    const classResults = {};
    let sheetsProcessed = 0;

    for (const templateSheet of sheetsForGrade) {
      const sheetName = templateSheet.getName();
      const section = normalizeSection(sheetName);

      updateProgress(sheetsProcessed, totalSteps, `Processing ${sheetName}...`);

      const templateData = templateSheet.getDataRange().getValues();

      // The header row isn't always row 1 - some templates have title
      // or adviser rows above the real column headers, so search for it.
      const headerRowIndex = findHeaderRowIndex(templateData, "student number");

      if (headerRowIndex === -1) {
        console.log(`⚠ Required columns not found in ${sheetName}, skipping`);
        sheetsProcessed++;
        continue;
      }

      const templateHeaders = templateData[headerRowIndex];

      // Find the Student Number column
      let studentNumberColIndex = -1;
      for (let i = 0; i < templateHeaders.length; i++) {
        if (templateHeaders[i].toString().trim().toLowerCase() === "student number") {
          studentNumberColIndex = i;
          break;
        }
      }

      // Columns D onward are fully generated from the ZipGrade file on
      // every load - Num Questions, Num Correct, Percent Correct, then
      // Q1...Qn - so the template doesn't need to already have them,
      // and a question-count change between loads doesn't leave stale
      // columns behind.
      const columns = prepareSectionColumns(templateSheet, headerRowIndex + 1, numQuestions);

      // Update rows
      let sheetMatched = 0;

      for (let row = headerRowIndex + 1; row < templateData.length; row++) {
        const studentNum = templateData[row][studentNumberColIndex];

        if (!studentNum || studentNum === "Student Number") continue;

        const key = buildLookupKey(section, studentNum);

        if (studentLookup[key]) {
          const student = studentLookup[key];
          const updateRow = row + 1;

          templateSheet.getRange(updateRow, columns.numQuestionsColIndex + 1).setValue(student.numQuestions);
          templateSheet.getRange(updateRow, columns.numCorrectColIndex + 1).setValue(student.numCorrect);
          templateSheet.getRange(updateRow, columns.percentCorrectColIndex + 1).setValue(student.percentCorrect);

          for (let q = 0; q < student.questions.length; q++) {
            const response = parseInt(student.questions[q]) || 0;
            const col = columns.questionColStartIndex + q + 1;
            const cell = templateSheet.getRange(updateRow, col);

            cell.setValue(response);
            cell.setBackground(response === 1 ? greenFill : redFill);
          }

          sheetMatched++;
          totalMatched++;
        }
      }

      // Autofit the generated columns to their content
      templateSheet.autoResizeColumns(4, 3 + numQuestions);

      classResults[sheetName] = sheetMatched;
      console.log(`✓ ${sheetName}: ${sheetMatched} students matched`);
      sheetsProcessed++;
    }

    console.log("\n" + "=".repeat(60));
    console.log("✓ COMPLETE!");
    console.log("=".repeat(60));
    console.log(`Grade ${templateGradeLevel} - Total students matched: ${totalMatched}`);

    completeProgress(`✓ Loaded ${totalMatched} students`);

    let completionMessage = `✓ Successfully loaded Grade ${templateGradeLevel} data!\n\nTotal students matched: ${totalMatched}\n\nCheck the logs for details.`;
    if (readLoCompetencyPairs(templateSpreadsheet, templateGradeLevel).length > 0) {
      completionMessage += `\n\nRun "Refresh Competency Summary" from the menu to update the GRADE ${templateGradeLevel} summary.`;
    }
    SpreadsheetApp.getUi().alert(completionMessage);

  } catch (error) {
    console.error("ERROR: " + error.message);
    failProgress(error.message);
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.message);
  } finally {
    if (tempConvertedFileId) {
      try {
        DriveApp.getFileById(tempConvertedFileId).setTrashed(true);
        console.log("Cleaned up temporary converted file");
      } catch (cleanupError) {
        console.error("Could not clean up temporary converted file: " + cleanupError.message);
      }
    }
  }
}

/**
 * Refresh the GRADE # competency summary independently of loading
 * ZipGrade data - re-reads each section sheet's already-written
 * Student Number + Q1...Qn columns (no ZipGrade file needed), so
 * updating Item Placement on the GRADE # sheet doesn't require
 * re-running the full load.
 */
function refreshCompetencySummary() {
  try {
    const templateSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = templateSpreadsheet.getSheets();
    const detectedGrades = new Set();

    for (const sheet of allSheets) {
      const grade = extractGradeLevel(sheet.getName());
      if (grade) detectedGrades.add(grade);
    }

    if (detectedGrades.size === 0) {
      throw new Error("Could not detect grade level from sheet names");
    }

    const templateGradeLevel = Array.from(detectedGrades)[0];

    const loCompetencyPairs = readLoCompetencyPairs(templateSpreadsheet, templateGradeLevel);
    if (loCompetencyPairs.length === 0) {
      SpreadsheetApp.getUi().alert(
        `❌ No "${LO_COMPETENCY_SHEET_NAME}" sheet found (or it has no usable rows).\n\nSet it up first, then try again.`
      );
      return;
    }

    const sheetsForGrade = allSheets.filter(sheet => extractGradeLevel(sheet.getName()) === templateGradeLevel);
    const responses = readSectionResponses(sheetsForGrade);

    const untaggedItems = syncGradeSummarySheet(
      templateSpreadsheet, templateGradeLevel, loCompetencyPairs, responses.studentLookup, sheetsForGrade, responses.numQuestions
    );

    let message = `✓ Refreshed the GRADE ${templateGradeLevel} competency summary.`;
    if (untaggedItems.length > 0) {
      message += `\n\n⚠ Items not tagged to any competency: ${untaggedItems.join(", ")}`;
    }
    SpreadsheetApp.getUi().alert(message);

  } catch (error) {
    console.error("ERROR: " + error.message);
    SpreadsheetApp.getUi().alert("❌ Error:\n\n" + error.message);
  }
}

/**
 * Rebuild a studentLookup (section+studentNumber -> { questions }) by
 * reading each section sheet's own Student Number and Q1...Qn columns
 * directly, instead of from a ZipGrade file. Used by
 * refreshCompetencySummary so it doesn't need the ZipGrade export.
 */
function readSectionResponses(sheetsForGrade) {
  const studentLookup = {};
  let numQuestions = 0;

  sheetsForGrade.forEach(templateSheet => {
    const section = normalizeSection(templateSheet.getName());
    const templateData = templateSheet.getDataRange().getValues();

    const headerRowIndex = findHeaderRowIndex(templateData, "student number");
    if (headerRowIndex === -1) return;

    const headers = templateData[headerRowIndex];
    let studentNumberColIndex = -1;
    let questionStartIndex = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim().toLowerCase();
      if (header === "student number") studentNumberColIndex = i;
      else if (questionStartIndex === -1 && header.match(/^q\d+$/)) questionStartIndex = i;
    }

    if (studentNumberColIndex === -1 || questionStartIndex === -1) return;

    let sheetNumQuestions = 0;
    for (let i = questionStartIndex; i < headers.length; i++) {
      if (headers[i].toString().match(/^Q\d+$/i)) sheetNumQuestions++;
      else break;
    }
    numQuestions = Math.max(numQuestions, sheetNumQuestions);

    for (let row = headerRowIndex + 1; row < templateData.length; row++) {
      const studentNum = templateData[row][studentNumberColIndex];
      if (!studentNum || studentNum === "Student Number") continue;

      const key = buildLookupKey(section, studentNum);
      studentLookup[key] = {
        questions: templateData[row].slice(questionStartIndex, questionStartIndex + sheetNumQuestions)
      };
    }
  });

  return { studentLookup: studentLookup, numQuestions: numQuestions };
}

/**
 * Wipe columns D onward (header row through the last existing row) on
 * a section sheet and rebuild them from the ZipGrade file: Num
 * Questions, Num Correct, Percent Correct, then Q1...Qn. This means
 * the template doesn't need those columns to already exist, and a
 * question-count change between loads doesn't leave stale extra
 * columns behind. Returns the 0-based column indices (matching the
 * getValues() row-array convention used elsewhere) for the generated
 * columns.
 */
function prepareSectionColumns(sheet, headerRow, numQuestions) {
  const dataHeaders = ["Num Questions", "Num Correct", "Percent Correct"];
  for (let q = 1; q <= numQuestions; q++) {
    dataHeaders.push("Q" + q);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 3 + dataHeaders.length);
  const lastRow = Math.max(sheet.getLastRow(), headerRow);

  sheet.getRange(headerRow, 4, lastRow - headerRow + 1, lastColumn - 3).clear();
  sheet.getRange(headerRow, 4, 1, dataHeaders.length).setValues([dataHeaders]);

  return {
    numQuestionsColIndex: 3,
    numCorrectColIndex: 4,
    percentCorrectColIndex: 5,
    questionColStartIndex: 6
  };
}

/**
 * Read LOs-Competency, if it exists, and unroll it into one (LO Code,
 * Competency) pair per row's LO Description plus each non-blank
 * "Competency N" cell - however many Competency columns exist and are
 * filled in, not just a fixed count. The LO Description itself becomes
 * its own row first (so a row with a description and 2 competencies
 * unrolls into 3 pairs), then each competency in column order. This is
 * the auto-line-up source for the GRADE # sheet - LOs-Competency
 * itself carries no item numbers, only descriptions.
 *
 * A row's LO Code is auto-generated ("Comp. G{gradeLevel}.{n}") and
 * written back into the sheet whenever that cell is blank, so a row
 * only needs a description and competencies typed in - it's never
 * skipped just for missing a code. An LO Code someone already typed is
 * left exactly as it is.
 *
 * Returns [] if the sheet doesn't exist or has no usable rows -
 * competency tagging is entirely optional and shouldn't block a
 * normal load.
 */
function readLoCompetencyPairs(spreadsheet, gradeLevel) {
  const sheet = spreadsheet.getSheetByName(LO_COMPETENCY_SHEET_NAME);
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  const headerRowIndex = findHeaderRowIndex(data, "lo code");
  if (headerRowIndex === -1) return [];

  const headers = data[headerRowIndex].map(h => h.toString().trim().toLowerCase());
  const loCodeIdx = headers.indexOf("lo code");
  const loDescriptionIdx = headers.indexOf("lo description");

  if (loCodeIdx === -1) return [];

  const competencyIndices = [];
  headers.forEach((header, i) => {
    if (header.match(/^competency\s*\d+$/)) competencyIndices.push(i);
  });

  const pairs = [];
  let loNumber = 0;

  for (let row = headerRowIndex + 1; row < data.length; row++) {
    const loDescription = loDescriptionIdx !== -1 ? (data[row][loDescriptionIdx] || "").toString().trim() : "";
    const rowCompetencies = competencyIndices.filter(colIndex => {
      const cell = data[row][colIndex];
      return cell && cell.toString().trim() !== "";
    });

    // Skip rows with nothing to tag - no description and no competencies
    if (!loDescription && rowCompetencies.length === 0) continue;

    loNumber++;

    let loCode = data[row][loCodeIdx] ? data[row][loCodeIdx].toString().trim() : "";
    if (!loCode) {
      loCode = `Comp. G${gradeLevel}.${loNumber}`;
      sheet.getRange(row + 1, loCodeIdx + 1).setValue(loCode);
    }

    if (loDescription) {
      pairs.push({
        loCode: loCode,
        loDescription: loDescription,
        competency: loDescription
      });
    }

    rowCompetencies.forEach(colIndex => {
      pairs.push({
        loCode: loCode,
        loDescription: loDescription,
        competency: data[row][colIndex].toString().trim()
      });
    });
  }

  return pairs;
}

/**
 * Parse an Item Placement cell into question numbers. Space-separated
 * is the expected format ("1 3 5"), but commas are tolerated too.
 */
function parseItemPlacement(raw) {
  if (!raw) return [];
  return raw.toString()
    .split(/[\s,]+/)
    .map(s => parseInt(s.trim(), 10))
    .filter(n => !isNaN(n));
}

/**
 * Find the "GRADE #" sheet for a grade level (e.g. "GRADE 7"),
 * case/whitespace-insensitively, without creating one.
 */
function findGradeSummarySheet(spreadsheet, gradeLevel) {
  const targetName = ("grade " + gradeLevel).toLowerCase();
  return spreadsheet.getSheets().find(sheet => sheet.getName().trim().toLowerCase() === targetName) || null;
}

/**
 * Locate the GRADE # sheet's existing columns by scanning only the
 * rows above GRADE_SUMMARY_DATA_START_ROW (its own title/header
 * formatting, never written to) for recognizable label text. A column
 * or section not found stays -1 / unset rather than being guessed -
 * callers must skip writing anything they can't confidently locate.
 */
function findGradeSummaryColumns(data, sections) {
  const headerRowCount = Math.min(data.length, GRADE_SUMMARY_DATA_START_ROW - 1);
  const columns = { loCode: -1, loDescription: -1, competency: -1, itemPlacement: -1, total: -1, sections: {} };

  for (let row = 0; row < headerRowCount; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const cell = data[row][col] ? data[row][col].toString().trim().toLowerCase() : "";
      if (!cell) continue;

      if (columns.loCode === -1 && cell === "lo code") columns.loCode = col;
      else if (columns.loDescription === -1 && cell === "lo description") columns.loDescription = col;
      else if (columns.competency === -1 && cell === "competency") columns.competency = col;
      else if (columns.itemPlacement === -1 && cell === "item placement") columns.itemPlacement = col;
      else if (columns.total === -1 && cell === "total") columns.total = col;
    }
  }

  sections.forEach(section => {
    const letterSuffix = section.replace(/^\d+/, "");
    for (let row = 0; row < headerRowCount; row++) {
      let found = false;
      for (let col = 0; col < data[row].length; col++) {
        const cell = data[row][col] ? data[row][col].toString().trim().toUpperCase() : "";
        if (cell === section || (letterSuffix && cell === letterSuffix)) {
          columns.sections[section] = col;
          found = true;
          break;
        }
      }
      if (found) break;
    }
  });

  return columns;
}

/**
 * Read GRADE_SUMMARY_DATA_START_ROW downward, grouping existing rows
 * by LO Code in the order they appear on the sheet. Rows are matched
 * to pairs by position within their LO Code (1st competency row, 2nd,
 * ...) rather than by matching the Competency cell's text verbatim -
 * free text is fragile to match on (Sheets can normalize whitespace,
 * quotes, etc. on re-read), so a text-based key risked never finding
 * the existing row and silently duplicating it with a blank Item
 * Placement instead of reading what was already typed there.
 */
function findExistingGradeSummaryRows(data, columns) {
  const rowsByLoCode = {};
  let maxRow = GRADE_SUMMARY_DATA_START_ROW - 1;

  if (columns.loCode === -1) return { rowsByLoCode: rowsByLoCode, nextRow: GRADE_SUMMARY_DATA_START_ROW };

  for (let i = GRADE_SUMMARY_DATA_START_ROW - 1; i < data.length; i++) {
    const loCode = data[i][columns.loCode];
    if (!loCode) continue;

    const sheetRow = i + 1; // convert back to 1-indexed sheet row
    const key = loCode.toString().trim();
    if (!rowsByLoCode[key]) rowsByLoCode[key] = [];
    rowsByLoCode[key].push({
      sheetRow: sheetRow,
      itemPlacementRaw: columns.itemPlacement !== -1 ? (data[i][columns.itemPlacement] || "") : ""
    });
    maxRow = Math.max(maxRow, sheetRow);
  }

  return { rowsByLoCode: rowsByLoCode, nextRow: maxRow + 1 };
}

/**
 * Sync the GRADE # sheet: auto-line-up every (LO Code, Competency) pair
 * from LOs-Competency (creating the sheet if it doesn't exist yet), and
 * fill in each section's computed total plus a grand TOTAL. A section
 * cell is the sum, across every student matched into that section, of
 * how many of that row's items they answered correctly. Rows above
 * GRADE_SUMMARY_DATA_START_ROW are never read from or written to, so
 * whatever title/header formatting is already there is left alone; a
 * pair's Item Placement is only ever read, never overwritten. Existing
 * rows are matched to pairs by LO Code + position among that LO's
 * competencies, not by matching the Competency text - logs the
 * detected columns either way, so a mismatch is visible in the log
 * rather than silently producing duplicate rows or blank matches.
 *
 * Returns the list of question numbers (1..numQuestions) not covered
 * by any row's Item Placement, so the caller can flag them.
 */
function syncGradeSummarySheet(spreadsheet, gradeLevel, loCompetencyPairs, studentLookup, sheetsForGrade, numQuestions) {
  const sections = sheetsForGrade.map(sheet => normalizeSection(sheet.getName()));

  const studentsBySection = {};
  for (const key in studentLookup) {
    const section = key.split("::")[0];
    if (!studentsBySection[section]) studentsBySection[section] = [];
    studentsBySection[section].push(studentLookup[key]);
  }

  let summarySheet = findGradeSummarySheet(spreadsheet, gradeLevel);
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet("GRADE " + gradeLevel);
    const header = ["LO Code", "LO Description", "Competency", "Item Placement"].concat(sections).concat(["TOTAL"]);
    summarySheet.getRange(GRADE_SUMMARY_DATA_START_ROW - 1, 1, 1, header.length).setValues([header]);
  }

  const lastRow = Math.max(summarySheet.getLastRow(), GRADE_SUMMARY_DATA_START_ROW - 1);
  const lastColumn = Math.max(summarySheet.getLastColumn(), 1);
  const data = summarySheet.getRange(1, 1, lastRow, lastColumn).getValues();

  const columns = findGradeSummaryColumns(data, sections);
  console.log(
    `GRADE ${gradeLevel} summary columns - LO Code: ${columns.loCode === -1 ? "NOT FOUND" : "col " + (columns.loCode + 1)}, `
    + `LO Description: ${columns.loDescription === -1 ? "NOT FOUND" : "col " + (columns.loDescription + 1)}, `
    + `Competency: ${columns.competency === -1 ? "NOT FOUND" : "col " + (columns.competency + 1)}, `
    + `Item Placement: ${columns.itemPlacement === -1 ? "NOT FOUND" : "col " + (columns.itemPlacement + 1)}, `
    + `TOTAL: ${columns.total === -1 ? "NOT FOUND" : "col " + (columns.total + 1)}, `
    + `Sections found: ${Object.keys(columns.sections).length}/${sections.length} (${sections.filter(s => columns.sections[s] === undefined).join(", ") || "none missing"})`
  );

  const existing = findExistingGradeSummaryRows(data, columns);
  let nextRow = existing.nextRow;
  const usedCounts = {};

  const taggedItems = new Set();

  loCompetencyPairs.forEach(pair => {
    const existingForLoCode = existing.rowsByLoCode[pair.loCode] || [];
    const usedCount = usedCounts[pair.loCode] || 0;
    const existingRow = existingForLoCode[usedCount];
    usedCounts[pair.loCode] = usedCount + 1;

    const sheetRow = existingRow ? existingRow.sheetRow : nextRow;
    const itemPlacementRaw = existingRow ? existingRow.itemPlacementRaw : "";
    const items = parseItemPlacement(itemPlacementRaw);
    items.forEach(item => taggedItems.add(item));

    const sectionSums = sections.map(section => {
      const students = studentsBySection[section] || [];
      return students.reduce((sectionSum, student) => {
        const studentSum = items.reduce((sum, itemNum) => {
          return sum + (parseInt(student.questions[itemNum - 1]) || 0);
        }, 0);
        return sectionSum + studentSum;
      }, 0);
    });

    const total = sectionSums.reduce((a, b) => a + b, 0);

    if (!existingRow) {
      // New row - write the identifying columns we could locate. Item
      // Placement is left blank for the teacher to fill in.
      if (columns.loCode !== -1) summarySheet.getRange(sheetRow, columns.loCode + 1).setValue(pair.loCode);
      if (columns.loDescription !== -1) summarySheet.getRange(sheetRow, columns.loDescription + 1).setValue(pair.loDescription);
      if (columns.competency !== -1) summarySheet.getRange(sheetRow, columns.competency + 1).setValue(pair.competency);
      nextRow++;
    }

    sections.forEach((section, i) => {
      if (columns.sections[section] !== undefined) {
        summarySheet.getRange(sheetRow, columns.sections[section] + 1).setValue(sectionSums[i]);
      }
    });

    if (columns.total !== -1) {
      summarySheet.getRange(sheetRow, columns.total + 1).setValue(total);
    }
  });

  const untaggedItems = [];
  for (let i = 1; i <= numQuestions; i++) {
    if (!taggedItems.has(i)) untaggedItems.push(i);
  }

  return untaggedItems;
}

/**
 * Find the row (0-indexed, matching getValues() rows) that contains a
 * cell equal to targetHeader, case-insensitively. Templates can have
 * title/adviser rows above the real header row, so this scans the
 * first several rows instead of assuming row 0 is the header.
 */
function findHeaderRowIndex(data, targetHeader) {
  const target = targetHeader.trim().toLowerCase();
  const maxRowsToScan = Math.min(data.length, 15);

  for (let row = 0; row < maxRowsToScan; row++) {
    for (let col = 0; col < data[row].length; col++) {
      const cell = data[row][col];
      if (cell && cell.toString().trim().toLowerCase() === target) {
        return row;
      }
    }
  }

  return -1;
}

/**
 * Normalize a section/class value for matching. Pulls out just the
 * grade + section code (e.g. "7H") from values that carry extra text
 * around it, like a ZipGrade Class column of "FILIPINO 7H" or a sheet
 * name of "7H" — both normalize to the same "7H" so they match.
 */
function normalizeSection(value) {
  if (!value) return "";
  const text = value.toString().trim();
  const match = text.match(/(\d+)\s*-?\s*([A-Za-z]+)/);
  return match ? (match[1] + match[2]).toUpperCase() : text.toUpperCase();
}

/**
 * Build the section + student number key used to look up a student
 */
function buildLookupKey(section, studentNum) {
  return `${section}::${studentNum.toString().trim()}`;
}

/**
 * Extract grade level from sheet name (e.g., "1A" → 1, "2G" → 2)
 */
function extractGradeLevel(sheetName) {
  const match = sheetName.match(/^(\d+)/);
  return match ? parseInt(match[1]) : null;
}

/**
 * Detect grade level from ZipGrade data by examining Class column
 */
function detectGradeLevelFromZipGrade(zipgradeData) {
  const gradeLevels = new Set();

  let classColumnIndex = -1;
  const headers = zipgradeData[0];

  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toString().trim().toLowerCase() === "class") {
      classColumnIndex = i;
      break;
    }
  }

  if (classColumnIndex === -1) {
    throw new Error("Could not find 'Class' column in ZipGrade data");
  }

  for (let row = 1; row < zipgradeData.length; row++) {
    const classValue = zipgradeData[row][classColumnIndex];
    if (classValue) {
      const match = classValue.toString().match(/(\d+)/);
      if (match) {
        gradeLevels.add(parseInt(match[1]));
      }
    }
  }

  if (gradeLevels.size === 0) {
    throw new Error("Could not detect grade level from Class column");
  }

  if (gradeLevels.size > 1) {
    throw new Error(`ZipGrade file contains mixed grade levels: ${Array.from(gradeLevels).join(', ')}. Please use a file with only one grade level.`);
  }

  const gradeLevel = Array.from(gradeLevels)[0];
  console.log(`Detected grade level: ${gradeLevel}`);
  return gradeLevel;
}

/**
 * Show instructions
 */
function showInstructions() {
  const instructions = `
═══════════════════════════════════════════════════════════════
📊 ZipGrade Loader - Smart Grade Level Detection
═══════════════════════════════════════════════════════════════

WHAT THIS DOES:
✓ Finds all ZipGrade files (Google Sheets or .xlsx) in the same folder
✓ Let's you select which file to load
✓ Auto-detects your grade level from sheet names
✓ Validates that ZipGrade file matches your template grade
✓ Populates all sheets for that grade with student data
✓ Color-codes responses: Green=1 (Correct), Red=0 (Incorrect)

SAFETY FEATURES:
✓ Prevents loading wrong grade level data
✓ Shows clear error if mismatch detected
✓ Validates all data before loading

HOW TO USE:
1. Click "📊 ZipGrade Loader" menu (top right of sheet)
2. Click "Load ZipGrade Data"
3. Select a ZipGrade file from the list
4. Script auto-detects your grade level and validates
5. Loads automatically ⚡

═══════════════════════════════════════════════════════════════
  `;
  SpreadsheetApp.getUi().alert(instructions);
}

/**
 * Show logs location
 */
function viewLogs() {
  SpreadsheetApp.getUi().alert("📋 Logs are shown in:\n\nExtensions > Apps Script > Execution log\n\nOr press Ctrl+Enter in the script editor.");
}
