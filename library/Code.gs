/**
 * ZIPGRADE DATA LOADER - PRACTICAL SKILLS TEMPLATE
 *
 * Smart grade level detection and validation
 * Automatically finds ZipGrade files in the same folder as template
 * Dynamically detects all template sheets
 * Remembers the template format for future loads
 */

/**
 * Create menu when sheet opens
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

/**
 * Process the selected file
 */
function processSelectedFile(fileId) {
  let tempConvertedFileId = null;
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
    let numCorrectIndex = -1;
    let percentCorrectIndex = -1;
    let classIndex = -1;
    let questionStartIndex = -1;

    for (let i = 0; i < headers.length; i++) {
      const header = headers[i].toString().trim().toLowerCase();
      if (header === "zipgrade id") zipgradeIdIndex = i;
      else if (header === "first name") firstNameIndex = i;
      else if (header === "last name") lastNameIndex = i;
      else if (header === "num correct") numCorrectIndex = i;
      else if (header === "percent correct") percentCorrectIndex = i;
      else if (header === "class") classIndex = i;
      else if (header.match(/^q\d+$/)) {
        if (questionStartIndex === -1) questionStartIndex = i;
      }
    }

    // Validate columns
    if (zipgradeIdIndex === -1) throw new Error("Column 'ZipGrade ID' not found");
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

    // Process sheets for this grade
    let totalMatched = 0;
    const classResults = {};

    for (const templateSheet of sheetsForGrade) {
      const sheetName = templateSheet.getName();
      const section = normalizeSection(sheetName);

      const templateData = templateSheet.getDataRange().getValues();

      // The header row isn't always row 1 - some templates have title
      // or adviser rows above the real column headers, so search for it.
      const headerRowIndex = findHeaderRowIndex(templateData, "student number");

      if (headerRowIndex === -1) {
        console.log(`⚠ Required columns not found in ${sheetName}, skipping`);
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
      // every load - Num Correct, Percent Correct, then Q1...Qn - so the
      // template doesn't need to already have them, and a question-count
      // change between loads doesn't leave stale columns behind.
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

      classResults[sheetName] = sheetMatched;
      console.log(`✓ ${sheetName}: ${sheetMatched} students matched`);
    }

    console.log("\n" + "=".repeat(60));
    console.log("✓ COMPLETE!");
    console.log("=".repeat(60));
    console.log(`Grade ${templateGradeLevel} - Total students matched: ${totalMatched}`);

    SpreadsheetApp.getUi().alert(`✓ Successfully loaded Grade ${templateGradeLevel} data!\n\nTotal students matched: ${totalMatched}\n\nCheck the logs for details.`);

  } catch (error) {
    console.error("ERROR: " + error.message);
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
 * Wipe columns D onward (header row through the last existing row) on
 * a section sheet and rebuild them from the ZipGrade file: Num Correct,
 * Percent Correct, then Q1...Qn. This means the template doesn't need
 * those columns to already exist, and a question-count change between
 * loads doesn't leave stale extra columns behind. Returns the 0-based
 * column indices (matching the getValues() row-array convention used
 * elsewhere) for the three generated columns.
 */
function prepareSectionColumns(sheet, headerRow, numQuestions) {
  const dataHeaders = ["Num Correct", "Percent Correct"];
  for (let q = 1; q <= numQuestions; q++) {
    dataHeaders.push("Q" + q);
  }

  const lastColumn = Math.max(sheet.getLastColumn(), 3 + dataHeaders.length);
  const lastRow = Math.max(sheet.getLastRow(), headerRow);

  sheet.getRange(headerRow, 4, lastRow - headerRow + 1, lastColumn - 3).clear();
  sheet.getRange(headerRow, 4, 1, dataHeaders.length).setValues([dataHeaders]);

  return {
    numCorrectColIndex: 3,
    percentCorrectColIndex: 4,
    questionColStartIndex: 5
  };
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
