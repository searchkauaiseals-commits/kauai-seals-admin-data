/**
 * @file code.gs
 * @version 1.0.0
 * @author Larry Ward - searchkauaiseals
 *
 * @description
 * Administrative backend for the Kauai Seals Google Sheets system.
 * This file implements all server-side logic for:
 *
 *  - UI menus and modal dialogs
 *  - Record CRUD operations (Add / Edit / Delete)
 *  - Search and navigation
 *  - Age/Sz yearly updates
 *  - Backup and retention management
 *  - Activity and function logging
 *
 * The script acts as the orchestration layer between:
 *  - Google Sheets (data persistence)
 *  - HTML front-end forms (AddRecord, EditRecord, Results, etc.)
 *  - Google Drive (backups)
 *
 * Design principles:
 *  - Spreadsheet is the single source of truth.
 *  - All mutations are logged in ActivityLog.
 *  - All user actions can be traced via FunctionLog (optional).
 *  - Business rules (age, Sz, dead status) live only here.
 *
 * Conventions:
 *  - Row 1–4 are headers.
 *  - Data begins at row 5.
 *  - Column D = PermID.
 *  - Column BD = Combined searchable text.
 *  - Column BF = Status (e.g. "dead").
 *
 * Triggers:
 *  - onOpen(): UI initialization.
 *  - Time-based trigger: updateAgeAndSz() (daily, but acts only Jan 1).
 */

/**
 * ============================================================
 * ARCHITECTURE OVERVIEW
 * ============================================================
 *
 * USER (Spreadsheet UI)
 *   |
 *   v
 * Custom Menu (onOpen)
 *   |
 *   +--> Search Records --> openSearchForm()
 *   |                         |
 *   |                         v
 *   |                    sheetSearchButton()
 *   |                         |
 *   |                         v
 *   |                   sheetSearchRecord()
 *   |                         |
 *   |                         v
 *   |              setFocusAndHighlightLater()
 *   |
 *   +--> Add Record --> OpenAddRecordForm()
 *   |                         |
 *   |                         v
 *   |                  addNewRecord()
 *   |                         |
 *   |                         v
 *   |                   logActivity("ADD")
 *   |
 *   +--> Edit Record --> OpenEditDialog()
 *   |                         |
 *   |                         v
 *   |                   updateRecord()
 *   |                         |
 *   |                         v
 *   |                   logActivity("EDIT")
 *   |
 *   +--> Backup File --> backupSpreadsheet()
 *   |                         |
 *   |                         v
 *   |                   manageBackups()
 *   |
 *   +--> January Age Update --> updateAgeAndSz()
 *                             |
 *                             v
 *                        incrementSz()
 *                             |
 *                             v
 *                        logActivity("EDIT")
 *
 * SUPPORT SYSTEMS
 * ------------------------------------------------------------
 * - ActivityLog sheet: immutable audit trail.
 * - FunctionLog sheet: optional debug trace.
 * - PropertiesService: persistent config (folderId, backupLimit).
 */


//----------------------------- Format Sheet and Custom Menu Functions ---------------------------------

/**
 * Triggered automatically when the spreadsheet is opened.
 *
 * Responsibilities:
 *  - Builds the custom "Admin Menu".
 *  - Applies formatting to sheets.
 *  - Clears search box.
 *  - Rebuilds combined BD column formulas.
 * Architectural role:
 *  - System bootstrapper.
 * @trigger onOpen
 */
function onOpen() {
  logCall("onOpen", {});
  var ui = SpreadsheetApp.getUi(); // Get the UI for menu creation
  
  // Create custom menu 'Admin Menu' with sub-items for record management
  ui.createMenu('Admin Menu')
    .addItem('Search Records', 'openSearchForm')
    .addItem('Last Seen', 'getLastSeenByID')  // Option to display 10 previous sightings of specific seal 
    .addItem('Add Record', 'OpenAddRecordForm') // Option to add a new record
    .addItem('Edit Record', 'OpenEditDialog') // Option to edit an existing record
    .addItem('Delete Record', 'OpenDeleteDialog') // Option to delete a record
    .addItem('Backup File', 'backupSpreadsheetUI') // Option to backup the spreadsheet
    .addSubMenu(ui.createMenu('Backup Settings') // Submenu for settings
      .addItem('Set no. of backup', 'setNumberOfBackups') // Option to set number of backups
      .addItem('Destination', 'setDestinationFolder')) // Option to set backup destination folder
    .addItem('January Age Update', 'OpenAgeAndSzDialog') // Option to increment age/Sz/Sz1 on Jan 1.
    .addItem('Sync Seal Log', 'syncSightingLogManual') // Option to manually sync the Complete Seal Log
    .addToUi(); // Add the menu to the UI

  // Call functions to format sheets and clear search box
  formatSheets(); // Custom function to apply formatting
  clearSearchBox(); // Clear search input when the sheet is opened
}

/**
* Applies layout and formatting rules to target sheets.
* - Freezes headers
* - Sets column widths
* - Enables wrapping for long-text columns
*
* @returns {void}
*/
function formatSheets() {
  logCall("formatSheets", {});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsToFormat = ["Data"]; // Array of sheet names to format
  var columnWidths = [76, 80, 62, 71, 65, 42, 44, 76, 54, 77, 75, 67, 67, 77, 75, 67, 67, 209, 296];
  var wrapColumns = [2, 4, 11, 12, 13, 15, 16, 17, 18, 19];

  sheetsToFormat.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return; // Skip if the sheet is not found

    // Freeze the first 4 rows for better readability and static view when scrolling
    sheet.setFrozenRows(4);

    // Set specific column widths for columns 1 to 19 to maintain consistent appearance
    for (var i = 0; i < columnWidths.length; i++) {
      sheet.setColumnWidth(i + 1, columnWidths[i]); // Apply width to each column
    }

    // Set text wrapping for specific columns to handle long text
    wrapColumns.forEach(function(col) {
      sheet.getRange(1, col, sheet.getMaxRows(), 1).setWrap(true); // Enable text wrapping for the entire column
    });
  });
}

/**
* Populates or refreshes the BD column for all active rows.
* - Skips header rows
* - Clears BD for dead records (BF = 'dead')
* - Appends normalized T:BB values when present
*
* This is a bulk operation intended for onOpen or maintenance runs.
*
* @returns {void}
*/
function populateCombinedColumnFormula() {
  logCall("populateCombinedColumnFormula", {});
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  // Column indices (1-based)
  const bfColumnIndex = 58; // BF
  const bdColumnIndex = 56; // BD

  const lastRow = sheet.getLastRow();

  if (lastRow < 5) {
    Logger.log('No data found from row 5 onwards in Column A.');
    return;
  }


/*
  // Find last row with data in column A starting at row 5
  const lastRowRange = sheet.getRange("A5:A");
  const nextDataCell = lastRowRange.getNextDataCell(SpreadsheetApp.Direction.DOWN);
  const lastRow = nextDataCell.getRow();

  if (lastRow < 5 || nextDataCell.getValue() === "") {
    Logger.log('No data found from row 5 onwards in Column A. Exiting populateCombinedColumnFormula.');
    return;
  }
*/
  Logger.log('Last Row with data in Column A: ' + lastRow);

  // 🔑 FIX: Use DISPLAY values so formula results are captured
  const sourceDataTBB = sheet
    .getRange('T5:BB' + lastRow)
    .getDisplayValues();

  // Base BD formula (static portion)
  const baseBdFormula = `=LOWER(TRIM(TEXTJOIN(" ", TRUE,
    IF(IFERROR(INDEX($H:$H, ROW()), "") = "male", CHAR(160) & LOWER(IFERROR(INDEX($H:$H, ROW()), "")), LOWER(IFERROR(INDEX($H:$H, ROW()), ""))),
    IF(IFERROR(INDEX($I:$I, ROW()), "") = "tag", CHAR(160) & LOWER(IFERROR(INDEX($I:$I, ROW()), "")), LOWER(IFERROR(INDEX($I:$I, ROW()), ""))),
    TEXT(IFERROR(INDEX($C:$C, ROW()), ""), "mm/dd/yyyy"),
    LOWER(TRIM(TEXTJOIN(" ", TRUE,
      IFERROR(INDEX($A:$B, ROW()), ""),
      IFERROR(INDEX($D:$G, ROW()), ""),
      IFERROR(INDEX($J:$R, ROW()), "")
    )))
  )))`;

  // Iterate rows
  for (let row = 5; row <= lastRow; row++) {
    const bfValue = sheet.getRange(row, bfColumnIndex).getValue();
    const isDead =
      typeof bfValue === 'string' &&
      bfValue.trim().toLowerCase() === 'dead';

    if (isDead) {
      sheet.getRange(row, bdColumnIndex).clearContent();
      Logger.log(`Row ${row}: BF is 'dead'. Cleared BD content.`);
      continue;
    }

    // Build T:BB string (normalized)
    const concatRowString = sourceDataTBB[row - 5]
      .map(cell => cell.toLowerCase().trim())
      .filter(cell => cell !== "")
      .map(cell => cell.replace(/"/g, '""'))
      .join(" ");

    // 🔑 Skip comma if nothing to append
    const finalFormula = concatRowString
      ? `${baseBdFormula} & ", ${concatRowString},"`
      : baseBdFormula;

    Logger.log(`Row ${row}: Setting formula in BD: ${finalFormula}`);
    sheet.getRange(row, bdColumnIndex).setFormula(finalFormula);
  }
}


/*
function populateCombinedColumnFormula() {
  logCall("populateCombinedColumnFormula", {});
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  // Define column indices for clarity
  const bfColumnIndex = 58; // Column BF (1-indexed)
  const bdColumnIndex = 56; // Column BD (1-indexed)

  // Find the last row with data in column A, starting from A5 downwards
  // Check if there are any rows with data below row 4
  const lastRowRange = sheet.getRange("A5:A");
  const nextDataCell = lastRowRange.getNextDataCell(SpreadsheetApp.Direction.DOWN);
  const lastRow = nextDataCell.getRow();

  if (lastRow < 5 || nextDataCell.getValue() === "") {
    Logger.log('No data found from row 5 onwards in Column A. Exiting populateCombinedColumnFormula.');
    return; // Exit if there is no data beyond row 4
  }

  Logger.log('Last Row with data in Column A: ' + lastRow);

  // Get data for T:BB for all relevant rows in one go for efficiency
  const sourceDataTBB = sheet.getRange('T5:BB' + lastRow).getValues();

  // Define the base formula for BD once
  const baseBdFormula = `=LOWER(TRIM(TEXTJOIN(" ", TRUE,
    IF(IFERROR(INDEX($H:$H, ROW()), "") = "male", CHAR(160) & LOWER(IFERROR(INDEX($H:$H, ROW()), "")), LOWER(IFERROR(INDEX($H:$H, ROW()), ""))),
    IF(IFERROR(INDEX($I:$I, ROW()), "") = "tag", CHAR(160) & LOWER(IFERROR(INDEX($I:$I, ROW()), "")), LOWER(IFERROR(INDEX($I:$I, ROW()), ""))),
    TEXT(IFERROR(INDEX($C:$C, ROW()), ""), "mm/dd/yyyy"),
    LOWER(TRIM(TEXTJOIN(" ", TRUE, IFERROR(INDEX($A:$B, ROW()), ""), IFERROR(INDEX($D:$G, ROW()), ""), IFERROR(INDEX($J:$R, ROW()), ""))))
    )))`; // Note: No trailing comma here, it will be added dynamically

  // Iterate from row 5 to the last data row
  for (let row = 5; row <= lastRow; row++) {
    const bfValue = sheet.getRange(row, bfColumnIndex).getValue();
    const isDead = typeof bfValue === 'string' && bfValue.trim().toLowerCase() === 'dead';

    if (isDead) {
      // If the record is 'dead', clear BD content for this row
      sheet.getRange(row, bdColumnIndex).clearContent();
      Logger.log(`Row ${row}: BF is 'dead'. Cleared BD content.`);
    } else {
      // If not 'dead', construct and set the formula
      const concatRowString = sourceDataTBB[row - 5] // Adjust for zero-based index
                                .filter(cell => cell !== "")
                                .map(cell => String(cell).replace(/"/g, '""')) // Ensure cell is string before replace
                                .join(" ");

      const finalFormula = `${baseBdFormula} & ", ${concatRowString}"`;

      Logger.log(`Row ${row}: Setting formula in BD: ${finalFormula}`);
      sheet.getRange(row, bdColumnIndex).setFormula(finalFormula);
    }
  }
}
*/
function createSheetTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < allTriggers.length; i++) {
    ScriptApp.deleteTrigger(allTriggers[i]);
  }
  Logger.log("All existing triggers for this project have been deleted.");

  ScriptApp.newTrigger('updateAgeAndSz')
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  Logger.log("Daily trigger for 'updateAgeAndSz' created (runs daily at 2:00 AM-2:59 AM, but logic only executes Jan 1st).");

  ScriptApp.newTrigger('backupSpreadsheet')
    .timeBased()
    .everyDays(1)
    .atHour(22)
    .create();
  Logger.log("Daily trigger for 'backupSpreadsheet' created (10:00 PM-10:59 PM).");

  ScriptApp.newTrigger('syncCompleteSealLogTrigger')
    .timeBased()
    .everyHours(2)
    .create();
  Logger.log("Every-2-hour trigger for 'syncCompleteSealLogTrigger' created (runs 7am-7pm Hawaii only).");

  Logger.log("All specified triggers have been created successfully.");
}

/**
 * Returns true if the current time falls within the sighting sync window
 * (7:00am–7:00pm Hawaii time, Pacific/Honolulu).
 * Used by syncCompleteSealLogTrigger() to skip overnight runs.
 *
 * @returns {boolean}
 */
function isSyncWindowOpen() {
  const now = new Date();
  const hawaiiTime = new Date(now.toLocaleString("en-US", { timeZone: "Pacific/Honolulu" }));
  const hour = hawaiiTime.getHours();
  return hour >= 7 && hour < 19;
}

/**
 * Time-trigger wrapper for syncCompleteSealLog().
 * Skips the sync if called outside 7am–7pm Hawaii time.
 * Called every 2 hours by the time-based trigger in createSheetTriggers().
 */
function syncCompleteSealLogTrigger() {
  if (!isSyncWindowOpen()) {
    Logger.log("syncCompleteSealLogTrigger: Outside sync window. Skipping.");
    return;
  }
  try {
    syncCompleteSealLog();
  } catch (e) {
    Logger.log("syncCompleteSealLogTrigger: ERROR — " + e.message);
  }
}

/**
 * Manual sync wrapper called from Admin Menu → Sync Seal Log.
 * Bypasses the time window check since the admin is explicitly requesting it.
 * Shows a UI alert on completion or failure.
 */
function syncSightingLogManual() {
  try {
    syncCompleteSealLog();
    SpreadsheetApp.getUi().alert("Seal Log sync completed successfully.");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Seal Log sync failed: " + e.message);
  }
}

/**
* Simple onEdit dispatcher for the Data sheet.
* - J2 checkbox toggles dead rows
* - Edits in T:BB or BF trigger row-level BD rebuild
*
* @param {GoogleAppsScript.Events.SheetsOnEdit} e
* @returns {void}
*/
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  // Only run on Data sheet
  if (sheet.getName() !== 'Data') return;

  const row = range.getRow();
  const col = range.getColumn();
  const a1 = range.getA1Notation();

  // -----------------------------
  // 1️⃣ Existing behavior (J2)
  // -----------------------------
  if (a1 === 'J2') {
    toggleDeadRows();
    return;
  }

  // Ignore header rows
  if (row < 5) return;

  // -----------------------------
  // 2️⃣ Auto-update BD when T:BB changes
  // -----------------------------
  const isTtoBB = col >= 20 && col <= 54; // T → BB
  const isBF = col === 58;                // BF (dead)

  if (isTtoBB || isBF) {
    populateCombinedColumnFormulaForRow(row);
  }
}

/**
* Rebuilds the BD formula for a single row.
* Used by onEdit to avoid expensive full-sheet recalculation.
*
* @param {number} row - 1-based row number to update
* @returns {void}
*/
function populateCombinedColumnFormulaForRow(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const bfColumnIndex = 58; // BF
  const bdColumnIndex = 56; // BD

  const bfValue = sheet.getRange(row, bfColumnIndex).getValue();
  const isDead =
    typeof bfValue === 'string' &&
    bfValue.trim().toLowerCase() === 'dead';

  if (isDead) {
    sheet.getRange(row, bdColumnIndex).clearContent();
    return;
  }

  const sourceRow = sheet
    .getRange('T' + row + ':BB' + row)
    .getDisplayValues()[0];

  const baseBdFormula = `=LOWER(TRIM(TEXTJOIN(" ", TRUE,
    IF(IFERROR(INDEX($H:$H, ROW()), "") = "male", CHAR(160) & LOWER(IFERROR(INDEX($H:$H, ROW()), "")), LOWER(IFERROR(INDEX($H:$H, ROW()), ""))),
    IF(IFERROR(INDEX($I:$I, ROW()), "") = "tag", CHAR(160) & LOWER(IFERROR(INDEX($I:$I, ROW()), "")), LOWER(IFERROR(INDEX($I:$I, ROW()), ""))),
    TEXT(IFERROR(INDEX($C:$C, ROW()), ""), "mm/dd/yyyy"),
    LOWER(TRIM(TEXTJOIN(" ", TRUE,
      IFERROR(INDEX($A:$B, ROW()), ""),
      IFERROR(INDEX($D:$G, ROW()), ""),
      IFERROR(INDEX($J:$R, ROW()), "")
    )))
  )))`;

  const concatRowString = sourceRow
    .map(cell => cell.toLowerCase().trim())
    .filter(cell => cell !== "")
    .map(cell => cell.replace(/"/g, '""'))
    .join(" ");

  const finalFormula = concatRowString
    ? `${baseBdFormula} & " ${concatRowString}"`
    : baseBdFormula;

  sheet.getRange(row, bdColumnIndex).setFormula(finalFormula);
}


/**
 * Toggles the visibility of rows in the 'Data' sheet based on the value in Column BF.
 * Rows with 'dead' in Column BF are hidden or shown.
 * This function is intended to be linked directly to a button on the sheet.
 * It uses cell J2 to maintain the current visibility state.
 * @returns {void}
 */
function toggleDeadRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const checkBoxCell = sheet.getRange('J2');
  const showAll = checkBoxCell.getValue() === true;

  const dataStartRow = 5;
  const bfCol = 58;
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - dataStartRow + 1;

  if (numRows < 1) return;

  const bfValues = sheet.getRange(dataStartRow, bfCol, numRows, 1).getValues();

  for (let i = 0; i < bfValues.length; i++) {
    const rowNum = dataStartRow + i;
    const value = String(bfValues[i][0] || '').trim().toLowerCase();

    if (value === 'dead') {
      if (showAll) {
        sheet.unhideRow(sheet.getRange(rowNum, 1));
      } else {
        sheet.hideRows(rowNum);
      }
    } else {
      sheet.unhideRow(sheet.getRange(rowNum, 1)); // Always show live seals
    }
  }

  // Always keep header rows unhidden
  for (let r = 1; r <= 4; r++) {
    sheet.unhideRow(sheet.getRange(r, 1));
  }

  SpreadsheetApp.flush();
}

/**
 * Retrieves recent non-dead sightings from the Data sheet.
 *
 * This function:
 *  - Scans column D to find the last populated row.
 *  - Reads rows A5:BF{lastRow}.
 *  - Filters out records where status (column BF) = "dead".
 *  - Returns a compact dataset for UI display.
 *
 * Returned fields (in order):
 *  1. PermID        (column D)
 *  2. Date          (column E)
 *  3. Location      (column F)
 *  4. Sex           (column H)
 *  5. Tag           (column J)
 *  6. Descriptive   (column N)
 *  7. Site          (column R)
 *  8. Region        (column S)
 *  9. CombinedText  (column BD)
 *
 * Architectural role:
 *  - Read-only reporting endpoint.
 *  - Used by "Last Seen" menu feature.
 *  - Never mutates spreadsheet state.
 *
 * @returns {Array<Array<string>>}
 *   Array of rows for recent sightings display.
 */
function getRecentSightingsData() {
  const externalId = "1uFhaG4Vc_LyfAFmcA8X5cFnFNQJNJ_g8Z-z49JD_6VM"; // external spreadsheet
  const extSS = SpreadsheetApp.openById(externalId); // ✅ open external spreadsheet
  const colsToExtract = [1, 4, 5, 10, 11]; // A,D,E,J,K (1-based)

  const extractRange = (sheetName) => {
    const sheet = extSS.getSheetByName(sheetName); // ✅ use external spreadsheet
    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found.`);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const values = sheet.getRange(2, 1, lastRow - 1, Math.max(...colsToExtract)).getValues();
    return values.map(row => colsToExtract.map(col => row[col - 1]));
  };

  const monthData = extractRange("Month");
  const yearData = extractRange("Year");

  const combined = [...monthData, ...yearData];

  // Sort by date (column 0, which is Col A) descending
  combined.sort((a, b) => new Date(b[0]) - new Date(a[0]));

  // Sort by date (column 0 = A) descending
  combined.sort((a, b) => new Date(b[0]) - new Date(a[0]));

  // ✅ Log first 10 rows
  Logger.log("First 10 recent sightings:");
  combined.slice(0, 10).forEach((row, i) => {
    Logger.log(`${i + 1}: ${row.join(" | ")}`);
  });

  return combined;
}

/**
 * Synchronizes the local "Complete Seal Log" sheet with the master
 * seal log data stored in the external mirror spreadsheet.
 *
 * This function:
 *  - Creates the "Complete Seal Log" sheet if it does not exist.
 *  - Pulls seal sighting records from:
 *      • "Month" tab
 *      • "Year" tab
 *      • Yearly tabs from 2020 up to last year (e.g. "2020", "2021", ...)
 *  - Extracts and normalizes the following fields:
 *      Date, Location, Descriptive Location, ID Perm, ID Temp
 *  - Sorts all entries by most recent date first.
 *  - Clears only columns A–E in the local sheet (preserving formulas elsewhere).
 *  - Writes the refreshed data back to columns A–E.
 *  - Formats the Date column.
 *
 * Intended to be run as a scheduled job (e.g. nightly refresh).
 *
 * @throws {Error} If the source spreadsheet cannot be accessed or data processing fails.
 */
function syncCompleteSealLog() {
  const SOURCE_SEAL_LOG_SHEET_ID =
    '1uAi3NYrkHM-79SzdCkH4xX7OKYFW4KwhIxKiwHujZdA';

  const LOCAL_COMPLETE_LOG_TAB = 'Complete Seal Log';

  try {
    Logger.log("=== syncCompleteSealLog START ===");

    const localSS = SpreadsheetApp.getActiveSpreadsheet();
    let targetSheet = localSS.getSheetByName(LOCAL_COMPLETE_LOG_TAB);

    // Create sheet if it doesn't exist
    if (!targetSheet) {
      Logger.log("Creating Complete Seal Log sheet...");
      targetSheet = localSS.insertSheet(LOCAL_COMPLETE_LOG_TAB);

      const headers = [
        'Date',
        'Location',
        'Descriptive Location',
        'ID Perm',
        'ID Temp'
      ];

      targetSheet.appendRow(headers);

      targetSheet
        .getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setFontColor('#ffffff')
        .setBackground('#3B82F6');

      targetSheet.setFrozenRows(1);
    }

    // Open source spreadsheet
    Logger.log("Opening source sheet: " + SOURCE_SEAL_LOG_SHEET_ID);
    const sourceSpreadsheet =
      SpreadsheetApp.openById(SOURCE_SEAL_LOG_SHEET_ID);

    const currentYear = new Date().getFullYear();
    const validRows = [];

    // ---------------------------------------------------------
    // Process ALL Year sheets dynamically (2020, 2021, ... 2026, etc.)
    // ---------------------------------------------------------
    const allSheets = sourceSpreadsheet.getSheets();
    
    allSheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const yearNum = parseInt(sheetName);

      // Only process 4-digit numeric sheets (the yearly archives)
      if (!isNaN(yearNum) && sheetName.length === 4) {
        Logger.log(`Processing Year: ${sheetName}`);
        
        let data;
        if (yearNum === 2020) {
          data = extractSealLogData2020(sheet);
        } else {
          data = extractSealLogData(sheet);
        }
        
        if (data && data.length > 0) {
          validRows.push(...data);
        }
      }
    });

    Logger.log("Total rows collected: " + validRows.length);

    // Sort newest first
    validRows.sort((a, b) => new Date(b[0]) - new Date(a[0]));

    // -----------------------------
    // Clear ONLY columns A:E
    // -----------------------------
    const targetLastRow = targetSheet.getLastRow();
    if (targetLastRow > 1) {
      targetSheet
        .getRange(2, 1, targetLastRow - 1, 5)
        .clearContent();
    }

    // -----------------------------
    // Write new data (A:E only)
    // -----------------------------
    if (validRows.length > 0) {
      targetSheet
        .getRange(2, 1, validRows.length, 5)
        .setValues(validRows);

      targetSheet
        .getRange(2, 1, validRows.length, 1)
        .setNumberFormat('M/d/yyyy');
    }

    Logger.log(
      `SUCCESS: Synced ${validRows.length} complete seal log entries`
    );
    Logger.log("=== syncCompleteSealLog END ===");

  } catch (e) {
    Logger.log("ERROR (syncCompleteSealLog): " + e.message);
    Logger.log("Stack: " + e.stack);
    throw new Error(
      "Failed to sync complete seal log: " + e.message
    );
  }
}



/**
 * Extracts and normalizes seal sighting data from a given source sheet.
 *
 * Reads raw rows from the provided sheet and returns only valid records
 * containing a parsable date. Each output row is normalized into the
 * standard format used by the Complete Seal Log:
 *
 * Output columns:
 *  [0] Date (Date object)
 *  [1] Location (string)
 *  [2] Descriptive Location (string)
 *  [3] ID Perm (string or "-")
 *  [4] ID Temp (string, normalized to "Temp X" or "PK...")
 *
 * Normalization rules:
 *  - Rows with empty or invalid dates are skipped.
 *  - Empty ID Perm becomes "-".
 *  - ID Temp:
 *      • If empty → "-"
 *      • If starts with "PK" → left as-is
 *      • Otherwise → prefixed with "Temp "
 *
 * Extracts columns  - A, D, E, J, K
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 *        Source sheet to extract seal log data from.
 *
 * @returns {Array<Array<any>>}
 *          A 2D array of normalized seal log rows.
 */
function extractSealLogData(sheet) {
  const values = sheet.getDataRange().getValues();
  const extractedRows = [];
  
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rawDate = row[0]; // Column A
    
    if (!rawDate) continue;
    
    const parsedDate = new Date(rawDate);
    if (isNaN(parsedDate)) continue;
    
    const location = row[3] || '';           // Column D
    const descriptiveLocation = row[4] || ''; // Column E
    const idPerm = row[9];                    // Column J
    const idTempRaw = row[10]; // Column K

    let idTempFinal = '-';

    if (idTempRaw !== '' && idTempRaw != null) {
      const idTempStr = String(idTempRaw).trim();

      // Prepend "Temp " unless value already starts with PK
      if (/^PK/i.test(idTempStr)) {
        idTempFinal = idTempStr;
      } else if (!/^Temp\s+/i.test(idTempStr)) {
        idTempFinal = `Temp ${idTempStr}`;
      } else {
        idTempFinal = idTempStr;
      }
    }

    extractedRows.push([
      parsedDate,
      location,
      descriptiveLocation,
      idPerm === '' || idPerm == null ? '-' : idPerm,
      idTempFinal
    ]);
  }
  return extractedRows;
}

/**
 * Extract 2020 seal log data (columns: A,D,E,I,J)
 */
function extractSealLogData2020(sheet) {
  const values = sheet.getDataRange().getValues();
  const extractedRows = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rawDate = row[0]; // Column A
    if (!rawDate) continue;

    const parsedDate = new Date(rawDate);
    if (isNaN(parsedDate)) continue;

    const location = row[3] || '';           // Column D
    const descriptiveLocation = row[4] || ''; // Column E
    const idPerm = row[8];                    // Column I
    const idTempRaw = row[9];                 // Column J

    let idTempFinal = '-';
    if (idTempRaw !== '' && idTempRaw != null) {
      const idTempStr = String(idTempRaw).trim();
      if (/^PK/i.test(idTempStr)) {
        idTempFinal = idTempStr;
      } else if (!/^Temp\s+/i.test(idTempStr)) {
        idTempFinal = `Temp ${idTempStr}`;
      } else {
        idTempFinal = idTempStr;
      }
    }

    extractedRows.push([
      parsedDate,
      location,
      descriptiveLocation,
      idPerm === '' || idPerm == null ? '-' : idPerm,
      idTempFinal
    ]);
  }

  return extractedRows;
}


/**
 * Displays last sightings for a PermID.
 *
 * Architectural role:
 *  - Read-only analytics endpoint.
 */
function getLastSeenByID() {
  const ui = SpreadsheetApp.getUi();
  const idPrompt = ui.prompt(
    "Last Seen",
    "Enter the PermID or TempID to search for:",
    ui.ButtonSet.OK_CANCEL
  );
  if (idPrompt.getSelectedButton() !== ui.Button.OK) return;

  const searchId = idPrompt.getResponseText().trim().toLowerCase();
  if (!searchId) {
    ui.alert("Invalid ID entered.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Complete Seal Log");
  if (!sheet) {
    ui.alert("Sheet 'Complete Seal Log' not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("Complete Seal Log is empty.");
    return;
  }

  // A–E = Date, Location, Desc Location, ID Perm, ID Temp
  const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();

  const matches = data.filter(row =>
    String(row[3]).trim().toLowerCase() === searchId ||
    String(row[4]).trim().toLowerCase() === searchId
  );

  if (matches.length === 0) {
    ui.alert("No sightings found for ID: " + searchId);
    return;
  }

  // Sort by Date descending
  matches.sort((a, b) => new Date(b[0]) - new Date(a[0]));

  // -----------------------------
  // Top 10 sightings display
  // -----------------------------
  const headers = ["Date", "Location", "Descriptive Location"];
  const maxRows = 10;

  const formattedRows = matches.slice(0, maxRows).map(row => {
    const parts = [];

    // Date (always included)
    parts.push(formatDate(row[0]));

    // Location
    if (row[1] && row[1] !== "-") {
      parts.push(row[1]);
    }

    // Descriptive Location
    if (row[2] && row[2] !== "-") {
      parts.push(row[2]);
    }

    // IMPORTANT:
    // Do NOT include ID Perm (row[3])
    // Do NOT include ID Temp (row[4])

    return parts.join("  •  ");
  });


  // -----------------------------
  // Yearly summary block
  // -----------------------------
  const currentYear = new Date().getFullYear();
  const minYear = 2020;

  const yearCounts = {};

  matches.forEach(row => {
    const year = new Date(row[0]).getFullYear();
    if (year >= minYear && year <= currentYear) {
      yearCounts[year] = (yearCounts[year] || 0) + 1;
    }
  });

  // Build ordered year list (descending)
  const orderedYears = [];
  for (let y = currentYear; y >= minYear; y--) {
    if (yearCounts[y]) {
      orderedYears.push(`${y} - ${yearCounts[y]}`);
    }
  }

  // Format into aligned columns (alert-safe)
  const columnCount = 3;
  const rowsPerColumn = Math.ceil(orderedYears.length / columnCount);

  // Build grid first
  const grid = Array.from({ length: rowsPerColumn }, () =>
    Array(columnCount).fill("")
  );

  for (let i = 0; i < orderedYears.length; i++) {
    const col = Math.floor(i / rowsPerColumn);
    const row = i % rowsPerColumn;
    grid[row][col] = orderedYears[i];
  }

  // Find max width per column
  const colWidths = [];

  for (let c = 0; c < columnCount; c++) {
    colWidths[c] = Math.max(
      ...grid.map(row => row[c].length)
    );
  }

  // Build formatted rows
  // Add empty spacer columns between data columns
  const MIN_COL_WIDTH = 12;
  const SPACER = " "; // single space, but as a real column

  const adjustedWidths = colWidths.map(w =>
    Math.max(w, MIN_COL_WIDTH)
  );

  const DOT_SPACER = "  ................  ";

  const summaryRows = grid.map(row =>
    row
      .filter(cell => cell) // remove empty cells
      .join(DOT_SPACER)
  );

  const summaryBlock =
    summaryRows.length > 0
      ? "\n\nYearly Totals:\n" + summaryRows.join("\n")
      : "";

  // -----------------------------
  // Final output
  // -----------------------------
  const output =
    [headers.join(" | "), ...formattedRows].join("\n") +
    summaryBlock;

  const displayId = formatDisplayId(searchId);
  ui.alert(`Recent Sightings for ${displayId}:\n\n${output}`);
}

function formatDisplayId(rawId) {
  const id = String(rawId).trim();

  // Match temp IDs (temp 625, Temp625, TEMP 625, etc.)
  const tempMatch = id.match(/^temp\s*(.+)$/i);

  if (tempMatch) {
    // Normalize to "Temp ###"
    return `Temp ${tempMatch[1].trim()}`;
  }

  // Otherwise uppercase everything
  return id.toUpperCase();
}

// Optional date formatter
function formatDate(date) {
  if (Object.prototype.toString.call(date) === "[object Date]" && !isNaN(date)) {
    return Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return date;
}

function getModalData() {
  logCall("getModalData", {});
  Logger.log("fxn getModalData called");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  // Find last non-empty row in column D (index 4)
  const colD = sheet.getRange("D5:D").getValues();
  let lastRow = 4;
  for (let i = colD.length - 1; i >= 0; i--) {
    if (colD[i][0] !== "") {
      lastRow = i + 5;
      break;
    }
  }

  // Get range A5:BF{lastRow} to include BF (col 58) for filtering only
  const data = sheet.getRange('A5:BF' + lastRow).getValues();

  // Filter out rows where BF (index 57) = 'Dead'
  const filtered = data.filter(row => {
    const status = (row[57] || "").toString().trim().toLowerCase();
    return status !== "dead";
  });

  // Return only these columns: D (3), E (4), F (5), H (7), J (9), N (13), R (17), S (18), BD (55)
  return filtered.map(row => [row[3], row[4], row[5], row[7], row[9], row[13], row[17], row[18], row[55]]);
}

// Function to clear the contents of the search box and set focus back to it
function clearSearchBox() {
  logCall("clearSearchBox", {});
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Search'); // Get the "Search" sheet
  sheet.getRange('sheetSearchBox').clearContent(); // Clear the content of the search box range
}

// Function to set focus to the search box DATA SHEET
function setFocusToSearchBox() {
  logCall("setFocusToSearchBox", {});
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Search'); // Get the "Search" sheet
  var cell = sheet.getRange('sheetSearchBox'); // Get the search box range
  sheet.setActiveRange(cell); // Set the search box as the active cell
}

/**
 * Opens the main search interface in a modal dialog.
 *
 * Architectural role:
 *  - UI gateway to Search subsystem.
 */
function openSearchForm() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Results')
      .setWidth(1100)
      .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Search Records');
}

/**
 * Opens the Add Record form.
 *
 * Architectural role:
 *  - UI entry point for CREATE pipeline.
 */
function OpenAddRecordForm() {
  logCall("OpenAddRecordForm", {});
  var htmlOutput = HtmlService.createHtmlOutputFromFile('AddRecord') // Load HTML file for the form
      .setWidth(1060) // Set dialog width
      .setHeight(560); // Set dialog height
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Add Record'); // Display the form
}

/**
 * Opens the Edit Record form.
 *
 * Architectural role:
 *  - UI entry point for UPDATE pipeline.
 */
function OpenEditDialog() {
  logCall("OpenEditDialog", {});
  var htmlOutput = HtmlService.createHtmlOutputFromFile('EditRecord') // Load HTML file for the form
      .setWidth(1060) // Set dialog width
      .setHeight(565); // Set dialog height
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Edit Record'); // Display the form
}

/**
 * Opens the Delete Record form.
 *
 * Architectural role:
 *  - UI entry point for DELETE pipeline.
 */
function OpenDeleteDialog() {
  logCall("OpenDeleteDialog", {});
  var htmlOutput = HtmlService.createHtmlOutputFromFile('DeleteRecord') // Load HTML file for the form
      .setWidth(1060) // Set dialog width
      .setHeight(570); // Set dialog height
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Delete Record'); // Display the form
}

/**
 * Opens the January Age/Sz update dialog.
 *
 * Architectural role:
 *  - Manual trigger for annual automation.
 */
function OpenAgeAndSzDialog() {
  logCall("OpenAgeAndSzDialog", {});
  var htmlOutput = HtmlService.createHtmlOutputFromFile('Jan1Modal') // Load HTML file for the form
      .setWidth(600) // Set dialog width
      .setHeight(200); // Set dialog height
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Jan 1st Age Update'); // Display the form
}

/**
 * Central server-side error handler.
 *
 * @param {Error} error
 */
function handleError(error) {
  logCall("handleError", {error: error});
  console.error(error); // Log the error to the console for debugging
  alert('Error: Failed to submit data!'); // Display an alert to the user in case of failure
}

/**
 * Logs functions as they are called to the "FunctionLog" sheet of the active spreadsheet.
 * To use function, set isLoggingEnabled = true.  To disable, set = false.
 * 
 * The FunctionLog sheet is hidden.  Unhide to view log !
 * 
 * @param {string} functionName - The name of the function called.
 * @param {string} params - The parameters passed with the function call.
 */

// Global variable to enable or disable logging
var isLoggingEnabled = false; // Set to true when you want to enable logging

/**
* Appends a function call entry to the hidden `functionLog` sheet.
* Intended for debugging and audit tracing.
*
* @param {string} functionName
* @param {Object} params
* @returns {void}
*/
function logCall(functionName, params) {
  if (!isLoggingEnabled) return; // Exit if logging is disabled
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('functionLog');
  if (sheet) {
    // Prepare the log details
    var timestamp = new Date();
    var paramString = JSON.stringify(params);  // Convert the params object to a string
    
    // Append the log entry as a new row
    sheet.appendRow([timestamp, functionName, paramString]);
  } else {
    Logger.log("Log sheet not found: functionLog");
  }
}

//---------------------------------- Backup Spreadsheet Functions ---------------------------------
/**
 * Creates a timestamped backup copy of the active spreadsheet
 * in the configured Google Drive folder.
 *
 * After backup, enforces retention rules via manageBackups().
 *
 * @returns {void}
 */
// Global variables
var properties = PropertiesService.getScriptProperties();
var backupLimit = properties.getProperty('backupLimit') ? parseInt(properties.getProperty('backupLimit')) : 5; // Default backup limit
var folderId = properties.getProperty('folderId') ? properties.getProperty('folderId') : ''; // Folder ID for backups

function backupSpreadsheetUI() {
  try {
    backupSpreadsheet();
    SpreadsheetApp.getUi().alert('Backup completed successfully.');
  } catch (e) {
    SpreadsheetApp.getUi().alert('Backup failed: ' + e.message);
  }
}

function backupSpreadsheet() {
  logCall("backupSpreadsheet", {});

  if (folderId === '') {
    Logger.log('ERROR: Destination folder is not set.');
    return;
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var destinationFolder;

  try {
    destinationFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('ERROR: Failed to access folder: ' + e.message);
    return;
  }

  try {
    DriveApp.getFileById(ss.getId()).makeCopy(
      ss.getName() + " Backup " +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"),
      destinationFolder
    );
  } catch (e) {
    Logger.log('ERROR: Backup failed: ' + e.message);
    return;
  }

  manageBackups();
}

function setNumberOfBackups() {
  logCall("setNumberofBackups", {});
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Set Number of Backups', 'Enter the number of backups to keep:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var num = parseInt(response.getResponseText());
    if(isNaN(num) || num < 1) {
      ui.alert('Invalid number. Please enter a valid integer greater than 0.');
      return;
    }
    backupLimit = num;
    properties.setProperty('backupLimit', num.toString());
    ui.alert('Backup limit set to ' + backupLimit);
  }
}

function setDestinationFolder() {
  logCall("setDestinationFolder", {});
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Destination Folder', 'Enter the Google Drive Folder ID:', ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    folderId = response.getResponseText();
    properties.setProperty('folderId', folderId);
    ui.alert('Destination folder set.' + folderId);
  }
}

/**
 * Enforces backup retention policy.
 * Deletes (trashes) oldest backups if count exceeds backupLimit.
 *
 * Uses:
 *  - folderId from PropertiesService
 *  - backupLimit from PropertiesService
 *
 * @returns {void}
 */
function manageBackups() {
  logCall("manageBackups", {});
  var folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (e) {
    Logger.log('Error accessing folder with ID: ' + folderId + '. Error: ' + e.toString());
    return;
  }

  var files = folder.getFiles();
  var fileArray = [];
  while (files.hasNext()) {
    var file = files.next();
    fileArray.push({id: file.getId(), date: file.getDateCreated()});
  }
  if (fileArray.length > backupLimit) {
    fileArray.sort(function(a, b) { return a.date - b.date; }); // Sort by date, oldest first
    for (var i = 0; i < fileArray.length - backupLimit; i++) {
      DriveApp.getFileById(fileArray[i].id).setTrashed(true); // Trash the oldest files beyond limit
    }
  }
}

//---------------------------------- Age and Increment Sz Functions --------------------------------
/**
 * Testing utility for age automation.
 * Resets 'lastupdatedate' to Jan 1 of previous year
 */
function resetLastUpdateDate() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Clear the existing property
  scriptProperties.deleteProperty("lastUpdateDate");

  // Get the current year
  const currentYear = new Date().getFullYear();
  
  // Calculate the previous year
  const previousYear = currentYear - 1;

  // Set the new date to January 1 of the previous year
  const newDate = new Date(previousYear, 0, 1); // January is 0
  
  // Save the new date in ISO string format
  scriptProperties.setProperty("lastUpdateDate", newDate.toISOString());
}

/**
 * Annual automation.
 *
 * Increments Age and Sz.
 *
 * @trigger time-based (daily)
 */
function updateAgeAndSz() {
  const today = new Date();

  // ***** NEW: Check if today is January 1st *****
  // getMonth() returns 0 for January, 1 for February, etc.
  if (today.getMonth() !== 0 || today.getDate() !== 1) {
    // If it's not January 1st, log and exit silently.
    // This handles the daily trigger firing on non-Jan 1st days.
    Logger.log("updateAgeAndSz: Not January 1st (" + today.toLocaleDateString() + "). Silently exiting.");
    return;
  }
  // ********************************************

  logCall("updateAgeAndSz", {});
  Logger.log('updateAgeAndSz function called');

  const scriptProperties = PropertiesService.getScriptProperties();
  const lastRunDate = scriptProperties.getProperty("lastUpdateDate");
  const currentYear = today.getFullYear();
  Logger.log("Last run date = ", lastRunDate);
  // Check if already run this year
  if (lastRunDate) {
    const lastRun = new Date(lastRunDate);
    if (lastRun.getFullYear() === currentYear) {
      const message = `Age, Sz and Sz1 were updated on ${lastRun.toLocaleString()}. This can only be run once a year.`;
      Logger.log("updateAgeAndSz: " + message + " Skipping.");
      return;
    }
  }

  // Perform Update
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!sheet) {
    Logger.log("updateAgeAndSz: ERROR — 'Data' sheet not found. Aborting.");
    return;
  }
  const lastRow = sheet.getLastRow();

  for (let row = 5; row <= lastRow; row++) {
    const ageCell = sheet.getRange(row, 6); // Column F (Age)
    const szCell = sheet.getRange(row, 7); // Column G (Sz)
    const age = ageCell.getValue(); // Get the current age
    let sz = szCell.getValue(); // Get the current Sz

    if (age !== "N/A" && !isNaN(parseInt(age))) {
      // Increment age
      //ageCell.setValue(parseInt(age) + 1);

      // Update Sz based on the current Sz
      if (sz === "P" || sz === "W") {
        sz = incrementSz(sz);
      } else if (sz.startsWith("J") || sz.startsWith("S")) {
        sz = incrementSz(sz);
      } else if (sz.startsWith("A")) {
        const currentAdultStage = parseInt(sz.substring(1)) || 5; // Default adult stage to A5 if undefined
        sz = "A" + (currentAdultStage + 1);
      }

      szCell.setValue(sz); // Update Sz
    }
  }

  // Log successful execution
  scriptProperties.setProperty("lastUpdateDate", today.toISOString());
  Logger.log("Age, Sz, Sz1 Updated Successfully");

  // Log the activity
  try {
    const userEmail = "trigger/updateAgeAndSz";
    const details = "Age and Sz Updated.";
    const permId = "All";
    logActivity("EDIT", permId, details, userEmail);
    Logger.log('Activity logged successfully.');
  } catch (logError) {
    Logger.log('Error logging activity: ' + logError.toString());
    // We don't throw this error as it shouldn't stop the main function
  }
}

/**
 * Calculates next Sz stage.
 *
 * @param {string} currentSz
 * @returns {string}
 */
function incrementSz(currentSz) {
  logCall("incrementSz", { currentSz: currentSz });
  const szSequence = ["P", "W", "J1", "J2", "S3", "S4", "A5"]; // Define the Sz sequence
  const index = szSequence.indexOf(currentSz);

  if (index !== -1 && index < szSequence.length - 1) {
    return szSequence[index + 1]; // Return the next stage in the sequence
  } else if (currentSz.startsWith("A")) {
    const currentAdultStage = parseInt(currentSz.substring(1)) || 5;
    return "A" + (currentAdultStage + 1); // Increment the adult stage
  } else {
    return currentSz; // Return the current Sz if it's not in the sequence
  }
}

//---------------------------------- Log Activity Functions -----------------------------------------------------
/**
 * Logs an activity in the "ActivityLog" sheet of the active spreadsheet.
 * @param {string} action - The action performed.
 * @param {string} permId - The permission ID associated with the action.
 * @param {string} details - Additional details about the action.
 * @param {string} userEmail - The email of the user performing the action.
 * @returns {number} The row number where the log entry was added.
 */
function logActivity(action, permId, details, userEmail) {
  logCall("logActivity", { action: action, permId: permId, details: details, userEmail: userEmail});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("ActivityLog");
  var timestamp = new Date();

  // Create sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet("ActivityLog");
    logSheet.getRange(3, 1, 1, 7).setValues([["Timestamp", "User", "Action", "PermID", "Details", "Month", "Year"]]);
    logSheet.getRange(3, 1, 1, 7).setFontWeight("bold");
    logSheet.getRange(3, 1, 1, 7).setBackground("#c9daf8");
  }

  // Insert new row at top of data (row 5)
  logSheet.insertRowBefore(5);
  var nextRow = 5;

  // Add the new log entry (columns A-E)
  logSheet.getRange(nextRow, 1, 1, 5).setValues([[timestamp, userEmail, action, permId, details]]);

  // Add formulas for Month (F) and Year (G)
  logSheet.getRange(nextRow, 6).setFormulaR1C1('=IF(RC1="","",DATE(YEAR(RC1),MONTH(RC1),1))');
  logSheet.getRange(nextRow, 7).setFormulaR1C1('=IF(RC[-1]="","",YEAR(RC[-1]))');

  // Format timestamp
  logSheet.getRange(nextRow, 1).setNumberFormat("mmm dd, yyyy HH:mm:ss");

  // Set column widths
  logSheet.setColumnWidth(1, 175);
  logSheet.setColumnWidth(2, 290);
  logSheet.setColumnWidth(3, 70);
  logSheet.setColumnWidth(4, 150);
  logSheet.setColumnWidth(5, 610);
  logSheet.setColumnWidth(6, 40); // Month (hidden)
  logSheet.setColumnWidth(7, 40); // Year (hidden)

  // Add borders to the data table
  var lastRow = logSheet.getLastRow();
  logSheet.getRange(4, 1, lastRow - 3, 7).setBorder(true, true, true, true, false, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // -----------------------------
  // Update pivot table on same sheet at I3 (dynamic height)
  // -----------------------------
  // Clear previous pivot area
  logSheet.getRange("I3:Z" + logSheet.getMaxRows()).clear();

  var pivotRange = logSheet.getRange("I3");
  var pivotTable = pivotRange.createPivotTable(logSheet.getRange(4, 1, lastRow - 3, 7));

  // Rows: User → Year → Month
  var userGroup = pivotTable.addRowGroup(2); // User
  userGroup.showTotals(false);

  var yearGroup = pivotTable.addRowGroup(7); // Year
  yearGroup.showTotals(false);

  var monthGroup = pivotTable.addRowGroup(6); // Month

  // Columns: Action
  pivotTable.addColumnGroup(3);

  // Values: count of Action
  pivotTable.addPivotValue(3, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);

  // Show grand totals for columns
  pivotTable.setShowTotals(true);

  // Header row (I3:O4)
  var pivotHeaderRow = logSheet.getRange("I3:O3");
  pivotHeaderRow.setBackground("#c9daf8").setFontWeight("bold");

  // Rename headers
  var headerNames = ["User","Year","Month","ADD","DELETE","EDIT","Grand Total"];
  pivotHeaderRow.setValues([headerNames]);


  // -----------------------------
  // Automatically style Action columns (ADD/DELETE/EDIT)
  // -----------------------------
  var headerValues = pivotHeaderRow.getValues()[0];

  headerValues.forEach(function(header, i){
    var col = i + 9; // I = column 9
    if(header === "ADD"){
      logSheet.getRange(3, col, logSheet.getMaxRows() - 2).setFontWeight("bold").setFontColor("green");
    }
    if(header === "DELETE"){
      logSheet.getRange(3, col, logSheet.getMaxRows() - 2).setFontWeight("bold").setFontColor("red");
    }
    if(header === "EDIT"){
      logSheet.getRange(3, col, logSheet.getMaxRows() - 2).setFontWeight("bold").setFontColor("#ff6d01");
    }
  });

  return nextRow;
}



function reverseActivityLogOnce() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ActivityLog");
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 4) return;

  // Sort rows 4 through lastRow by column 1 (Timestamp) descending
  sheet.getRange(4, 1, lastRow - 3, 7)
       .sort({ column: 1, ascending: false });
}


//---------------------------------- Auto-Populate Drop Down Functions ------------------------
/**
 * Returns dropdown list from Lists sheet.
 *
 * @param {string} column
 * @returns {Array<string>}
 */
function GetDropDownArray(column) {
  logCall("GetDropDownArray", { column: column });
  const ss = SpreadsheetApp.getActiveSpreadsheet();  // Get the active spreadsheet
  const ws = ss.getSheetByName("Lists");  // Access the "Lists" sheet
  
  // Get all the values from the specified column, starting from row 2 (excluding header)
  const data = ws.getRange(column + "2:" + column).getValues();
  let lastRow = 0;
  
  // Loop through the column data from the bottom to find the last non-empty row
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] !== '') {  // If a non-empty cell is found
      lastRow = i + 2;  // Set the last row number (add 2 since data starts at row 2)
      break;
    }
  }
  
  // Get the range of non-empty rows in the column and return the values as a 1D array
  const columnData = ws.getRange(column + "2:" + column + lastRow).getValues();
  return columnData.map(row => row[0]);  // Extract the first value from each row
}

//---------------------------------- Add Record Functions -----------------------------------------
/**
 * Checks if a given permId already exists in the 'Data' sheet.
 * This function prevents duplicates by scanning the sheet for matching permId values.
 * 
 * @param {string} permId - The permId to check for duplication.
 * @returns {boolean} - Returns true if the permId already exists, false otherwise.
 */
function duplicatePermidCheck(permId) {
  logCall("duplicatePermidCheck", { permId: permId });
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = sheet.getDataRange().getValues();
  
  // Assuming PermID is in column D (index 3)
  var permIdColumnIndex = 3;
  
  // Loop through the data, starting from the second row to skip the header
  for (var i = 1; i < data.length; i++) {
    var cellPermId = data[i][permIdColumnIndex].toString().trim().toLowerCase();
    
    // Check if the permId matches the input, case-insensitive
    if (cellPermId === permId.trim().toLowerCase()) {
      return true;  // Exact match found
    }
  }  
  return false;  // No match found
}

/**
 * Opens the 'EditRecord' page in a modal dialog and passes the permId to the page.
 * The permId is stored in PropertiesService to be accessible within the modal.
 * 
 * @param {string} permId - The permId to be edited in the form.
 */
function openEditRecordPage(permId) {
  logCall("openEditRecordPage", { permId: permId });
  var html = HtmlService.createHtmlOutputFromFile('EditRecord')
      .setWidth(1060)
      .setHeight(560);
  
  // Display the modal dialog for editing records
  SpreadsheetApp.getUi().showModalDialog(html, 'Edit Record');
  
  // Store the permId and group in script properties for access in EditRecord.html
  PropertiesService.getScriptProperties().setProperty('editPermId', permId);
}

/**
 * Finds last non-empty PermID row.
 *
 * @param {Sheet} sheet
 * @returns {number}
 */
function getLastDataRowInColumnD(sheet) {
  const colD = sheet.getRange("D:D").getValues();
  for (let i = colD.length - 1; i >= 0; i--) {
    if (colD[i][0] !== "" && colD[i][0] !== null) {
      return i + 1; // Rows are 1-based
    }
  }
  return 1; // Default to first row if nothing found
}

/**
 * Primary CREATE operation.
 *
 * Responsibilities:
 *  - Appends new row.
 *  - Injects formulas.
 *  - Applies formatting.
 *  - Logs activity.
 *
 * @param {Array} formData
 * @returns {{lastRow:number}}
 */
function appendRow(formData) {
  logCall("appendRow", { formData: formData });
  Logger.log('Received Form Data: ' + JSON.stringify(formData));
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  // Check if formData is correctly received
  if (!formData || !Array.isArray(formData)) {
    throw new Error('Invalid formData received');
  }

  // Create an array to store the field names
  var fieldNames = [
    'island', 'group', 'permid', 'sz', 'sx', 'tagged', 'lefttag', 'ltresight', 'ltconddorsal', 'ltcondventral',
    'righttag', 'rtresight', 'rtconddorsal', 'rtcondventral', 'comment',
    '1number', '1mark', '1location1', '1location1a', '1site',
    '2number', '2mark', '2location1', '2location1a', '2site',
    '3number', '3mark', '3location1', '3location1a', '3site',
    '4number', '4mark', '4location1', '4location1a', '4site',
    '5number', '5mark', '5location1', '5location1a', '5site',
    '6number', '6mark', '6location1', '6location1a', '6site',
    '7number', '7mark', '7location1', '7location1a', '7site'
  ];

  // Create an array to store the added fields
  var addedFields = [];

  // Assuming columns C, E, F, S are calculated and should not be overwritten
  var row = [
    formData[0], // A: 'island'
    formData[1], // B: 'group'
    "",          // C: year (leave blank for formula)
    formData[2], // D: 'permid'
    "",          // E: sz1 (leave blank for formula)
    "",          // F: age (leave blank for formula)
    formData[3], // G: 'sz'
    formData[4], // H: 'sx'
    formData[5], // I: 'tagged'
    formData[6], // J: 'lefttag'
    formData[7], // K: 'ltresight'
    formData[8], // L: 'ltconddorsal'
    formData[9], // M: 'ltcondventral'
    formData[10], // N: 'righttag'
    formData[11], // O: 'rtresight'
    formData[12], // P: 'rtconddorsal'
    formData[13], // Q: 'rtcondventral'
    formData[14], // R: 'comment'
    "",          // S: identifying marks (leave blank for formula)
    formData[15], // T: '1number'
    formData[16], // U: '1mark'
    formData[17], // V: '1location1'
    formData[18], // W: '1location1a'
    formData[19], // X: '1site'
    formData[20], // Y: '2number'
    formData[21], // Z: '2mark'
    formData[22], // AA: '2location1'
    formData[23], // AB: '2location1a'
    formData[24], // AC: '2site'
    formData[25], // AD: '3number'
    formData[26], // AE: '3mark'
    formData[27], // AF: '3location1'
    formData[28], // AG: '3location1a'
    formData[29], // AH: '3site'
    formData[30], // AI: '4number'
    formData[31], // AJ: '4mark'
    formData[32], // AK: '4location1'
    formData[33], // AL: '4location1a'
    formData[34], // AM: '4site'
    formData[35], // AN: '5number'
    formData[36], // AO: '5mark'
    formData[37], // AP: '5location1'
    formData[38], // AQ: '5location1a'
    formData[39], // AR: '5site'
    formData[40], // AS: '6number'
    formData[41], // AT: '6mark',
    formData[42], // AU: '6location1'
    formData[43], // AV: '6location1a'
    formData[44], // AW: '6site'
    formData[45], // AX: '7number'
    formData[46], // AY: '7mark'
    formData[47], // AZ: '7location1'
    formData[48], // BA: '7location1a'
    formData[49]  // BB: '7site'
  ];

  // Append the row to the sheet
  const lastRowD = getLastDataRowInColumnD(ws);
  const nextRow = lastRowD + 1;
  try {
    Logger.log('lastRowD: ' + lastRowD);
    Logger.log('appendRow fxn nextRow: ' + nextRow);
    ws.getRange(nextRow, 1, 1, row.length).setValues([row]);
    Logger.log('Row successfully added');
  } catch (e) {
    Logger.log('Error adding row: ' + e.toString());
    throw new Error('Error adding row: ' + e.toString());
  }

  // Define column indexes for better readability
  const yearCol = 3;  // Column C
  const sz1Col = 5;   // Column E
  const ageCol = 6;   // Column F
  const marksCol = 19; // Column S

  // IMPORTANT: Use relative references in formulas (e.g., G instead of G' + nextRow + ')
  // This allows formulas to auto-adjust if rows are inserted/moved.
  const yearFormula = '=IF(G="NA", "", IF(OR(G="P", G="W"), DATE(YEAR(TODAY()), 1, 1), IF(AND(LEN(G)<2, G<>"P", G<>"W"), "", IF(LEN(G)<2, "", DATE(YEAR(TODAY())-VALUE(RIGHT(G, LEN(G)-1)), 1, 1))))))';
  const sz1Formula = '=IF(AND(F="N/A", G=""), "", IF(AND(F="N/A", G="A"), "Adult", IF(AND(F="N/A", G="J"), "Juvenile", IF(INT(F)>4, "Adult", IF(INT(F)>2, "Subadult", IF(INT(F)>=1, "Juvenile", IF(G="W", "Weanling", IF(G="P", "Pup", IF(G="NA", "N/A", IF(G="", "", IF(LEN(G)<2, G))))))))))))';
  const ageFormula = '=IF(G="NA", "N/A", IF(AND(C="", LEN(G)<2), "N/A", IF(AND(ISBLANK(C), LEN(G)>=2), VALUE(RIGHT(G, LEN(G)-1)), INT(YEARFRAC(C, TODAY(), 1)))))';
  const marksFormula = '=SUBSTITUTE(TEXTJOIN(" ", TRUE, T:BB), ",", CHAR(10))';

  // Set dynamic formulas using relative references with error handling
  // Removed the redundant individual setFormula calls and consolidated with the formulas array.
  try {
    ws.getRange(nextRow, yearCol).setFormula(yearFormula);
    Logger.log("Year formula set for row " + nextRow + ": " + yearFormula);
  } catch (e) {
    Logger.log('Error setting formula for yearCol (Column ' + yearCol + ') on row ' + nextRow + ': ' + e.toString());
  }

  try {
    ws.getRange(nextRow, sz1Col).setFormula(sz1Formula);
    Logger.log("Sz1 formula set for row " + nextRow + ": " + sz1Formula);
  } catch (e) {
    Logger.log('Error setting formula for sz1Col (Column ' + sz1Col + ') on row ' + nextRow + ': ' + e.toString());
  }

  try {
    ws.getRange(nextRow, ageCol).setFormula(ageFormula);
    Logger.log("Age formula set for row " + nextRow + ": " + ageFormula);
  } catch (e) {
    Logger.log('Error setting formula for ageCol (Column ' + ageCol + ') on row ' + nextRow + ': ' + e.toString());
  }

  try {
    ws.getRange(nextRow, marksCol).setFormula(marksFormula);
    Logger.log("Markings formula set for row " + nextRow + ": " + marksFormula);
  } catch (e) {
    Logger.log('Error setting formula for marksCol (Column ' + marksCol + ') on row ' + nextRow + ': ' + e.toString());
  }

  // Ensure all changes are applied after setting formulas
  SpreadsheetApp.flush();

  // Set data formats for the columns
  ws.getRange(nextRow, 1, 1, 2).setNumberFormat('@STRING@'); // A:B as plain text
  ws.getRange(nextRow, 3).setNumberFormat('m/d/yyyy'); // C as date
  ws.getRange(nextRow, 4, 1, ws.getLastColumn() - 3).setNumberFormat('@STRING@'); // D:BB as plain text

  // Populate the addedFields array only with non-empty fields
  for (var i = 0; i < formData.length; i++) {
    if (formData[i] && formData[i].toString().trim() !== "") {
      addedFields.push(fieldNames[i].toUpperCase() + ": " + formData[i]);
    }
  }

  Logger.log('Formulas set successfully for row: ' + nextRow);

  // Log the activity
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var permId = formData[2]; // Assuming permId is the third element in formData
    var details = "New record added. Fields: " + addedFields.join(", ");
    logActivity("ADD", permId, details, userEmail);
    Logger.log('Activity logged successfully.');
  } catch (logError) {
    Logger.log('Error logging activity: ' + logError.toString());
    // We don't throw this error as it shouldn't stop the main function
  }

  return { lastRow: nextRow };
}

//-----------------------------------Data Sheet Search Record Functions --------------------------------------------------
/**
 * Handles the action when a match is selected from the list of matches.
 * Sets the focus on the corresponding row number in the Google Sheet.
 *
 * Outcomes:
 *  - Single match → focus row.
 *  - Multiple matches → modal dialog.
 *  - No matches → alert.
 * @returns {void}
 * @param {number} index - The index of the selected match in the matches array.
 */

function handleSelectedMatch(index) {
  logCall("handleSelectedMatch", { index: index });
  var matches = getMatches();
  var selectedMatch = matches[index];
  setFocusAndHighlightLater(selectedMatch.rowNumber);
}

/**
 * Displays a modal dialog showing multiple matches found for a search term.
 * The dialog's height is adjusted based on the number of matches.
 *
 * @param {Array<Object>} matches - An array of match objects containing PermID and row information.
 */
function showSheetMultipleMatchesDialog(matches) {
  logCall("showSheetMultipleMatchesDialog", { matches: matches });
  var html = HtmlService.createHtmlOutputFromFile('MatchesModal')
      .setWidth(400)
      .setHeight(200 + matches.length * 40); // Adjust height based on number of matches
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Matching Records');
  
  // Store matches in PropertiesService for later retrieval
  PropertiesService.getUserProperties().setProperty('matches', JSON.stringify(matches));
}

/**
 * Retrieves the matches stored in user properties (per-user, not shared).
 *
 * @returns {Array<Object>} - An array of match objects parsed from JSON.
 */
function getMatches() {
  logCall("getmatches", {});
  var matchesJson = PropertiesService.getUserProperties().getProperty('matches');
  return JSON.parse(matchesJson);
}

/**
 * Main function triggered by the search button.
 * Reads the search box, validates input, and delegates
 * to sheetSearchRecord().
 */
function sheetSearchButton() {
  logCall("sheetSearchButton", {});
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Data');

  Logger.log("Starting searchButton function");

  var range = sheet.getRange('C2');
  var rawValue = range.getValue();
  Logger.log("Raw value from C2: " + rawValue);

  var permId = rawValue.toString().trim();
  Logger.log("Processed permId: " + permId);

  if (permId === "") {
    Logger.log("permId is empty");
    SpreadsheetApp.getUi().alert("No record found. Please enter a valid PermID in cell C2.");
    return;
  }

  Logger.log("Calling SearchRecord with permId: " + permId);
  var result = sheetSearchRecord(permId);

  if (result.singleMatch) {
    Logger.log("Single match found. lastRowNumber: " + result.data.rowNumber);
    setFocusAndHighlightLater(result.data.rowNumber);
  } else if (result.multipleMatches) {
    Logger.log("Multiple matches found");
    showSheetMultipleMatchesDialog(result.matches);
  } else {
    Logger.log("No match found");
    SpreadsheetApp.getUi().alert('No matching record found for search term: ' + permId);
  }
}

/**
 * Searches the Data sheet for rows matching a PermID substring.
 *
 * @param {string} permId - Search token (case-insensitive).
 * @returns {Object} Search result:
 *   {
 *     singleMatch: boolean,
 *     multipleMatches: boolean,
 *     data?: Object,
 *     matches?: Array<Object>
 *   }
 */
function sheetSearchRecord(permId) {
  logCall("sheetSearchRecord", {permId: permId });
  var lowerPermId = permId.toLowerCase();
  Logger.log("Searching for PermID in sheet: " + lowerPermId);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = sheet.getDataRange().getValues();
  
  var matches = [];
  
  for (var i = 4; i < data.length; i++) { // Start from row 5 to skip header rows
    var row = data[i];
    var cellPermId = row[3].toString().toLowerCase(); // Assuming PermID is in column D (index 3)
    Logger.log("Checking row " + (i + 1) + " with PermID: " + cellPermId);
    
    if (cellPermId.indexOf(lowerPermId) !== -1) {
      Logger.log("Match found in row " + (i + 1));
      matches.push({permid: row[3], row: row, rowNumber: i + 1});
    }
  }
  
  if (matches.length === 1) {
    return {singleMatch: true, data: matches[0]};
  } else if (matches.length > 1) {
    return {multipleMatches: true, matches: matches};
  }
  
  Logger.log("No match found");
  return {singleMatch: false, multipleMatches: false, error: "No matching records found."};
}

//---------------------------------- Edit Record Functions -----------------------------------------------------

// Declare a global variable to store the row number
var lastRowNumber = 0;

/**
 * Retrieves the edit PermID stored in script properties.
 * Clears the property after retrieval to ensure it is only used once.
 *
 * @returns {string|null} - The edit PermID, or null if not found.
 */
function getEditPermId() {
  logCall("getEditPermId", {});
  var permId = PropertiesService.getScriptProperties().getProperty('editPermId');
  PropertiesService.getScriptProperties().deleteProperty('editPermId'); // Clear the property after use  
  return permId;
}

/**
 * Searches for a given PermID in the 'Data' sheet.
 * Returns the row data if a single match is found, or multiple matches if more than one is found.
 *
 * @param {string} permId - The PermID to search for.
 * @returns {Array<Object>|Object|null} - Row data if a single match is found, an object containing multiple matches, or null if no match is found.
 */
function SearchRecord(permId) {
  logCall("SearchRecord", { permId: permId });
  var lowerPermId = permId.toLowerCase();
  Logger.log("Searching for PermID: " + lowerPermId);
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var data = sheet.getDataRange().getValues();
  
  var matches = [];
  
  for (var i = 1; i < data.length; i++) { // Start from 1 to skip header row
    var row = data[i];
    var cellPermId = row[3].toString().toLowerCase(); // Assuming PermID is in column D (index 3)
    Logger.log("Checking row " + (i + 1) + " with PermID: " + cellPermId);
    
    if (cellPermId.indexOf(lowerPermId) !== -1) {
      Logger.log("Match found in row " + (i + 1));
      matches.push({row: row, rowNumber: i + 1});
    }
  }
  
  if (matches.length === 1) {
    lastRowNumber = matches[0].rowNumber;
    return matches[0].row;
  } else if (matches.length > 1) {
    return {multipleMatches: true, matches: matches};
  }
  Logger.log("No match found");
  return null;
}

/*
// Use for testing - call from EditRecord.html  testServerLogging()
// Fxn commented out until needed
function testServerLogging() {
  console.log('Server: Test log from testServerLogging function');
  return 'Logging test successful';
}
*/

/**
 * Retrieves and formats the record data for a specific PermID.
 * Returns structured data for easy access or indicates if no matches are found.
 *
 * @param {string} permId - The PermID for which to retrieve data.
 * @returns {Object} - An object containing structured record data or an error message.
 */
function GetRecordData(permId) {
  logCall("GetRecordData", { permId: permId });
  console.log('Server: GetRecordData function called with permId:', permId);
  Logger.log("Getting record data for PermID: " + permId);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    var lowerPermId = permId.toLowerCase();
    var matches = [];

    for (var i = 1; i < values.length; i++) { // Start from 1 to skip header row
      var row = values[i];
      var cellPermId = row[3].toString().toLowerCase(); // Assuming PermID is in column D (index 3)
      
      if (cellPermId.indexOf(lowerPermId) !== -1) {
        matches.push({row: row, rowNumber: i + 1});
      }
    }

    if (matches.length === 1) {
      console.log('Server: Single match found, formatting data');
      var formattedData = formatRecordData(matches[0].row, matches[0].rowNumber);
      return {
        singleMatch: true,
        data: formattedData
      };
    } else if (matches.length > 1) {
      console.log('Server: Multiple matches found, formatting data');
      var formattedMatches = matches.map(match => formatRecordData(match.row, match.rowNumber));
      return {
        multipleMatches: true,
        matches: formattedMatches
      };
    } else {
      console.log('Server: No matches found');
      return { error: "No matching records found" };
    }
  } catch (error) {
    console.error('Server: Error in GetRecordData:', error);
    Logger.log("Error in GetRecordData: " + error.message);
    return { error: "An error occurred while processing the request" };
  } finally {
    console.log('Server: GetRecordData function completed');
  }
}

/**
 * Formats the record data into a structured object.
 *
 * @param {Array} record - The row data from the sheet.
 * @param {number} rowNumber - The row number of the record in the sheet.
 * @returns {Object} - An object containing structured record data.
 */
function formatRecordData(record, rowNumber) {
  logCall("formatRecordData", { record: record, rowNumber: rowNumber });
  return {
    rowNumber: rowNumber,
    island: record[0],
    group: record[1],
    permid: record[3],
    sz: record[6],
    sx: record[7],
    tagged: record[8],
    lefttag: record[9],
    ltresight: record[10],
    ltconddorsal: record[11],
    ltcondventral: record[12],
    righttag: record[13],
    rtresight: record[14],
    rtconddorsal: record[15],
    rtcondventral: record[16],
    comment: record[17],
    number1: record[19],
    mark1: record[20],
    location1_1: record[21],
    location1a_1: record[22],
    site1: record[23],
    number2: record[24],
    mark2: record[25],
    location1_2: record[26],
    location1a_2: record[27],
    site2: record[28],
    number3: record[29],
    mark3: record[30],
    location1_3: record[31],
    location1a_3: record[32],
    site3: record[33],
    number4: record[34],
    mark4: record[35],
    location1_4: record[36],
    location1a_4: record[37],
    site4: record[38],
    number5: record[39],
    mark5: record[40],
    location1_5: record[41],
    location1a_5: record[42],
    site5: record[43],
    number6: record[44],
    mark6: record[45],
    location1_6: record[46],
    location1a_6: record[47],
    site6: record[48],
    number7: record[49],
    mark7: record[50],
    location1_7: record[51],
    location1a_7: record[52],
    site7: record[53],
    dead: record[57],
  };
}

/**
 * Edits a record in the 'Data' sheet based on the provided form data.
 *
 * @param {Object} formData - The form data containing updated values.
 * @param {number} rowNumber - The row number of the record to be edited.
 * @returns {number} - The row number of the edited record.
 * @throws {Error} - Throws an error if the row number is invalid.
 */
function EditRecord(formData, rowNumber) {
  logCall("EditRecord", { formData: formData, rowNumber: rowNumber });
  Logger.log('EditRecord(formData) received formData: ' + JSON.stringify(formData));
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var rowNumber = formData.rowNumber;

  // Checks for a valid row number
  Logger.log("rowNumber = formData.rowNumber = " + rowNumber);
  if (!rowNumber || rowNumber <= 0) {
    throw new Error("Invalid row number: " + rowNumber);
  }

  // Retrieve the current data in the row
  var currentRowData = ws.getRange(rowNumber, 1, 1, ws.getLastColumn()).getValues()[0];

  // Current permid and group from pre-edit form
  var currentPermId = formData.currentPermId;
  var preEditGroup = formData.preEditGroup;
    
  // editPermId from post-edit form data 
  var editPermId = currentRowData[3];
  Logger.log('Edit PermId (from form): ' + editPermId);

  // Define the updated row data, split into parts
  var updatedRowPart1 = currentRowData.slice(0, 2);   // Columns A-B
  var updatedRowPart2 = currentRowData.slice(3, 4);   // Column D
  var updatedRowPart3 = currentRowData.slice(6, 18);  // Columns G-R
  var updatedRowPart4 = currentRowData.slice(19);     // Columns T-BB

  // Initialize an array to store the changes
  var changes = [];

  // Update fields for each part
  if (formData.island && formData.island !== currentRowData[0]) {
    updatedRowPart1[0] = formData.island;  // A: island
    changes.push("ISLAND: " + formData.island);
  }
  if (formData.group && formData.group !== currentRowData[1]) {
    updatedRowPart1[1] = formData.group;  // B: group
    changes.push("GROUP changed from " + preEditGroup + ' to ' + formData.group);
  }   

  if (formData.permid && formData.permid !== currentRowData[3]) {
    updatedRowPart2[0] = formData.permid;  // D: permid
    changes.push("PERMID changed from " + currentPermId + ' to ' + formData.permid);
  }

  // Update fields for columns G-R
  const fieldsToUpdate = ['sz', 'sx', 'tagged', 'lefttag', 'ltresight', 'ltconddorsal', 'ltcondventral', 
                          'righttag', 'rtresight', 'rtconddorsal', 'rtcondventral', 'comment'];
  fieldsToUpdate.forEach((field, index) => {
    if (formData[field] && formData[field] !== currentRowData[index + 6]) {
      updatedRowPart3[index] = formData[field];
      changes.push(field.toUpperCase() + ": " + formData[field]);
    }
  });

  // Update marks, locations, and site fields (only if changed)
  const updateFields = (fieldPrefix, startIndex) => {
    for (let i = 1; i <= 7; i++) {
      const field = fieldPrefix + i;
      const colIndex = startIndex + (i - 1) * 5;
      if (formData[field] !== undefined && formData[field] !== currentRowData[colIndex]) {
        updatedRowPart4[colIndex - 19] = formData[field];
        changes.push(field.toUpperCase() + ": " + formData[field]);
      }
    }
  };

  updateFields('number', 19);
  updateFields('mark', 20);
  updateFields('location1_', 21);
  updateFields('location1a_', 22);
  updateFields('site', 23);

  // Column BF is index 57 (0-indexed in currentRowData)
  // Calculate its index relative to updatedRowPart4
  const bfSliceIndex = 57 - 19; // 57 (BF index) - 19 (start of updatedRowPart4) = 38

  if (formData.dead !== undefined && formData.dead !== currentRowData[57]) {
    updatedRowPart4[bfSliceIndex] = formData.dead; // Update the slice for BF
    changes.push("DEAD: " + formData.dead);
  }

  // Update the sheet with the new data, preserving the calculated columns C, E, F, and S
  ws.getRange(rowNumber, 1, 1, 2).setValues([updatedRowPart1]);
  ws.getRange(rowNumber, 4, 1, 1).setValues([updatedRowPart2]);
  ws.getRange(rowNumber, 7, 1, 12).setValues([updatedRowPart3]);
  ws.getRange(rowNumber, 20, 1, updatedRowPart4.length).setValues([updatedRowPart4]);

  // Log the activity
  try {
    var userEmail = Session.getActiveUser().getEmail();
    var permId = formData.permid;
    var details = "Changes: " + changes.join(", ");
    logActivity("EDIT", permId, details, userEmail);
    Logger.log('Activity logged successfully.');
  } catch (logError) {
    Logger.log('Error logging activity: ' + logError.toString());
    // We don't throw this error as it shouldn't stop the main function
  }

  return rowNumber;
}

//----------------- Set Focus / Position New Record / Highlight Functions ---------------------------------
/**
 * Focuses row and applies business highlighting rules.
 *
 * @param {number} rowNumber
 * @param {boolean} isNewRecord
 */
function setFocusAndHighlightLater(rowNumber, isNewRecord) {
  Logger.log("setFocusAndHighlightLater called with rowNumber: " + rowNumber + ", isNewRecord: " + isNewRecord);
  logCall("setFocusAndHighlightLater", { rowNumber: rowNumber, isNewRecord: isNewRecord });  

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');

  if (isNewRecord) {
    Logger.log("New record detected. Attempting to insert in order.");
    rowNumber = insertNewRecordInOrder(sheet, rowNumber);
    Logger.log("After insertion, new rowNumber: " + rowNumber);
  }

  // --- Highlight Logic for 'dead' records (Column BF) and Column BD management ---
const bdColumnIndex = 56; // Column BD (1-indexed)
const bfColumnIndex = 58; // Column BF (1-indexed)
const highlightRangeStartCol = 1; // Column A
const highlightRangeEndCol = 19;  // Column S

const lightGrayColor = "#d3d3d3"; // For 'dead' status
const colBColor = "#feffb3";      // Light yellow for Column B
const colJNColor = "#b7e1cd";     // Light green for Columns J and N
const colIColor = "#ffffff";      // White font color Column I 

try {
  const deadStatusValue = sheet.getRange(rowNumber, bfColumnIndex).getValue();
  Logger.log("Checking row " + rowNumber + ", Column BF value: '" + deadStatusValue + "'");

  const highlightRange = sheet.getRange(rowNumber, highlightRangeStartCol, 1, highlightRangeEndCol); // A:S

  const isDead = typeof deadStatusValue === 'string' && deadStatusValue.trim().toLowerCase() === 'dead';

  if (isDead) { 
    Logger.log("Column BF contains 'dead'. Highlighting A:S light gray, clearing BD, and setting text color in I.");

    // Set background for A:S
    highlightRange.setBackground(lightGrayColor);

    // Clear column BD content
    sheet.getRange(rowNumber, bdColumnIndex).clearContent();

    // Set text color in column I (column index 9)
    sheet.getRange(rowNumber, 9).setFontColor(lightGrayColor);
  } else {
    Logger.log("Column BF is not 'dead'. Applying column colors and setting formula in BD.");

    // Set background colors A:S
    const backgroundColors = new Array(highlightRangeEndCol - highlightRangeStartCol + 1).fill(null);
    backgroundColors[1] = colBColor;    // Column B
    backgroundColors[9] = colJNColor;   // Column J
    backgroundColors[13] = colJNColor;  // Column N
    highlightRange.setBackgrounds([backgroundColors]);
    sheet.getRange(rowNumber, 9).setFontColor(colIColor);

  }

} catch (e) {
  Logger.log("Error during highlighting/BD logic for row " + rowNumber + ": " + e.toString());
}

  // Set focus to the cell in column A
  sheet.setActiveRange(sheet.getRange(rowNumber, 4));
  
  // Make column D bold for this row
  var columnDCell = sheet.getRange(rowNumber, 4);
  columnDCell.setFontWeight("bold");

  populateCombinedColumnFormulaForRow(rowNumber); // Adds concatenated columns to BD
}

/**
 * Schedules background reset trigger.
 */
function scheduleReset() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = triggers.some(function(trigger) {
    return trigger.getHandlerFunction() == "checkAndResetBackground";
  });

  if (!triggerExists) {
    ScriptApp.newTrigger("checkAndResetBackground")
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

/**
 * Inserts a newly added record into its correct sorted position.
 *
 * This function:
 *  - Assumes the record was initially appended at the bottom.
 *  - Compares the new PermID with existing PermIDs.
 *  - Moves the row upward until alphabetical order is preserved.
 *
 * Why this exists:
 *  - Users expect Data sheet to remain ordered.
 *  - New inserts must not break visual navigation.
 *
 * How it works:
 *  1. Reads the PermID of the new row.
 *  2. Scans upward through column D.
 *  3. Finds the first row where newPermID < existingPermID.
 *  4. Moves the new row to that index.
 *
 * Architectural role:
 *  - Post-processing step of CREATE pipeline.
 *  - Guarantees Data sheet ordering invariant.
 *
 * @param {Sheet} sheet - Data sheet.
 * @param {number} rowNumber - Row where record was appended.
 * @returns {number} Final row index after reordering.
 */
function insertNewRecordInOrder(sheet, rowNumber) {
  Logger.log("insertNewRecordInOrder called with rowNumber: " + rowNumber);
  logCall("insertNewRecordInOrder", { sheet: sheet, rowNumber: rowNumber});

  // Define the standard relative formulas for C, E, F, S (consistent with appendRow)
  // These use INDIRECT("Column"&ROW()) for all relative column references
  const yearColFormula = '=IF(INDIRECT("G"&ROW())="NA", "", IF(OR(INDIRECT("G"&ROW())="P", INDIRECT("G"&ROW())="W"), DATE(YEAR(TODAY()), 1, 1), IF(AND(LEN(INDIRECT("G"&ROW()))<2, INDIRECT("G"&ROW())<>"P", INDIRECT("G"&ROW())<>"W"), "", IF(LEN(INDIRECT("G"&ROW()))<2, "", DATE(YEAR(TODAY())-VALUE(RIGHT(INDIRECT("G"&ROW()), LEN(INDIRECT("G"&ROW()))-1)), 1, 1)))))';

  // *** UPDATED SZ1 FORMULA: Added ISNUMBER check to prevent #VALUE! error ***
  const sz1ColFormula =
    '=IF(INDIRECT("G"&ROW())="NA", "N/A",' + // Priority 1: If G is "NA", sz1 is "N/A"
    'IF(AND(INDIRECT("F"&ROW())="N/A", INDIRECT("G"&ROW())=""), "",' + // Priority 2: If F is "N/A" and G is blank
    'IF(AND(INDIRECT("F"&ROW())="N/A", INDIRECT("G"&ROW())="A"), "Adult",' + // Priority 3: If F is "N/A" and G is "A"
    'IF(AND(INDIRECT("F"&ROW())="N/A", INDIRECT("G"&ROW())="J"), "Juvenile",' + // Priority 4: If F is "N/A" and G is "J"
    'IF(ISNUMBER(INDIRECT("F"&ROW())),' + // Priority 5: If F (Age) IS A NUMBER, proceed with age-based classification
    'IF(INT(INDIRECT("F"&ROW()))>4, "Adult",' +
    'IF(INT(INDIRECT("F"&ROW()))>2, "Subadult",' +
    'IF(INT(INDIRECT("F"&ROW()))>=1, "Juvenile",' +
    'IF(INDIRECT("G"&ROW())="W", "Weanling",' + // Fallback if F is numeric but no age-stage match, check G
    'IF(INDIRECT("G"&ROW())="P", "Pup",' +
    'IF(LEN(INDIRECT("G"&ROW()))<2, INDIRECT("G"&ROW()),' +
    '""' + // Final fallback for numeric F
    ')' +
    ')' +
    ')' +
    ')' +
    ')' +
    '),' + // End ISNUMBER(F) TRUE branch
    'IF(INDIRECT("G"&ROW())="W", "Weanling",' + // Priority 6: If F (Age) is NOT a number (e.g., still "N/A" but G is not "", "A", or "J"), use G-based codes
    'IF(INDIRECT("G"&ROW())="P", "Pup",' +
    'IF(LEN(INDIRECT("G"&ROW()))<2, INDIRECT("G"&ROW()),' +
    '""' + // Final fallback for non-numeric F
    ')' +
    ')' +
    ')' +
    ')' + // End ISNUMBER(F) FALSE branch
    ')' + // End F="N/A" G="J" IF
    ')' + // End F="N/A" G="A" IF
    ')' + // End F="N/A" G="" IF
    ')'; // End G="NA" IF


  const ageColFormula = '=IF(INDIRECT("G"&ROW())="NA", "N/A", IF(AND(INDIRECT("C"&ROW())="", LEN(INDIRECT("G"&ROW()))<2), "N/A", IF(AND(ISBLANK(INDIRECT("C"&ROW())), LEN(INDIRECT("G"&ROW()))>=2), VALUE(RIGHT(INDIRECT("G"&ROW()), LEN(INDIRECT("G"&ROW()))-1)), INT(YEARFRAC(INDIRECT("C"&ROW()), TODAY(), 1)))))';

  // Marks formula already uses INDIRECT, so it should be fine.
  const marksColFormula = '=SUBSTITUTE(TEXTJOIN(" ", TRUE, INDIRECT("T"&ROW()):INDIRECT("BB"&ROW())), ",", CHAR(10))';

  // Define the formulas that should go into BJ:BM.
  const bjFormula = yearColFormula; // Column BJ (62)
  const bkFormula = sz1ColFormula;  // Column BK (63)
  const blFormula = ageColFormula;   // Column BL (64)
  const bmFormula = marksColFormula; // Column BM (65)

  const bjbmFormulasToSet = [[
      bjFormula,
      bkFormula,
      blFormula,
      bmFormula
  ]];

  // Find the last row with data in column A
  var lastRowWithData = sheet.getRange("A:A").getValues().filter(String).length;
  Logger.log("Last row with data in column A: " + lastRowWithData);

  var data = sheet.getRange("A5:BB" + lastRowWithData).getValues();

  // Get values for the row to be moved (NOTE: getValues() retrieves calculated results, not formulas)
  var editedRowData = sheet.getRange(rowNumber, 1, 1, 54).getValues()[0];
  Logger.log("Edited record data: " + JSON.stringify(editedRowData));

  // Remove the current row
  sheet.deleteRow(rowNumber);

  // Re-fetch the data after deletion
  data = sheet.getRange("A5:BB" + (lastRowWithData - 1)).getValues();

  // Find the correct insertion index using editedRowData
  var insertIndex = findInsertIndexExcludingRow(data, editedRowData, rowNumber);
  Logger.log("Initial insert index found: " + insertIndex);

  // Ensure insertIndex is valid
  if (insertIndex >= lastRowWithData) {
    insertIndex = lastRowWithData;
  }

  Logger.log('Final insertIndex: ' + insertIndex);
  Logger.log('rowNumber: ' + rowNumber);

  // Insert a new row at the desired position
  if (insertIndex !== rowNumber) {
    sheet.insertRowBefore(insertIndex);
    var newRowRange = sheet.getRange(insertIndex, 1, 1, 54);

    // Set the values for the new row (this will overwrite any formulas in columns C, E, F, S)
    newRowRange.setValues([editedRowData]);

    // Apply formulas to BJ:BM in the newly inserted row from script templates.
    try {
        sheet.getRange(insertIndex, 62, 1, 4).setFormulas(bjbmFormulasToSet);
        Logger.log("Formulas set to BJ:BM in new row " + insertIndex + " from script templates.");
    } catch (e) {
        Logger.log('Error setting BJ:BM formulas from script templates: ' + e.toString());
    }

    // Re-apply formulas to columns C, E, F, and S in the new position (as they were overwritten by setValues)
    try {
      sheet.getRange(insertIndex, 3).setFormula(yearColFormula);   // Column C
      sheet.getRange(insertIndex, 5).setFormula(sz1ColFormula);    // Column E
      sheet.getRange(insertIndex, 6).setFormula(ageColFormula);    // Column F
      sheet.getRange(insertIndex, 19).setFormula(marksColFormula); // Column S (using corrected formula)
      Logger.log("Formulas for C,E,F,S re-applied to new row position " + insertIndex);
    } catch (e) {
      Logger.log('Error re-applying formulas in insertNewRecordInOrder (moved row): ' + e.toString());
    }

    Logger.log("Row moved from " + rowNumber + " to position " + (insertIndex));
    return insertIndex;

  } else { // No need to move row. It's already in the correct position.

    Logger.log("No need to move row. It's already in the correct position.");
    var newRowRange = sheet.getRange(insertIndex, 1, 1, 54);

    // Set the values for the new row (this will overwrite formulas)
    newRowRange.setValues([editedRowData]);

    // Apply formulas to BJ:BM even if the row didn't move (ensures they are always there if overwritten by edit)
    try {
        sheet.getRange(insertIndex, 62, 1, 4).setFormulas(bjbmFormulasToSet);
        Logger.log("Formulas set to BJ:BM in stationary row " + insertIndex + " from script templates.");
    } catch (e) {
        Logger.log('Error setting BJ:BM formulas for stationary row from script templates: ' + e.toString());
    }

    // Re-apply formulas to columns C, E, F, and S in the current position
    try {
      sheet.getRange(insertIndex, 3).setFormula(yearColFormula);   // Column C
      sheet.getRange(insertIndex, 5).setFormula(sz1ColFormula);    // Column E
      sheet.getRange(insertIndex, 6).setFormula(ageColFormula);    // Column F
      sheet.getRange(insertIndex, 19).setFormula(marksColFormula); // Column S (using corrected formula)
      Logger.log("Formulas for C,E,F,S re-applied to existing row position " + insertIndex);
    } catch (e) {
      Logger.log('Error re-applying formulas in insertNewRecordInOrder (stationary row): ' + e.toString());
    }

    return insertIndex;
  }
}

/**
 * Determines the correct sorted index for an edited record.
 *
 * This function:
 *  - Computes where a record *should* be located after its PermID changes.
 *  - Ignores the row currently being edited.
 *  - Does not move data itself — only calculates target index.
 *
 * Why this exists:
 *  - Editing a PermID may break Data sheet ordering.
 *  - The system must reposition the row without duplicating it.
 *
 * How it works:
 *  1. Reads the edited PermID.
 *  2. Scans column D from top to bottom.
 *  3. Skips the current row.
 *  4. Returns the first row where editedPermID < existingPermID.
 *
 * Architectural role:
 *  - Preprocessing step for UPDATE pipeline.
 *  - Used by editRecord() before moving rows.
 *
 * @param {Sheet} sheet - Data sheet.
 * @param {string} permId - Updated PermID.
 * @param {number} currentRow - Row being edited.
 * @returns {number|null}
 *   Target row index, or null if no move required.
 */
function findInsertIndexExcludingRow(data, editedRowData, originalRowNumber) {
  var newGroup = editedRowData[1]; // Assuming column 2 is the group (index 1)
  var newPermID = editedRowData[3]; // Assuming column 4 is the permID (index 3)
  Logger.log('originalRowNumber = ' + originalRowNumber);
  Logger.log("Finding insert index for new record. Group: " + newGroup + ", PermID: " + newPermID);
  logCall("findInsertIndexExcludingRow", { data: data, editedRowData: editedRowData, originalRowNumber: originalRowNumber});

  var lastIndex = data.length + 5; // Initialize to the last possible index (consider header)

  for (var i = 0; i < data.length; i++) {
    var currentRowNumber = i + 5; // Adjust to match sheet row numbers (consider header)
    
    if (currentRowNumber !== originalRowNumber) { // Skip the deleted row
      var currentGroup = data[i][1]; // Assuming column 2 is the group
      var currentPermID = data[i][3]; // Assuming column 4 is the permID
      
      Logger.log('Comparing currentGroup = ' + currentGroup + ' and newGroup = ' + newGroup);
      var groupComparison = compareGroups(newGroup, currentGroup);
      Logger.log('groupComparison = ' + groupComparison);

      if (groupComparison < 0) {
        return currentRowNumber; // Insert before this group
      } else if (groupComparison === 0) {
        Logger.log('Comparing currentPermID = ' + currentPermID + ' on row ' + currentRowNumber + ' and newPermID = ' + newPermID);
        var permIDComparison = comparePermIDs(newPermID, currentPermID, newGroup, currentGroup);
        Logger.log('permIDComparison = ' + permIDComparison + ' AND i : data.lentgh = ' + i + ' : ' + data.length);

        if (permIDComparison < 0) {
          return currentRowNumber; // Insert before this permID in the same group
        } else if (permIDComparison >= 0 && i === data.length - 1) {
          return currentRowNumber + 1; // Insert after this permID (new last record)
        }
      } else {
        lastIndex = currentRowNumber + 1; // Update lastIndex for potential insertion after this group
      }
    }
  }
  Logger.log('lastIndex = ' + lastIndex);
  return lastIndex; // Insert at the end if no earlier position is found
}


function isCohort(group) {
  logCall("isCohort", { group: group });
  return /[A-Z] Cohort \d{4}/.test(group);
}

function getGroupCategory(group, isCohort) {
  logCall("getGroupCategory", { group: group});
  if (isCohort) {
    return 'letter-Cohort-year'; // Cohort group
  } else if (group === 'Non-Cohort Tags') {
    return 'Non-Cohort Tags'; // Non-Cohort Tags group
  } else if (group === 'Untagged') {
    return 'Untagged'; // Untagged group
  } else if (group === 'Temp IDs') {
    return 'Temp IDs'; // Temp IDs group
  } else {
    return 'Other'; // Default fallback category
  }
}

/**
 * Top-level comparison dispatcher for two PermIDs.
 *
 * Determines which comparison strategy applies based on ID type.
 *
 * Group priority (highest → lowest):
 *  1. Cohort IDs
 *  2. Non-cohort permanent IDs
 *  3. Untagged IDs
 *  4. Temporary IDs
 *
 * Architectural role:
 *  - Master comparator for ordering engine.
 *  - Used by all row reordering logic.
 *
 * @param {string} a - First group.
 * @param {string} b - Second group.
 * @returns {number}
 *   Negative if a < b, positive if a > b, 0 if equal.
 */
function compareGroups(group1, group2) {
  Logger.log("Comparing groups: " + group1 + " and " + group2);
  logCall("compareGroups", {group1: group1, group2: group2});

  // Store the cohort checks in variables
  const isGroup1Cohort = isCohort(group1);
  const isGroup2Cohort = isCohort(group2);
  
  // If both groups are cohorts, compare first letter and year
  if (isGroup1Cohort && isGroup2Cohort) {
    var year1 = parseInt(group1.slice(-4));
    var year2 = parseInt(group2.slice(-4));

    if (year1 !== year2) {
      return year1 - year2; // Compare years
    }

    return group1.charAt(0).localeCompare(group2.charAt(0)); // Compare first letter if years are the same
  }

  // If only one of the groups is a cohort, place it above the non-cohort
  if (isGroup1Cohort) return -1;
  if (isGroup2Cohort) return 1;

  // Special handling for "Temp IDs", "Non-Cohort Tags", and "Untagged"
  var order = ['letter-Cohort-year', 'Non-Cohort Tags', 'Untagged', 'Temp IDs'];
  var index1 = order.indexOf(getGroupCategory(group1, isGroup1Cohort));
  var index2 = order.indexOf(getGroupCategory(group2, isGroup2Cohort));

  if (index1 !== -1 && index2 !== -1) {
    Logger.log('index1= ' + index1 + '   index2= ' + index2);
    return index1 - index2; // Compare based on predefined order
  }

  // Fallback to string comparison for other groups
  return group1.localeCompare(group2);
}

/**
 * Compares two permanent (non-cohort) PermIDs.
 *
 * Rules:
 *  - Letter prefix is compared alphabetically.
 *  - Numeric portion is compared numerically.
 *
 * Example ordering:
 *  RG58 < RH92 < RK42
 *
 * @param {string} permID1
 * @param {string} permID2
 * @param {string} group1
 * @param {string} group2
 * @returns {number}
 */
function comparePermIDs(permID1, permID2, group1, group2) { 
  var base1, base2;
  Logger.log('group1 = ' + group1 + ' group2 = ' + group2);
  logCall("comparePermIDs", {permID1: permID1, permID2: permID2, group1: group1, group2: group2});

  // Extract values
  var base1 = extractPermIDBeforeSpecialChars(permID1);
  var base2 = extractPermIDBeforeSpecialChars(permID2);

  Logger.log('base1 =' + base1 + ' base2 = ' + base2);

  // Check for null values and handle them if necessary
  if (!base1 || !base2) {
    Logger.log('Error: Invalid base1 or base2 extracted.');
    return 0;
  }

  switch (group1) { // Since both groups are the same
    case 'letter-Cohort-year':
      return compareCohortIDs(base1, base2);
    case 'Non-Cohort Tags':
      return compareNonCohortIDs(base1, base2);
    case 'Untagged':
      return compareUntaggedIDs(base1, base2);
    case 'Temp IDs':
      Logger.log('calling compareTempIDs : base1 = ' + base1 + ' base2 = ' + base2);
      return compareTempIDs(base1, base2);
    default:
      Logger.log('Falling back to default string comparison');
      return base1.localeCompare(base2);
  }
}

/*
function compareCategoriesOrder(cat1, cat2) {
  const order = ['cohort', 'non-cohort', 'untagged', 'temp', 'other'];
  return order.indexOf(cat1) - order.indexOf(cat2);
}
*/

/**
 * Compares two cohort-style PermIDs.
 *
 * Cohort IDs typically encode:
 *  - Year
 *  - Cohort letter
 *
 * Example:
 *  U Cohort 2025
 *
 * @param {string} id1
 * @param {string} id2
 * @returns {number}
 */
function compareCohortIDs(id1, id2) {
  logCall("compareCohortIDs", { id1: id1, id2: id2});
  const m1 = id1.match(/^([A-Z]) Cohort (\d{4})/);
  const m2 = id2.match(/^([A-Z]) Cohort (\d{4})/);
  if (!m1 || !m2) {
    const bad = !m1 ? id1 : id2;
    Logger.log("compareCohortIDs: WARNING — malformed cohort PermID '" + bad + "'. Falling back to localeCompare.");
    SpreadsheetApp.getUi().alert("Sort warning: PermID '" + bad + "' does not match the expected cohort format (e.g. 'U Cohort 2025'). The record has been placed but may be out of order. Please review and correct this PermID.");
    return id1.localeCompare(id2);
  }
  var [, letter1, year1] = m1;
  var [, letter2, year2] = m2;

  if (year1 !== year2) return year1 - year2;
  if (letter1 !== letter2) return letter1.localeCompare(letter2);

  // If cohort identifiers are the same, compare the rest
  return compareAlphaNumeric(id1.split(': ')[1], id2.split(': ')[1]);
}

/**
 * Compares two non-cohort permanent IDs.
 *
 * Used when IDs are permanent but do not match
 * cohort format.
 *
 * @param {string} id1
 * @param {string} id2
 * @returns {number}
 */
function compareNonCohortIDs(id1, id2){
  logCall("compareNonCohortIDs", { id1: id1, id2: id2});
  Logger.log("compare Non-Cohort Ids fxn");
  Logger.log("ID1: " + id1 + " -> part1: " + id1);
  Logger.log("ID2: " + id2 + " -> part2: " + id2);

  // Function to split ID into letter and number parts, allowing digit comparison first
  function splitID(id) {
    logCall("splitID", { id: id });
    let match = id.match(/^([A-Za-z]+)(\d*)(.*)/); // Match letters, digits, and any non-digit part
    if (match) {
      return {
        letter: match[1],       // Extract the letter part
        number: match[2] ? match[2].split('') : [],  // Split number into digits for comparison
        rest: match[3]          // Extract any remaining non-digit part
      };
    }
    return { letter: '', number: [], rest: '' }; // Default if no match
  }

  let parts1 = splitID(id1);
  let parts2 = splitID(id2);

  // Compare letter parts first
  if (parts1.letter !== parts2.letter) {
    return parts1.letter.localeCompare(parts2.letter); // Sort alphabetically by the letter part
  }

  // Compare digit by digit after the letter
  let maxLength = Math.max(parts1.number.length, parts2.number.length);
  for (let i = 0; i < maxLength; i++) {
    let digit1 = parts1.number[i] ? parseInt(parts1.number[i], 10) : 0; // Pad with 0 if missing
    let digit2 = parts2.number[i] ? parseInt(parts2.number[i], 10) : 0; // Pad with 0 if missing

    if (digit1 !== digit2) {
      Logger.log('Returning difference between digit1 and digit2: ' + digit1 + ' - ' + digit2);
      return digit1 - digit2; // Return difference if digits differ
    }
  }

  // If both have the same numeric part, compare any remaining non-numeric characters
  if (parts1.rest !== parts2.rest) {
    Logger.log('Returning alphabetical comparison of non-numeric parts: ' + parts1.rest + ' vs ' + parts2.rest);
    return parts1.rest.localeCompare(parts2.rest);
  }

  Logger.log('IDs are equal after all comparisons');
  return 0; // If everything is the same
}

/*
function compareNonCohortIDs(id1, id2) {
  var part1 = id1.match(/^([A-Z][A-Z0-9]+)/)[1];
  var part2 = id2.match(/^([A-Z][A-Z0-9]+)/)[1];
  return compareAlphaNumeric(id1, id2);
}
*/

/**
 * Compares two untagged IDs.
 *
 * Untagged records are always sorted
 * after permanent IDs but before temp IDs.
 *
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function compareUntaggedIDs(id1, id2) {
  logCall("compareUntaggedIDs", { id1: id1, id2: id2});
  // Extract the alphanumeric part before "(was temp" for comparison
  const r1 = id1.match(/^([A-Z][A-Z0-9]+)/);
  const r2 = id2.match(/^([A-Z][A-Z0-9]+)/);
  if (!r1 || !r2) {
    const bad = !r1 ? id1 : id2;
    Logger.log("compareUntaggedIDs: WARNING — malformed untagged PermID '" + bad + "'. Falling back to localeCompare.");
    SpreadsheetApp.getUi().alert("Sort warning: PermID '" + bad + "' does not match the expected untagged format (e.g. 'RK42'). The record has been placed but may be out of order. Please review and correct this PermID.");
    return id1.localeCompare(id2);
  }
  var part1 = r1[1];
  var part2 = r2[1];
  Logger.log('Non-Cohort Tag part1 : part2 = ' + part1 + ' : ' + part2);
  return compareAlphaNumeric(part1, part2);
}

/**
 * Compares two temporary IDs.
 *
 * Temporary IDs are always lowest priority
 * in global ordering.
 *
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function compareTempIDs(ID1, ID2) {
  logCall("compareTempIDs", { ID1: ID1, ID2: ID2});
  Logger.log("compareTempIDs fxn");
  Logger.log("ID1: " + ID1 + " -> part1: " + ID1);
  Logger.log("ID2: " + ID2 + " -> part2: " + ID2);

  // Function to split ID into letter and number parts
  function splitID(id) {
    logCall("splitID", { id: id});
    let match = id.match(/^([A-Za-z]*)(\d+)$/);
    if (match) {
      return {
        letter: match[1],
        number: parseInt(match[2], 10)
      };
    }
    return { letter: '', number: parseInt(id, 10) };
  }

  let parts1 = splitID(ID1);
  let parts2 = splitID(ID2);

  // Compare letter parts first
  if (parts1.letter !== parts2.letter) {
    // If letter parts are different, sort alphabetically
    return parts1.letter.localeCompare(parts2.letter);
  }

  // If letter parts are the same (including both empty), compare numbers
  Logger.log('Returning parts1.number - parts2.number:  ' + parts1.number + ' - ' + parts2.number);
  return parts1.number - parts2.number;
}


//function extractTempNumber(id) {
//  var match = id.match(/\d+/);
//  return match ? parseInt(match[0], 10) : 0;
//}

/**
 * Generic alphanumeric comparator.
 *
 * Splits strings into:
 *  - Alphabetic prefix
 *  - Numeric suffix
 *
 * Then compares:
 *  - prefix lexicographically
 *  - suffix numerically
 *
 * Used by all higher-level comparators.
 *
 * @param {string} a
 * @param {string} b
 * @returns {number}
 */
function compareAlphaNumeric(str1, str2) {
  logCall("compareAlphaNumeric", { str1: str1, str2: str2});

  var re = /(\d+|\D+)/g; // Matches sequences of digits or non-digits
  var s1 = str1.match(re);
  var s2 = str2.match(re);
  
  while (s1.length && s2.length) {
    var a = s1.shift();
    var b = s2.shift();
    
    var diff = (a - b) || a.localeCompare(b);
    if (diff) return diff;
  }
  Logger.log('str1 = '+ str1 + '  str2 = ' + str2);
  return s1.length - s2.length; // If one string is shorter than the other
}

/**
 * Extracts alphanumeric parts of permID before special characters.
 *
 * @param {string} permID - The input permID string.
 * @return {string} The extracted alphanumeric part.
 * @customfunction
 */
function extractPermIDBeforeSpecialChars(permID) {
  logCall("extractPermIDBeforeSpecialchars", { permID: permID});
  console.log("Input permID:", permID);

  if (typeof permID !== 'string') {
    return '';
  }

  var result = '';

  // Check if the permID starts with 'temp'
  if (permID.toLowerCase().startsWith('temp')) {
    // If it starts with 'temp', match everything after 'temp' until a space or special character
    var tempMatch = permID.match(/^temp\s*([A-Za-z0-9]+)/i);
    if (tempMatch) {
      result = tempMatch[1];
    }
  } else {
    // For non-'temp' cases, match letter(s) followed by numbers, or just the alphanumeric part
    var match = permID.match(/^([A-Z]+\d+|[A-Z0-9]+)(?=[^A-Z0-9]|$)/i);
    if (match) {
      result = match[1];
    }
  }

  result = result.trim();
  console.log("Extracted permID:", result);
  return result;
}

// Use to test alphanumeric extraction from permid.
function testExtractPermID() {
  var testCases = [
    "temp V91 (4/2023)",
    "temp 301 (2014)",
    "temp 301 2014)",
    "temp 301(2014)",
    "RK36=R4DI",
    "RK36",
    "R8HT",
    "N1AA (Black tags)",
    "N1AA(Black tags)",
    "temp V95",
    "R608 (was V5)"
  ];
  
  for (var i = 0; i < testCases.length; i++) {
    var input = testCases[i];
    var result = extractPermIDBeforeSpecialChars(input);
    console.log("Test case: '" + input + "' -> Result: '" + result + "'");
  }
}

/**
 * Periodically resets highlight styles.
 *
 * @trigger time-based
 */
function checkAndResetBackground() {
  logCall("checkAndResetBackground", { });
  const props = PropertiesService.getScriptProperties();
  const highlightedRows = JSON.parse(props.getProperty('highlightedRows') || '{}');
  const now = new Date().getTime();
  let hasResetAny = false;

  for (const [rowNumber, timestamp] of Object.entries(highlightedRows)) {
    const elapsed = now - parseInt(timestamp);

    if (elapsed >= 10000) { // 10 seconds
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
      const range = sheet.getRange(parseInt(rowNumber), 1, 1, sheet.getLastColumn());
      const backgrounds = range.getBackgrounds()[0]; // Get current backgrounds

      // Set background colors for each column
      for (let i = 0; i < backgrounds.length; i++) {
        if (i + 1 === 2) { // Column B
          backgrounds[i] = '#feffb3';
        } else if (i + 1 === 10 || i + 1 === 14) { // Columns J and N
          backgrounds[i] = '#b7e1cd';
          } else if (i + 1 >= 20 && i + 1 <= 54) { // Columns T:BB
          backgrounds[i] = '#f3f3f3';
          } else if (i + 1 >= 62 && i + 1 <= 65) { // Columns BJ:BM
          backgrounds[i] = '#b7b7b7';
        } else {
          backgrounds[i] = '#ffffff'; // All other columns
        }
      }

      // Apply background colors to row
      range.setBackgrounds([backgrounds]);
      Logger.log("Background colors reset for row " + rowNumber);

      // Apply borders only to columns A:BB
      const borderRange = sheet.getRange(parseInt(rowNumber), 1, 1, 54); // Columns A to BB
      borderRange.setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);

      // Remove row from highlighted rows after reset
      delete highlightedRows[rowNumber];
      hasResetAny = true;      
    }
  }

  if (hasResetAny) {
    props.setProperty('highlightedRows', JSON.stringify(highlightedRows));
  }

  // Update highlightedrows property if any rows were reset
  if (Object.keys(highlightedRows).length === 0) {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'checkAndResetBackground') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  }
}

/**
 * Deletes highlight reset trigger.
 */
function deleteCARRtrigger() {
        // Delete the trigger
      const triggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'checkAndResetBackground') {
          ScriptApp.deleteTrigger(triggers[i]);
        }
      }
}

//---------------------------------- Delete Record Functions --------------------------------------------
/**
 * Primary DELETE operation.
 *
 * Responsibilities:
 *  - Verifies PermID.
 *  - Deletes row.
 *  - Logs full record snapshot.
 *
 * @param {number} rowNumber
 * @param {string} permId
 * @returns {{success:boolean,message:string}}
 */
function DeleteRecord(rowNumber, permId) {
  logCall("DeleteRecord", { rowNumber: rowNumber, permId: permId});
  console.log('DeleteRecord function started');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Data");

  console.log('Attempting to delete row:', rowNumber, 'with permId:', permId);

  // Verify that the permId in the specified row matches the one we're trying to delete
  var permIdInSheet = sheet.getRange(rowNumber, 4).getValue(); // permId is in column D
  console.log('PermId in sheet:', permIdInSheet);

  var range = sheet.getRange(rowNumber, 1);
  sheet.setActiveRange(range);
  
  // Check if the searched permId is a substring of the permId in the sheet
  if (permIdInSheet.toString().toUpperCase().includes(permId.toString().toUpperCase())) {

    // Retrieve the row data from columns A:S
    var fieldNames = ['island', 'group', 'year', 'permid', 'sz1', 'age', 'sz', 'sx', 'tagged', 'lefttag', 'ltresight', 'ltconddorsal', 'ltcondventral', 'righttag', 'rtresight', 'rtconddorsal', 'rtcondventral', 'comment', 'identifyingmarks'];
    var rowData = sheet.getRange(rowNumber, 1, 1, fieldNames.length).getValues()[0]; // Get data from columns A:S (19 columns)
    
    // Format the 'year' column (column C) value to m/d/yyyy
    if (rowData[2] !== "" && rowData[2] !== null) {
      rowData[2] = Utilities.formatDate(new Date(rowData[2]), Session.getScriptTimeZone(), "M/d/yyyy");
    }

    // Construct details string with non-empty fields and uppercase headings
    var detailsArray = [];
    for (var i = 0; i < fieldNames.length; i++) {
      if (rowData[i] !== "" && rowData[i] !== null) {  // Check for non-empty values
        detailsArray.push(fieldNames[i].toUpperCase() + ": " + rowData[i]);
      }
    }
    
    var details = "Deleted record for permId: " + permIdInSheet.toString().toUpperCase() + "; " + detailsArray.join(", ");

    // Delete the row
    sheet.deleteRow(rowNumber);
    console.log('Row deleted successfully');

    // Log the activity
    try {
      var userEmail = Session.getActiveUser().getEmail();
      logActivity("DELETE", permIdInSheet.toString().toUpperCase(), details, userEmail);
      Logger.log('Activity logged successfully.');
    } catch (logError) {
      Logger.log('Error logging activity: ' + logError.toString());
    }
    return {success: true, message: "Record deleted successfully"};
  } else {
    console.log('PermId mismatch. Deletion aborted.');
    throw new Error("PermId mismatch. Deletion aborted.");
  }
}

function testActiveUserEmail() {
  const email = Session.getActiveUser().getEmail();
  Logger.log("Active user email: " + email);
}


