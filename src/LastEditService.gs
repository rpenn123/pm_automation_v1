/**
 * @OnlyCurrentDoc
 * LastEditService.gs
 * Manages the creation, updating, and initialization of "Last Edit" tracking columns.
 */

/**
 * Iterates through all sheets specified in `CONFIG.LAST_EDIT.TRACKED_SHEETS`
 * and ensures that the required "Last Edit" columns exist on each one.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet to process.
 */
function ensureAllLastEditColumns(ss) {
  const trackedSheets = CONFIG.LAST_EDIT.TRACKED_SHEETS;
  trackedSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    // Only proceed if the sheet actually exists in the workbook
    if (sh) ensureLastEditColumns(sh);
  });
}

/**
 * Ensures that a specific sheet has the necessary "Last Edit" columns: a hidden timestamp column
 * and a visible relative time formula column. If the columns do not exist by their header names,
 * they are created at the end of the sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check and modify.
 * @returns {{tsCol: number, relCol: number}} An object containing the 1-based column indices for the timestamp (`tsCol`) and relative time (`relCol`) columns.
 */
function ensureLastEditColumns(sheet) {
  const { AT_HEADER, REL_HEADER } = CONFIG.LAST_EDIT;
  let startingLastCol = sheet.getLastColumn();

  // 1. Handle Timestamp (Hidden) Column
  let tsCol = getHeaderColumnIndex(sheet, AT_HEADER);
  if (tsCol === -1) {
    // Insert column at the end if not found
    sheet.insertColumnAfter(Math.max(1, startingLastCol));
    tsCol = Math.max(1, startingLastCol) + 1;
    sheet.getRange(1, tsCol).setValue(AT_HEADER);
    sheet.hideColumns(tsCol); // Keep raw timestamp hidden
  }

  // 2. Handle Relative Time (Visible) Column
  // Recompute last column as we might have just inserted one
  let afterTsLastCol = sheet.getLastColumn();
  let relCol = getHeaderColumnIndex(sheet, REL_HEADER);
  if (relCol === -1) {
    // Insert column at the end if not found
    sheet.insertColumnAfter(afterTsLastCol);
    relCol = afterTsLastCol + 1;
    sheet.getRange(1, relCol).setValue(REL_HEADER);
  }

  return { tsCol, relCol };
}

/**
 * Updates the "Last Edit" timestamp and the relative time formula for a specific row.
 * This function is typically called by an `onEdit` trigger. It skips header rows.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet where the edit occurred.
 * @param {number} row The 1-based row number that was edited.
 */
function updateLastEditForRow(sheet, row) {
  if (row <= 1) return; // Skip headers
  
  try {
    const cols = ensureLastEditColumns(sheet);
    const tsCell = sheet.getRange(row, cols.tsCol);
    const relCell = sheet.getRange(row, cols.relCol);

    // Set the raw timestamp
    const now = new Date();
    tsCell.setValue(now);

    // Set the relative time formula
    const tsA1 = tsCell.getA1Notation();
    const formula = _generateOptimizedRelativeTimeFormula(tsA1);
    relCell.setFormula(formula);

  } catch (error) {
    // Log and notify if this specific update fails, as it's a critical tracking feature
    Logger.log(`Failed to update Last Edit for ${sheet.getName()} Row ${row}: ${error}`);
    notifyError(`Last Edit update failed on ${sheet.getName()}`, error, sheet.getParent());
  }
}

/**
 * Generates an optimized Google Sheets formula to calculate a human-readable relative time
 * (e.g., "5 minutes", "2 hours", "3 days") from a timestamp cell. It uses the `LET` function
 * for improved performance and readability within the sheet.
 *
 * @private
 * @param {string} tsA1 The A1 notation of the cell containing the raw timestamp (e.g., "Z2").
 * @returns {string} The complete Google Sheets formula.
 */
function _generateOptimizedRelativeTimeFormula(tsA1) {
  return `=IF(${tsA1}="","", 
      LET(duration, NOW()-${tsA1}, 
          minutes, duration*1440, 
          hours, duration*24, 
          days, duration, 
          weeks, duration/7,
          IF(minutes<1, ROUND(minutes)&" minutes",
          IF(hours<1, ROUND(hours)&" hours",
          IF(days<7, ROUND(days)&" days", 
          ROUND(weeks)&" weeks")))))`;
}

/**
 * Initializes or re-applies the "Last Edit" relative time formulas for all existing data rows
 * on the tracked sheets. This is a utility function that can be run manually, useful for
 * backfilling the formulas after the feature has been added to a sheet with existing data.
 */
function initializeLastEditFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackedSheets = CONFIG.LAST_EDIT.TRACKED_SHEETS;

  trackedSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const cols = ensureLastEditColumns(sh);
    const relRange = sh.getRange(2, cols.relCol, lastRow - 1, 1);

    // Determine the column letter of the timestamp column (e.g., "Z")
    const tsA1Header = sh.getRange(1, cols.tsCol).getA1Notation();
    const tsColLetter = tsA1Header.replace(/\d+/g, "");

    const formulas = [];
    for (let r = 2; r <= lastRow; r++) {
      const tsA1 = tsColLetter + r;
      const formula = _generateOptimizedRelativeTimeFormula(tsA1);
      formulas.push([formula]);
    }
    // Apply formulas in a single batch operation
    relRange.setFormulas(formulas);
  });
}