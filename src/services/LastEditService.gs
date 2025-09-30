/**
 * @OnlyCurrentDoc
 * LastEditService.gs
 * Manages the creation, updating, and initialization of "Last Edit" tracking columns.
 * This service provides a user-friendly way to see how recently a row has been modified.
 */

/**
 * Ensures that all sheets designated for edit tracking have the necessary "Last Edit" columns.
 * It iterates through the sheet names listed in `CONFIG.LAST_EDIT.TRACKED_SHEETS`,
 * finds each corresponding sheet object, and calls `ensureLastEditColumns` on it.
 * This function is idempotent and safe to run multiple times.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet instance containing the sheets to process.
 * @returns {void}
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
 * Ensures a specific sheet has the "Last Edit" columns: a hidden timestamp and a visible relative time.
 * If the columns do not exist by their header names (defined in `CONFIG.LAST_EDIT`),
 * this function creates them at the end of the sheet. The raw timestamp column is hidden
 * from users to reduce clutter.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to check and potentially modify.
 * @returns {{tsCol: number, relCol: number}} An object containing the 1-based column indices
 *   for the timestamp (`tsCol`) and relative time (`relCol`) columns.
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
 * Updates the "Last Edit" timestamp and relative time formula for a specific row.
 * This function is the core of the live-updating "Last Edit" feature and is
 * typically called from a main `onEdit` trigger. It performs the following steps:
 * 1. Ensures the necessary "Last Edit" columns exist.
 * 2. Writes the current timestamp to the hidden timestamp column.
 * 3. Sets a formula in the visible "Last Edit" column that calculates the relative time.
 * It gracefully handles and logs errors to prevent halting the entire `onEdit` process.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet where the edit occurred.
 * @param {number} row The 1-based row number that was edited.
 * @returns {void}
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
 * Generates an optimized and correct Google Sheets formula to calculate a human-readable relative time.
 * This corrected formula (e.g., "5 min. ago", "2 hr. ago") uses `LET` for performance.
 * It checks progressively larger time units and provides a user-friendly output.
 *
 * @private
 * @param {string} tsA1 The A1 notation of the cell containing the raw timestamp (e.g., "Z2").
 * @returns {string} The complete, corrected Google Sheets formula.
 */
function _generateOptimizedRelativeTimeFormula(tsA1) {
  // This formula uses LET for performance by calculating the duration once.
  // It then checks progressively larger units of time and returns the first appropriate match.
  return `=IF(${tsA1}="", "",
    LET(
      diff, NOW() - ${tsA1},
      minutes, diff * 1440,
      hours, diff * 24,
      days, diff,
      weeks, diff / 7,
      IF(minutes < 1, "just now",
      IF(minutes < 60, ROUND(minutes) & " min. ago",
      IF(hours < 24, ROUND(hours) & " hr. ago",
      IF(days < 7, ROUND(days) & " day(s) ago",
      ROUND(weeks) & " wk. ago"
    ))))))`;
}

/**
 * Initializes or re-applies "Last Edit" formulas for all data rows in tracked sheets.
 * This function is designed to be run manually from the script editor or a custom menu.
 * Its primary use case is to backfill the relative time formulas for all existing data
 * after the "Last Edit" feature has been deployed for the first time. It processes
 * each tracked sheet in a batch operation for efficiency.
 *
 * @returns {void}
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