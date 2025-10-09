/**
 * @OnlyCurrentDoc
 *
 * LastEditService.gs
 *
 * Manages the creation, updating, and initialization of "Last Edit" tracking columns.
 * This service provides a user-friendly way to see how recently a row has been modified
 * by adding a human-readable relative timestamp (e.g., "5 min. ago").
 *
 * @version 1.0.0
 * @release 2024-07-29
 */

/**
 * Ensures that all sheets designated for edit tracking have the necessary "Last Edit" columns.
 * It iterates through the sheet names listed in `config.LAST_EDIT.TRACKED_SHEETS`, finds each
 * corresponding sheet object, and calls `ensureLastEditColumns` on it. This function is idempotent
 * and safe to run multiple times.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet instance containing the sheets to process.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function ensureAllLastEditColumns(ss, config) {
  const trackedSheets = config.LAST_EDIT.TRACKED_SHEETS;
  trackedSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    // Only proceed if the sheet actually exists in the workbook
    if (sh) ensureLastEditColumns(sh, config);
  });
}

/**
 * Ensures a specific sheet has the "Last Edit" columns: a hidden timestamp and a visible relative time.
 * If the columns do not exist by their header names (defined in `config.LAST_EDIT`), this function
 * creates them at the end of the sheet. The raw timestamp column is automatically hidden to reduce clutter.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to check and potentially modify.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {{tsCol: number, relCol: number}} An object containing the 1-based column indices for the timestamp (`tsCol`) and relative time (`relCol`) columns.
 */
function ensureLastEditColumns(sheet, config) {
  const { AT_HEADER, REL_HEADER } = config.LAST_EDIT;
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
 * This is the core of the live-updating "Last Edit" feature and is typically called from an `onEdit` trigger.
 * It ensures the necessary columns exist, writes the current timestamp, and sets a formula to calculate the relative time.
 * Errors are handled gracefully to prevent halting the entire `onEdit` process.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object where the edit occurred.
 * @param {number} row The 1-based row number that was edited.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function updateLastEditForRow(sheet, row, config) {
  if (row <= 1) return; // Skip headers
  
  try {
    const cols = ensureLastEditColumns(sheet, config);
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
    notifyError(`Last Edit update failed on ${sheet.getName()}`, error, sheet.getParent(), config);
  }
}

/**
 * Generates an optimized Google Sheets formula to calculate a human-readable relative time (e.g., "5 min. ago").
 * This private helper function uses the `LET` function for performance by calculating the time difference once.
 * It checks progressively larger time units to provide a user-friendly output.
 *
 * @private
 * @param {string} tsA1 The A1 notation of the cell containing the raw timestamp (e.g., "Z2").
 * @returns {string} The complete Google Sheets formula as a string.
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
 * Initializes or re-applies "Last Edit" formulas for all data rows in all tracked sheets.
 * This function is designed to be run manually from the script editor or a custom menu. Its primary use case
 * is to backfill the relative time formulas for all existing data after the "Last Edit" feature is first deployed.
 * It processes each tracked sheet in a batch operation for efficiency.
 *
 * @returns {void} This function does not return a value.
 */
function initializeLastEditFormulas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG; // Access global config here
  const trackedSheets = config.LAST_EDIT.TRACKED_SHEETS;

  trackedSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const cols = ensureLastEditColumns(sh, config);
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