/**
 * @OnlyCurrentDoc
 *
 * LastEditService.gs
 *
 * Manages the creation, updating, and initialization of "Last Edit" tracking columns.
 * This service provides a user-friendly way to see how recently a row has been modified
 * by adding a human-readable relative timestamp (e.g., "5 min. ago").
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

/**
 * Ensures that all sheets designated for edit tracking have the necessary "Last Edit" columns.
 * This function is a key part of the initial setup routine (`Setup.gs`). It iterates through the sheet
 * names listed in `config.LAST_EDIT.TRACKED_SHEETS`, finds each corresponding sheet object, and
 * calls `ensureLastEditColumns` on it. This function is idempotent, making it safe to run multiple times
 * as it will only add columns if they are missing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet instance containing the sheets to process.
 * @param {object} config The global configuration object (`CONFIG`), used to find the list of tracked sheets.
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
 * Ensures a specific sheet has the required "Last Edit" columns: a hidden raw timestamp and a visible relative time.
 * This function is idempotent. If the columns do not exist (identified by their header names defined in `config.LAST_EDIT`),
 * this function creates them at the end of the sheet. The raw timestamp column is automatically hidden from users to reduce UI clutter.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to check and potentially modify.
 * @param {object} config The global configuration object (`CONFIG`), used to get the required header names.
 * @returns {{tsCol: number, relCol: number}} An object containing the 1-based column indices for the newly verified or created timestamp (`tsCol`) and relative time (`relCol`) columns.
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
 * Updates the "Last Edit" timestamp and relative time formula for a specific row in response to an edit.
 * This function is the core of the live-updating "Last Edit" feature and is called from the main `onEdit` trigger
 * in `Automations.gs`. It first ensures the necessary columns exist, then writes the current timestamp to the hidden
 * column and sets a dynamic Google Sheets formula in the visible column to calculate the relative time.
 * Errors are handled gracefully to prevent a failure here from halting the entire `onEdit` process.
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
 * This private helper function constructs a formula that uses the `LET` function for improved performance by
 * calculating the time difference (`NOW() - timestamp`) only once. It then uses a series of nested `IF`
 * statements to check progressively larger time units (minutes, hours, days, weeks) and provide a clean,
 * user-friendly output like "just now", "10 min. ago", or "2 wk. ago".
 *
 * @private
 * @param {string} tsA1 The A1 notation of the cell containing the raw timestamp (e.g., "Z2").
 * @returns {string} The complete Google Sheets formula as a string, ready to be set in a cell.
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
 * Initializes or re-applies "Last Edit" formulas for all data rows across all tracked sheets.
 * This function is designed to be run manually from the custom menu (`🚀 Project Actions`). Its primary use case
 * is to backfill the relative time formulas for all existing data after the "Last Edit" feature is first deployed
 * or to repair formulas that have been accidentally deleted by users. It processes each tracked sheet and applies
 * the formulas in a single batch operation per sheet for efficiency.
 *
 * @returns {void} This function does not return a value; it modifies the spreadsheet directly.
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