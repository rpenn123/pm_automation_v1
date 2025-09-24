/**
 * @OnlyCurrentDoc
 * LastEditService.gs
 * Manages the creation, updating, and initialization of "Last Edit" tracking columns.
 */

/** Ensure Last Edit columns exist on all key sheets defined in CONFIG. */
function ensureAllLastEditColumns(ss) {
  const trackedSheets = CONFIG.LAST_EDIT.TRACKED_SHEETS;
  trackedSheets.forEach(name => {
    const sh = ss.getSheetByName(name);
    // Only proceed if the sheet actually exists in the workbook
    if (sh) ensureLastEditColumns(sh);
  });
}

/** 
 * Ensure Last Edit columns exist for a specific sheet. 
 * Returns { tsCol, relCol } (1-based indices).
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

/** Update Last Edit timestamp and relative formula for a specific row. */
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
 * Generates the spreadsheet formula for calculating relative time.
 * Optimized using the LET function for better readability and performance.
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
 * Optional: Initialize human-readable Last Edit formulas for all existing rows.
 * Useful if adding these columns to an existing dataset.
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