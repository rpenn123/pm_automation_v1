/**
 * @OnlyCurrentDoc
 * Utilities.gs
 * Shared utility functions for normalization, formatting, lookups, and data manipulation.
 */

// =================================================================
// ==================== NORMALIZATION ==============================
// =================================================================

/**
 * Normalize values for comparison to avoid synchronization loops.
 * Handles various types including Dates and Booleans.
 */
function normalizeForComparison(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "boolean") return val ? "true" : "false";
  if (val instanceof Date) {
    // Format date consistently for comparison
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(val).trim();
}

/** Normalize any value to a lowercased, trimmed string. */
function normalizeString(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toLowerCase();
}

/** True-like helper for checkbox or text TRUE/YES/Y/1. */
function isTrueLike(val) {
  const v = normalizeString(val);
  return v === "true" || v === "yes" || v === "y" || v === "1";
}

// =================================================================
// ==================== DATE UTILITIES =============================
// =================================================================

/**
 * Parses a value into a Date object and normalizes it to the start of the day (midnight).
 * Returns null if the value is not a valid date.
 */
function parseAndNormalizeDate(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const d = new Date(value);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (value) {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      d.setHours(0, 0, 0, 0);
      return d;
    }
  }
  return null;
}

/** Format values for key building (dates -> yyyy-MM-dd). Used in TransferEngine duplicate checks. */
function formatValueForKey(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return (value !== null && value !== undefined) ? String(value).trim().toLowerCase() : "";
}

/** Return YYYY-MM (padded) for a date (current if omitted). Used in LoggerService. */
function getMonthKeyPadded(d) {
  const dt = d || new Date();
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

// =================================================================
// ==================== SHEET UTILITIES ============================
// =================================================================

/** Find a header by text in row 1 and return its 1-based column index. */
function getHeaderColumnIndex(sheet, headerText) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  // Read headers and normalize
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  
  const normalizedTarget = headerText.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase() === normalizedTarget) return i + 1; // Return 1-based index
  }
  return -1;
}

/**
 * Robust row lookup by Project Name.
 * Tries TextFinder (entire cell, case-insensitive) then a fallback manual scan.
 * Returns 1-based row index or -1 if not found.
 */
function findRowByProjectNameRobust(sheet, projectName, projectNameCol) {
  if (!sheet || !projectName || typeof projectName !== "string" || !projectNameCol) return -1;
  const searchNameTrimmed = projectName.trim();
  if (!searchNameTrimmed) return -1;

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    const range = sheet.getRange(2, projectNameCol, lastRow - 1, 1);

    // 1. Try TextFinder (often faster, more reliable for exact matches)
    const tf = range.createTextFinder(searchNameTrimmed).matchCase(false).matchEntireCell(true);
    const found = tf.findNext();
    if (found) return found.getRow();

    // 2. Fallback: manual scan (handles edge cases TextFinder might miss)
    const vals = range.getValues();
    const target = searchNameTrimmed.toLowerCase();
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][0];
      if (v && String(v).trim().toLowerCase() === target) return i + 2; // +2 because 0-indexed array + starting from row 2
    }
    return -1;
  } catch (error) {
    Logger.log(`findRowByProjectNameRobust error on "${sheet.getName()}": ${error}`);
    // Use the centralized error notification system (LoggerService must be available globally)
    notifyError(`TextFinder lookup failed for "${projectName}" in "${sheet.getName()}"`, error, sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet());
    return -1;
  }
}

/** Helper to get or create a sheet. Used in Dashboard.gs. */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

// =================================================================
// ==================== OBJECT UTILITIES ===========================
// =================================================================

/** Create a mapping object from pairs [[sourceCol, destCol], ...]. Used for Transfers. */
function createMapping(pairs) {
  const o = {};
  for (const [source, dest] of pairs) {
    o[source] = dest;
  }
  return o;
}

/** Get max numeric value in an object's own property values. */
function getMaxValueInObject(obj) {
  let max = 0;
  for (const k in obj) {
    if (Object.prototype.hasOwnProperty.call(obj, k)) {
      const v = obj[k];
      if (typeof v === "number" && v > max) max = v;
    }
  }
  return max;
}

/** Return unique array of numbers. Uses V8 Set for efficiency. */
function uniqueArray(arr) {
  return [...new Set(arr)];
}