/**
 * @OnlyCurrentDoc
 * Utilities.gs
 * Shared utility functions for normalization, formatting, lookups, and data manipulation.
 */

// =================================================================
// ==================== NORMALIZATION ==============================
// =================================================================

/**
 * Normalizes a value for consistent comparison, crucial for preventing synchronization loops.
 * It handles various data types: null/undefined become an empty string, booleans are converted
 * to string representations, and Dates are formatted into a consistent ISO-like string.
 * All other values are converted to a trimmed string.
 *
 * @param {*} val The value to normalize.
 * @returns {string} The normalized string representation of the value.
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

/**
 * Normalizes any value into a lowercased, trimmed string.
 * If the value is null or undefined, it returns an empty string.
 *
 * @param {*} val The value to normalize.
 * @returns {string} The normalized, lowercased string.
 */
function normalizeString(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toLowerCase();
}

/**
 * Checks if a value is "true-like". This is useful for handling values from checkboxes
 * or user input where "true" could be represented in multiple ways.
 * It checks for "true", "yes", "y", or "1" in a case-insensitive manner.
 *
 * @param {*} val The value to check.
 * @returns {boolean} True if the value is considered "true-like", otherwise false.
 */
function isTrueLike(val) {
  const v = normalizeString(val);
  return v === "true" || v === "yes" || v === "y" || v === "1";
}

// =================================================================
// ==================== DATE UTILITIES =============================
// =================================================================

/**
 * Parses a value into a Date object and normalizes it to the beginning of the day (midnight).
 * This is useful for date-based comparisons where the time of day is irrelevant.
 *
 * @param {*} value The value to parse (can be a Date object, string, or number).
 * @returns {Date|null} A new Date object set to midnight, or null if the value is not a valid date.
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

/**
 * Formats a value for use in creating unique keys, especially for duplicate checks in TransferEngine.
 * Dates are formatted as "yyyy-MM-dd". Other values are normalized to a lowercased, trimmed string.
 *
 * @param {*} value The value to format.
 * @returns {string} The formatted string key.
 */
function formatValueForKey(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return (value !== null && value !== undefined) ? String(value).trim().toLowerCase() : "";
}

/**
 * Generates a padded, sortable month key in "YYYY-MM" format from a Date object.
 * If no date is provided, the current date is used. This is primarily used by the LoggerService
 * for naming monthly log sheets.
 *
 * @param {Date} [d=new Date()] The date to format. Defaults to the current date.
 * @returns {string} The formatted month key (e.g., "2024-07").
 */
function getMonthKeyPadded(d) {
  const dt = d || new Date();
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

/**
 * Generates an array of Date objects, representing the first day of each month
 * between a specified start and end date (inclusive).
 *
 * @param {Date} startDate The first month to include in the list.
 * @param {Date} endDate The last month to include in the list.
 * @returns {Date[]} An array of Date objects.
 */
function generateMonthList(startDate, endDate) {
    const months = [];
    let currentDate = new Date(startDate.getTime());
    while (currentDate <= endDate) {
        months.push(new Date(currentDate));
        currentDate.setMonth(currentDate.getMonth() + 1);
    }
    return months;
}

// =================================================================
// ==================== SHEET UTILITIES ============================
// =================================================================

/**
 * Finds the 1-based column index of a header in the first row of a sheet.
 * The search is case-insensitive and trims whitespace from both the header and the target text.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @param {string} headerText The text of the header to find.
 * @returns {number} The 1-based column index of the header, or -1 if not found.
 */
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
 * Performs a robust, case-insensitive search for a project name in a specified column and returns its 1-based row index.
 * It first attempts a fast search using `TextFinder` for an exact cell match. If that fails, it falls back to a manual
 * scan of the column values, which can handle edge cases `TextFinder` might miss.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {string} projectName The name of the project to find.
 * @param {number} projectNameCol The 1-based column index where project names are stored.
 * @returns {number} The 1-based row index of the project, or -1 if not found or if an error occurs.
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

/**
 * Gets a sheet by its name. If the sheet does not exist, it creates a new one with that name.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet to get the sheet from.
 * @param {string} sheetName The name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet.
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

/**
 * Clears and prepares a sheet by ensuring it has a fixed number of rows.
 * Deletes excess rows or adds missing rows to match a predefined count.
 * This is crucial for maintaining a consistent layout on sheets like the Dashboard.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to prepare.
 * @param {number} requiredRowCount The exact number of rows the sheet should have.
 */
function clearAndResizeSheet(sheet, requiredRowCount) {
  if (!sheet || typeof requiredRowCount !== 'number' || requiredRowCount < 1) {
    throw new Error("Invalid parameters provided to clearAndResizeSheet.");
  }

  // Clear all content, formatting, and data validations.
  sheet.clear();

  // Adjust row count to the fixed number.
  const maxRows = sheet.getMaxRows();
  if (maxRows < requiredRowCount) {
    sheet.insertRowsAfter(maxRows, requiredRowCount - maxRows);
  } else if (maxRows > requiredRowCount) {
    sheet.deleteRows(requiredRowCount + 1, maxRows - requiredRowCount);
  }
}

// =================================================================
// ==================== OBJECT UTILITIES ===========================
// =================================================================

/**
 * Creates a simple mapping object from an array of source-destination pairs.
 * This is used in the TransferEngine to define how columns from a source sheet
 * map to columns in a destination sheet.
 *
 * @param {Array<[number, number]>} pairs An array of pairs, where each pair is `[sourceColumn, destinationColumn]`.
 * @returns {Object<number, number>} An object where keys are source columns and values are destination columns.
 */
function createMapping(pairs) {
  const o = {};
  for (const [source, dest] of pairs) {
    o[source] = dest;
  }
  return o;
}

/**
 * Gets the maximum numeric value among an object's own properties.
 * This is useful for determining the required width of a row when building it from a mapping.
 *
 * @param {Object<any, any>} obj The object to inspect.
 * @returns {number} The highest numeric value found, or 0 if none exist.
 */
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

/**
 * Returns an array with all duplicate values removed.
 * It leverages the high performance of the V8 engine's `Set` object for efficiency.
 *
 * @param {Array<any>} arr The array to deduplicate.
 * @returns {Array<any>} A new array containing only the unique elements from the input array.
 */
function uniqueArray(arr) {
  return [...new Set(arr)];
}