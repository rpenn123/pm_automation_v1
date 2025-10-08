/**
 * @OnlyCurrentDoc
 * Utilities.gs
 * A collection of shared, reusable functions for data normalization, date/time handling,
 * sheet interactions, and general-purpose data manipulation. These utilities form the
 * foundational building blocks for the entire application.
 */

// =================================================================
// ==================== NORMALIZATION ==============================
// =================================================================

/**
 * Normalizes a value into a consistent string representation for reliable comparisons.
 * This function is **critical** for preventing infinite loops in `onEdit` triggers, where a
 * script's own edit might be detected as a new change. It handles various data types:
 * - `null`/`undefined` become an empty string.
 * - Booleans become `"true"` or `"false"`.
 * - Dates are formatted into a consistent, timezone-aware ISO-like string.
 * - All other types are converted to a string and trimmed.
 *
 * @param {*} val The value to normalize.
 * @returns {string} The normalized string representation of the value.
 */
function normalizeForComparison(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "boolean") return val ? "true" : "false";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
  }
  return String(val).trim();
}

/**
 * Normalizes any value into a lowercased, trimmed string. This is the standard
 * function for preparing user-entered strings for case-insensitive comparisons.
 *
 * @param {*} val The value to normalize.
 * @returns {string} The normalized, lowercased string. Returns an empty string if input is null/undefined.
 */
function normalizeString(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toLowerCase();
}

/**
 * Checks if a value is "true-like", useful for handling spreadsheet checkbox values or user input.
 * It evaluates to true for the boolean `true` or for case-insensitive strings like
 * "true", "yes", "y", or "1".
 *
 * @param {*} val The value to check.
 * @returns {boolean} `true` if the value is considered "true-like", otherwise `false`.
 */
function isTrueLike(val) {
  const v = normalizeString(val);
  return v === "true" || v === "yes" || v === "y" || v === "1";
}

// =================================================================
// ==================== DATE UTILITIES =============================
// =================================================================

/**
 * Parses a value into a Date object and normalizes it to the beginning of the day (midnight) in the script's timezone.
 * This is essential for date-based comparisons where the time of day is irrelevant.
 * It robustly handles native Date objects, ISO-like strings (YYYY-MM-DD), and common US date formats (M/D/YYYY).
 *
 * @param {*} input The value to parse (can be a Date object, a date string, or a number).
 * @returns {Date|null} A new Date object set to midnight, or `null` if the value is not a valid date.
 */
function parseAndNormalizeDate(input) {
  if (!input) return null;

  // 1. If it's already a valid Date object, normalize and return a copy.
  if (input instanceof Date) {
    if (isNaN(input.getTime())) return null; // Check for invalid Date objects
    const date = new Date(input.getTime());
    date.setHours(0, 0, 0, 0);
    return date;
  }

  // 2. If it's a string, attempt to parse common formats.
  if (typeof input === 'string') {
    const trimmedInput = input.trim();
    let date;

    // Try to parse YYYY-MM-DD or YYYY/MM/DD format (handles ISO-like dates)
    const isoMatch = trimmedInput.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (isoMatch) {
      const year = parseInt(isoMatch[1], 10);
      const month = parseInt(isoMatch[2], 10);
      const day = parseInt(isoMatch[3], 10);
      date = new Date(year, month - 1, day);
    } else {
      // Fallback for other formats recognized by the native parser, like MM/DD/YYYY.
      // This is less reliable but provides a fallback.
      date = new Date(trimmedInput);
    }

    if (date && !isNaN(date.getTime())) {
      date.setHours(0, 0, 0, 0);
      return date;
    }
  }

  // 3. If it's a number (potentially from a spreadsheet), let the constructor handle it.
  if (typeof input === 'number') {
      const date = new Date(input);
      if (date && !isNaN(date.getTime())) {
          date.setHours(0, 0, 0, 0);
          return date;
      }
  }

  // 4. If all parsing attempts fail, return null.
  return null;
}

/**
 * Formats a value for use in creating unique string-based keys, a core part of the
 * `TransferEngine`'s duplicate checking mechanism. It handles Dates and other values differently:
 * - Dates are formatted as `"yyyy-MM-dd"`.
 * - Other values are normalized to a lowercased, trimmed string.
 *
 * @param {*} value The value to format.
 * @returns {string} The formatted string, ready for use as a key.
 */
function formatValueForKey(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, "UTC", "yyyy-MM-dd");
  }
  return (value !== null && value !== undefined) ? String(value).trim().toLowerCase() : "";
}

/**
 * Generates a padded, sortable month key in `"YYYY-MM"` format from a Date object.
 * This is primarily used by the `LoggerService` to create standardized, chronologically
 * sortable names for monthly log sheets.
 *
 * @param {Date} [d=new Date()] The date to format. Defaults to the current date if not provided.
 * @returns {string} The formatted month key (e.g., `"2024-07"`).
 */
function getMonthKeyPadded(d) {
  const dt = d || new Date();
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
}

// =================================================================
// ==================== SHEET UTILITIES ============================
// =================================================================

/**
 * Finds the 1-based column index of a header by its text in the first row of a sheet.
 * The search is case-insensitive and trims whitespace to be robust against user formatting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @param {string} headerText The text of the header to find (e.g., "Project Name").
 * @returns {number} The 1-based column index of the header, or -1 if not found.
 */
function getHeaderColumnIndex(sheet, headerText) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return -1;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  
  const normalizedTarget = headerText.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase() === normalizedTarget) return i + 1; // Return 1-based index
  }
  return -1;
}

/**
 * Performs a robust, case-insensitive search for a project by name and returns its 1-based row index.
 * This function uses a two-stage lookup for both performance and accuracy:
 * 1.  **TextFinder:** It first uses the highly optimized `TextFinder` for an exact, full-cell match.
 * 2.  **Manual Scan:** If `TextFinder` fails, it falls back to a manual row-by-row scan, which can
 *     sometimes catch edge cases or differently formatted data that TextFinder might miss.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {string} projectName The name of the project to find.
 * @param {number} projectNameCol The 1-based column index where project names are stored.
 * @returns {number} The 1-based row index of the project, or -1 if not found.
 */
function findRowByProjectNameRobust(sheet, projectName, projectNameCol) {
  if (!sheet || !projectName || typeof projectName !== "string" || !projectNameCol) return -1;
  const searchNameTrimmed = projectName.trim();
  if (!searchNameTrimmed) return -1;

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return -1;
    const range = sheet.getRange(2, projectNameCol, lastRow - 1, 1);

    const tf = range.createTextFinder(searchNameTrimmed).matchCase(false).matchEntireCell(true);
    const found = tf.findNext();
    if (found) return found.getRow();

    const vals = range.getValues();
    // Use formatValueForKey to handle dates and other types consistently.
    const targetKey = formatValueForKey(searchNameTrimmed);
    for (let i = 0; i < vals.length; i++) {
      const v = vals[i][0];
      // By using the same key format, we can correctly compare strings to Date objects.
      if (v && formatValueForKey(v) === targetKey) {
        return i + 2; // +2 for 0-index and header row
      }
    }
    return -1;
  } catch (error) {
    Logger.log(`findRowByProjectNameRobust error on "${sheet.getName()}": ${error}`);
    notifyError(`TextFinder lookup failed for "${projectName}" in "${sheet.getName()}"`, error, sheet.getParent() || SpreadsheetApp.getActiveSpreadsheet());
    return -1;
  }
}


/**
 * Finds a row by an exact, case-sensitive match for a value in a specific column.
 * This function is optimized for looking up unique identifiers like a Salesforce ID (SFID)
 * where an exact, case-sensitive match is required.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search.
 * @param {*} value The exact value to find. Must not be null, undefined, or an empty string.
 * @param {number} column The 1-based column index to search within.
 * @returns {number} The 1-based row index of the first match, or -1 if not found.
 */
function findRowByValue(sheet, value, column) {
  if (!sheet || !value || !column) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  const searchValue = String(value);

  for (let i = 0; i < values.length; i++) {
    // Trim the sheet value to make the comparison robust against whitespace.
    if (String(values[i][0]).trim() === searchValue) {
      return i + 2; // +2 adjustment: 0-indexed loop variable + data starts on row 2
    }
  }
  return -1;
}

/**
 * Implements an "SFID-first" lookup strategy to robustly find a row in a sheet.
 * It first tries to find a match using the `sfid`. If no `sfid` is provided or if no match
 * is found, it falls back to searching by the `projectName` for backward compatibility with
 * legacy data that may not have an SFID.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @param {string} sfid The Salesforce ID to search for. Can be null or empty.
 * @param {number} sfidCol The 1-based column index for SFIDs.
 * @param {string} projectName The project name to use as a fallback identifier.
 * @param {number} projectNameCol The 1-based column index for project names.
 * @returns {number} The 1-based row index of the matched row, or -1 if not found.
 */
function findRowByBestIdentifier(sheet, sfid, sfidCol, projectName, projectNameCol) {
  if (sfid) {
    const row = findRowByValue(sheet, sfid, sfidCol);
    if (row !== -1) {
      return row; // Found a definitive match by SFID.
    }
  }
  return findRowByProjectNameRobust(sheet, projectName, projectNameCol);
}

/**
 * Gets a sheet by its name. If the sheet does not exist, it creates and returns the new sheet.
 * This is a convenient idempotent operation used throughout the script to ensure target sheets exist.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {string} sheetName The name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet object.
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

/**
 * Clears and resizes a sheet to a fixed number of rows and (optionally) columns.
 * This is crucial for maintaining a consistent layout on report sheets like the Dashboard,
 * preventing them from growing or shrinking unexpectedly.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to prepare.
 * @param {number} requiredRowCount The exact number of rows the sheet must have.
 * @param {number} [requiredColCount] Optional. The exact number of columns the sheet must have.
 * @returns {void}
 */
function clearAndResizeSheet(sheet, requiredRowCount, requiredColCount) {
  if (!sheet || typeof requiredRowCount !== 'number' || requiredRowCount < 1) {
    throw new Error("Invalid parameters provided to clearAndResizeSheet.");
  }

  sheet.clear();

  const maxRows = sheet.getMaxRows();
  if (maxRows < requiredRowCount) {
    sheet.insertRowsAfter(maxRows, requiredRowCount - maxRows);
  } else if (maxRows > requiredRowCount) {
    sheet.deleteRows(requiredRowCount + 1, maxRows - requiredRowCount);
  }

  if (typeof requiredColCount === 'number' && requiredColCount > 0) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols < requiredColCount) {
      sheet.insertColumnsAfter(maxCols, requiredColCount - maxCols);
    } else if (maxCols > requiredColCount) {
      sheet.deleteColumns(requiredColCount + 1, maxCols - requiredColCount);
    }
  }
}

// =================================================================
// ==================== OBJECT UTILITIES ===========================
// =================================================================

/**
 * Creates a simple mapping object from an array of [source, destination] pairs.
 * This provides a clean, readable way to define column mappings for the `TransferEngine`.
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
 * Gets the maximum numeric value among an object's values.
 * This is used by the `TransferEngine` to determine the required width of a new row when
 * building it from a destination column mapping, ensuring the new row is wide enough.
 *
 * @param {Object<any, number>} obj The object to inspect, expecting numeric values.
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
 * Returns an array with all duplicate values removed, preserving the order of the first appearance.
 * It leverages the high performance of the V8 engine's `Set` object for efficiency.
 *
 * @param {any[]} arr The array to deduplicate.
 * @returns {any[]} A new array containing only the unique elements from the input array.
 */
function uniqueArray(arr) {
  return [...new Set(arr)];
}

// =================================================================
// ==================== LOCKING UTILITIES ==========================
// =================================================================

/**
 * Attempts to acquire a script lock with a retry mechanism, making it more resilient
 * to brief contention.
 *
 * @param {GoogleAppsScript.Lock.Lock} lock The LockService lock object to acquire.
 * @param {number} [maxRetries=3] The maximum number of times to attempt acquiring the lock.
 * @param {number} [delayMs=1000] The delay in milliseconds between retry attempts.
 * @returns {boolean} `true` if the lock was acquired, `false` otherwise.
 */
function acquireLockWithRetry(lock, maxRetries, delayMs) {
  const retries = maxRetries === undefined ? 3 : maxRetries;
  const delay = delayMs === undefined ? 1000 : delayMs;

  for (let i = 0; i < retries; i++) {
    // Try to acquire the lock, waiting up to 100ms.
    if (lock.tryLock(100)) {
      return true;
    }
    // If this wasn't the last attempt, wait before retrying.
    if (i < retries - 1) {
      Utilities.sleep(delay);
    }
  }

  // If the loop completes without acquiring the lock, return false.
  return false;
}