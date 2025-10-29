/**
 * @OnlyCurrentDoc
 *
 * Utilities.gs
 *
 * A collection of shared, reusable functions for data normalization, date/time handling,
 * sheet interactions, and general-purpose data manipulation. These utilities form the
 * foundational building blocks for the entire application.
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

// =================================================================
// ==================== NORMALIZATION ==============================
// =================================================================

/**
 * Normalizes a value into a consistent string representation for reliable comparisons.
 * This function is **critical** for preventing infinite loops in `onEdit` triggers by ensuring that a script's
 * own edit does not appear as a new change. It handles various data types as follows:
 * - `null` or `undefined` become an empty string `""`.
 * - Booleans become `"true"` or `"false"`.
 * - Date objects are formatted into a consistent, timezone-aware ISO-like string (`yyyy-MM-dd'T'HH:mm:ss`).
 * - All other types are converted to a string and trimmed.
 *
 * @param {*} val The value of any type to be normalized.
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
 * Normalizes any value into a lowercased, trimmed string.
 * This is the standard function for preparing user-entered strings for case-insensitive comparisons.
 *
 * @param {*} val The value of any type to normalize.
 * @returns {string} The normalized, lowercased string. Returns an empty string `""` if the input is null or undefined.
 */
function normalizeString(val) {
  if (val === null || val === undefined) return "";
  return String(val).trim().toLowerCase();
}

/**
 * Checks if a value is "true-like", which is useful for handling spreadsheet checkbox values or user input.
 * It evaluates to `true` for the boolean `true` or for case-insensitive strings such as "true", "yes", "y", or "1".
 *
 * @param {*} val The value to check.
 * @returns {boolean} Returns `true` if the value is considered "true-like", otherwise `false`.
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
 * This is essential for date-based comparisons where the time of day is irrelevant. It robustly handles native Date objects,
 * ISO-like strings (e.g., `YYYY-MM-DD`), and common US date formats (e.g., `M/D/YYYY`), while safely ignoring numeric strings.
 *
 * @param {any} input The value to parse, which can be a Date object, a date string, or a number.
 * @returns {Date|null} A new Date object set to midnight, or `null` if the input cannot be parsed as a valid date.
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

  // 2. If it's a string, strictly validate and manually parse it.
  if (typeof input === 'string') {
    const trimmedInput = input.trim();
    if (/^\d+$/.test(trimmedInput)) return null; // Ignore numeric strings.

    let date = null;
    let match;

    // B. Manually parse common date formats to avoid ambiguity with `new Date(string)`.
    // Case 1: MM/DD/YYYY or MM-DD-YYYY
    match = trimmedInput.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/);
    if (match) {
      const month = parseInt(match[1], 10);
      const day = parseInt(match[2], 10);
      const year = parseInt(match[3], 10);
      const tempDate = new Date(year, month - 1, day);
      // Verify no rollover (e.g., input wasn't 2/30/2024).
      if (tempDate.getFullYear() === year && tempDate.getMonth() === month - 1 && tempDate.getDate() === day) {
        date = tempDate;
      }
    } else {
      // Case 2: YYYY-MM-DD or YYYY/MM/DD
      match = trimmedInput.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
      if (match) {
        const year = parseInt(match[1], 10);
        const month = parseInt(match[2], 10);
        const day = parseInt(match[3], 10);
        const tempDate = new Date(year, month - 1, day);
        // Verify no rollover.
        if (tempDate.getFullYear() === year && tempDate.getMonth() === month - 1 && tempDate.getDate() === day) {
          date = tempDate;
        }
      }
    }

    // C. If a valid date was constructed, normalize and return it.
    if (date && !isNaN(date.getTime())) {
      date.setHours(0, 0, 0, 0);
      return date;
    }
  }

  // 3. **Bug Fix**: Do not attempt to parse raw numbers as dates. A project name might be a number,
  // and it should be treated as an identifier, not a date. If a date is intended, the sheet
  // should provide a Date object, which is handled by the first check in this function.
  if (typeof input === 'number') {
    return null;
  }

  // 4. If all parsing attempts fail, return null.
  return null;
}

/**
 * Formats a value for use in creating unique string-based keys, a core part of the `TransferEngine`'s duplicate checking mechanism.
 * It handles Date objects and other values differently to ensure consistent comparisons:
 * - Date objects and recognizable date strings are formatted as `"yyyy-MM-dd"`.
 * - All other values are normalized to a lowercased, trimmed string.
 *
 * @param {any} value The value to format.
 * @returns {string} The formatted string, ready for use as a key.
 */
function formatValueForKey(value) {
  // First, try to parse the value as a date. This handles both actual Date objects
  // and common date-string formats from the spreadsheet.
  const parsedDate = parseAndNormalizeDate(value);
  if (parsedDate) {
    // If it's a valid date, format it consistently.
    return Utilities.formatDate(parsedDate, "UTC", "yyyy-MM-dd");
  }
  // If it's not a date, fall back to the standard string normalization.
  return (value !== null && value !== undefined) ? String(value).trim().toLowerCase() : "";
}

/**
 * Generates a padded, sortable month key in `"YYYY-MM"` format from a Date object.
 * This is primarily used by the `LoggerService` to create standardized, chronologically sortable names for monthly log sheets.
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
 * The search is case-insensitive and trims whitespace to be robust against variations in user formatting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to search in.
 * @param {string} headerText The text of the header to find (e.g., "Project Name").
 * @returns {number} The 1-based column index of the header, or `-1` if not found.
 */
function getHeaderColumnIndex(sheet, headerText) {
  if (!sheet || !headerText) {
    throw new ValidationError("getHeaderColumnIndex requires a valid sheet and headerText.");
  }
  const lastCol = sheet.getMaxColumns();
  if (lastCol < 1) return -1;

  const headers = withRetry(() => {
    return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(v => String(v).trim());
  }, { functionName: "getHeaderColumnIndex:readHeaders" });
  
  const normalizedTarget = headerText.toLowerCase();
  for (let i = 0; i < headers.length; i++) {
    if (headers[i].toLowerCase() === normalizedTarget) return i + 1;
  }
  return -1;
}

/**
 * Performs a robust, case-insensitive search for a project by name and returns its 1-based row index.
 * This function uses a row-by-row scan that normalizes both the search term and the sheet values.
 * **Note on Bug Fix:** This function previously used `formatValueForKey`, which incorrectly parsed date-like strings
 * (e.g., "5/10/2024") as dates, leading to incorrect matches. It now uses `normalizeString` to ensure
 * project names are always treated as case-insensitive strings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to search.
 * @param {string} projectName The name of the project to find.
 * @param {number} projectNameCol The 1-based column index where project names are stored.
 * @returns {number} The 1-based row index of the project, or `-1` if not found.
 * @throws {ValidationError} If input parameters are invalid.
 * @throws {DependencyError} If reading from the sheet fails after retries.
 */
function findRowByProjectNameRobust(sheet, projectName, projectNameCol) {
  if (!sheet || !projectName || typeof projectName !== "string" || !projectNameCol) {
    throw new ValidationError("findRowByProjectNameRobust requires a valid sheet, projectName, and projectNameCol.");
  }
  const searchNameTrimmed = projectName.trim();
  if (!searchNameTrimmed) return -1;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const vals = withRetry(() => {
    return sheet.getRange(2, projectNameCol, lastRow - 1, 1).getValues();
  }, { functionName: "findRowByProjectNameRobust:readColumn" });

  const targetKey = normalizeString(searchNameTrimmed);
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i][0];
    if (v && normalizeString(v) === targetKey) {
      return i + 2;
    }
  }
  return -1;
}


/**
 * Finds a row by an exact, case-sensitive match for a value in a specific column.
 * This function is optimized for looking up unique identifiers like a Salesforce ID (SFID).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to search.
 * @param {any} value The exact value to find.
 * @param {number} column The 1-based column index to search within.
 * @returns {number} The 1-based row index of the first match, or `-1` if not found.
 * @throws {ValidationError} If input parameters are invalid.
 * @throws {DependencyError} If reading from the sheet fails after retries.
 */
function findRowByValue(sheet, value, column) {
  if (!sheet || value == null || !column) {
    throw new ValidationError("findRowByValue requires a valid sheet, value, and column.");
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = withRetry(() => {
    return sheet.getRange(2, column, lastRow - 1, 1).getValues();
  }, { functionName: "findRowByValue:readColumn" });

  const searchValue = String(value).trim();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === searchValue) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Implements an "SFID-first" lookup strategy to find a row in a sheet.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to search in.
 * @param {string} sfid The Salesforce ID to search for. Can be null or empty.
 * @param {number} sfidCol The 1-based column index for SFIDs.
 * @param {string} projectName The project name to use as a fallback identifier.
 * @param {number} projectNameCol The 1-based column index for project names.
 * @returns {number} The 1-based row index of the matched row, or `-1` if not found.
 * @throws {ValidationError} If input parameters are invalid.
 */
function findRowByBestIdentifier(sheet, sfid, sfidCol, projectName, projectNameCol) {
  if (!sheet || !sfidCol || !projectNameCol) {
    throw new ValidationError("findRowByBestIdentifier requires a valid sheet and column indices.");
  }
  if (sfid) {
    const row = findRowByValue(sheet, sfid, sfidCol);
    if (row !== -1) {
      return row;
    }
  }
  if (projectName) {
    return findRowByProjectNameRobust(sheet, projectName, projectNameCol);
  }
  return -1;
}

/**
 * Gets a sheet by its name. If the sheet does not exist, it creates and returns it.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {string} sheetName The name of the sheet to get or create.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The existing or newly created sheet object.
 * @throws {ValidationError} If input parameters are invalid.
 * @throws {DependencyError} If creating the sheet fails after retries.
 */
function getOrCreateSheet(ss, sheetName) {
  if (!ss || !sheetName) {
    throw new ValidationError("getOrCreateSheet requires a valid spreadsheet and sheetName.");
  }
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = withRetry(() => ss.insertSheet(sheetName), { functionName: "getOrCreateSheet:insertSheet" });
  }
  return sheet;
}

/**
 * Clears and resizes a sheet to a fixed number of rows and (optionally) columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to prepare.
 * @param {number} requiredRowCount The exact number of rows the sheet must have.
 * @param {number} [requiredColCount] Optional. The exact number of columns the sheet must have.
 * @returns {void} This function does not return a value.
 * @throws {ValidationError} If input parameters are invalid.
 * @throws {DependencyError} If any sheet modification fails after retries.
 */
function clearAndResizeSheet(sheet, requiredRowCount, requiredColCount) {
  if (!sheet || typeof requiredRowCount !== 'number' || requiredRowCount < 1) {
    throw new ValidationError("Invalid parameters provided to clearAndResizeSheet.");
  }

  withRetry(() => sheet.clear(), { functionName: "clearAndResizeSheet:clear" });

  const maxRows = sheet.getMaxRows();
  if (maxRows < requiredRowCount) {
    withRetry(() => sheet.insertRowsAfter(maxRows, requiredRowCount - maxRows), { functionName: "clearAndResizeSheet:insertRows" });
  } else if (maxRows > requiredRowCount) {
    withRetry(() => sheet.deleteRows(requiredRowCount + 1, maxRows - requiredRowCount), { functionName: "clearAndResizeSheet:deleteRows" });
  }

  if (typeof requiredColCount === 'number' && requiredColCount > 0) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols < requiredColCount) {
      withRetry(() => sheet.insertColumnsAfter(maxCols, requiredColCount - maxCols), { functionName: "clearAndResizeSheet:insertCols" });
    } else if (maxCols > requiredColCount) {
      withRetry(() => sheet.deleteColumns(requiredColCount + 1, maxCols - requiredColCount), { functionName: "clearAndResizeSheet:deleteCols" });
    }
  }
}

// =================================================================
// ==================== OBJECT UTILITIES ===========================
// =================================================================

/**
 * Creates a simple mapping object from an array of `[source, destination]` pairs.
 * This provides a clean, readable way to define column mappings for the `TransferEngine`.
 * Example: `createMapping([[1, 2], [3, 5]])` returns `{1: 2, 3: 5}`.
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
 * Gets the maximum numeric value from an object's values.
 * This is used by the `TransferEngine` to determine the required width of a new row when building it
 * from a destination column mapping, ensuring the new row is wide enough to hold all mapped data.
 *
 * @param {Object<any, number>} obj The object to inspect, expecting numeric values.
 * @returns {number} The highest numeric value found, or `0` if no numeric values exist.
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