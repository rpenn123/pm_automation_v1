/**
 * @OnlyCurrentDoc
 *
 * TransferEngine.gs
 *
 * This script provides a generic, reusable engine for transferring data between sheets based on a configuration object.
 * It is designed to be the robust backend for all data transfer operations, handling locking, duplicate checking,
 * flexible data mapping, and post-transfer actions like sorting.
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

/**
 * Executes a generic, configuration-driven data transfer from a source row to a destination sheet.
 * This function is the core of the transfer mechanism, designed to be called by trigger handlers in `Automations.gs`.
 * It manages the entire transfer lifecycle, including locking, validation, duplicate checking, data mapping, and post-transfer actions.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object passed from the trigger.
 * @param {object} config The configuration object that defines all aspects of the transfer.
 * @param {string} config.transferName A descriptive name for the transfer, used for logging (e.g., "Upcoming Transfer").
 * @param {string} config.destinationSheetName The name of the sheet to which data will be transferred.
 * @param {Object<number, number>} config.destinationColumnMapping An object mapping source column indices (keys) to destination column indices (values).
 * @param {string[]} [config.lastEditTrackedSheets] An array of sheet names where "Last Edit" timestamps should be updated.
 * @param {object} [config.duplicateCheckConfig] Configuration for preventing duplicate entries. If omitted, no duplicate check is performed.
 * @param {boolean} [config.duplicateCheckConfig.checkEnabled=true] If `false`, the duplicate check is skipped.
 * @param {number} [config.duplicateCheckConfig.sfidSourceCol] The 1-based column index of the Salesforce ID in the source sheet.
 * @param {number} [config.duplicateCheckConfig.sfidDestCol] The 1-based column index of the Salesforce ID in the destination sheet.
 * @param {number} config.duplicateCheckConfig.projectNameSourceCol The 1-based column index of the project name in the source sheet, used as a fallback identifier.
 * @param {number} config.duplicateCheckConfig.projectNameDestCol The 1-based column index of the project name in the destination sheet.
 * @param {number[]} [config.duplicateCheckConfig.compoundKeySourceCols] For fallback checks, an array of additional source column indices to create a compound key.
 * @param {number[]} [config.duplicateCheckConfig.compoundKeyDestCols] The corresponding destination column indices for the compound key.
 * @param {string} [config.duplicateCheckConfig.keySeparator="|"] The separator character for building compound keys.
 * @param {object} [config.postTransferActions] Actions to perform after a successful transfer.
 * @param {boolean} [config.postTransferActions.sort=false] If `true`, sorts the destination sheet after appending the new row.
 * @param {number} [config.postTransferActions.sortColumn] The 1-based column index to sort by.
 * @param {boolean} [config.postTransferActions.sortAscending] If `true`, sorts in ascending order; otherwise, descending.
 * @param {any[]} [preReadSourceRowData] Optional. The pre-read data from the source row to avoid another I/O call. If not provided, the function will read it.
 * @param {string} correlationId A unique ID for tracing the entire operation.
 * @returns {void} This function does not return a value.
 */
function executeTransfer(e, config, preReadSourceRowData, correlationId) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const sourceSheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const ss = e.source;
  let projectName = ""; // Initialize for logging

  try {
    withRetry(() => {
      lockAcquired = lock.tryLock(2500);
      if (!lockAcquired) throw new Error("Lock not acquired within the time limit.");
    }, { functionName: `${config.transferName}:acquireLock`, maxRetries: 2, initialDelayMs: 500 });

    const destinationSheet = ss.getSheetByName(config.destinationSheetName);
    if (!destinationSheet) {
      throw new ConfigurationError(`Destination sheet "${config.destinationSheetName}" not found`);
    }

    let sourceRowData = preReadSourceRowData;
    let readWidth;

    if (sourceRowData) {
      readWidth = sourceRowData.length;
    } else {
      const mappedSourceCols = Object.keys(config.destinationColumnMapping || {}).map(Number);
      const compoundKeyCols = (config.duplicateCheckConfig && config.duplicateCheckConfig.compoundKeySourceCols) || [];
      const dupCheckCols = (config.duplicateCheckConfig && [config.duplicateCheckConfig.sfidSourceCol, config.duplicateCheckConfig.projectNameSourceCol]) || [];
      const maxSourceColNeeded = Math.max(...(config.sourceColumnsNeeded || []), ...mappedSourceCols, ...compoundKeyCols, ...dupCheckCols);
      const actualLastSourceCol = sourceSheet.getMaxColumns();
      readWidth = Math.min(maxSourceColNeeded, actualLastSourceCol);
      sourceRowData = withRetry(() => sourceSheet.getRange(editedRow, 1, 1, readWidth).getValues()[0], { functionName: `${config.transferName}:readSourceRow` });
    }

    const { sfidSourceCol, projectNameSourceCol } = config.duplicateCheckConfig || {};
    let sfid = sfidSourceCol && sfidSourceCol <= readWidth ? sourceRowData[sfidSourceCol - 1] : null;
    projectName = projectNameSourceCol && projectNameSourceCol <= readWidth ? sourceRowData[projectNameSourceCol - 1] : "";
    sfid = sfid ? String(sfid).trim() : null;
    projectName = projectName ? String(projectName).trim() : "";

    if (!sfid && !projectName) {
      throw new ValidationError(`Row ${editedRow} is missing both SFID and Project Name.`);
    }

    // Perform Duplicate Check to find row for potential update or to confirm uniqueness.
    let duplicateRowIndex = -1;
    if (config.duplicateCheckConfig && config.duplicateCheckConfig.checkEnabled !== false) {
      duplicateRowIndex = findDuplicateRow(destinationSheet, sfid, projectName, sourceRowData, readWidth, config.duplicateCheckConfig, correlationId);
    }

    // Build the destination row data. This is needed for both append and update.
    const mapping = config.destinationColumnMapping || {};
    const maxMappedCol = getMaxValueInObject(mapping);
    const destLastCol = Math.max(destinationSheet.getMaxColumns(), maxMappedCol);
    const newRowData = new Array(destLastCol).fill("");

    for (const sourceColStr in mapping) {
      if (!Object.prototype.hasOwnProperty.call(mapping, sourceColStr)) continue;
      const sourceCol = Number(sourceColStr);
      const destCol = mapping[sourceColStr];
      if (sourceCol <= readWidth) {
        newRowData[destCol - 1] = sourceRowData[sourceCol - 1] ?? "";
      } else {
        handleError(new ConfigurationError(`Source col ${sourceCol} not available in read data. Skipped mapping.`), {
          correlationId, functionName: "executeTransfer", spreadsheet: ss,
          extra: { transferName: config.transferName, sourceSheet: sourceSheet.getName(), editedRow }
        }, CONFIG);
      }
    }

    // Decide whether to update, append, or skip.
    if (duplicateRowIndex !== -1) {
      if (config.syncOnDuplicate) {
        // --- SYNC/UPDATE PATH ---
        updateRowInDestination(destinationSheet, duplicateRowIndex, newRowData, config, correlationId);
        if (config.lastEditTrackedSheets && config.lastEditTrackedSheets.includes(config.destinationSheetName)) {
          updateLastEditForRow(destinationSheet, duplicateRowIndex, CONFIG);
        }
        logAudit(ss, {
          correlationId, action: config.transferName, sourceSheet: sourceSheet.getName(), sourceRow: editedRow,
          projectName, details: `Updated row ${duplicateRowIndex} in ${config.destinationSheetName}`, result: "success-updated"
        }, CONFIG);
      } else {
        // --- SKIP DUPLICATE PATH (original behavior) ---
        const logIdentifier = sfid ? `SFID ${sfid}` : `project "${projectName}"`;
        logAudit(ss, {
          correlationId, action: config.transferName, sourceSheet: sourceSheet.getName(), sourceRow: editedRow,
          sfid, projectName, details: `Duplicate detected for ${logIdentifier}.`, result: "skipped-duplicate"
        }, CONFIG);
      }
    } else {
      // --- APPEND PATH (no duplicate found) ---
      withRetry(() => destinationSheet.appendRow(newRowData), { functionName: `${config.transferName}:appendRow` });
      const appendedRow = destinationSheet.getLastRow();
      if (config.lastEditTrackedSheets && config.lastEditTrackedSheets.includes(config.destinationSheetName)) {
        updateLastEditForRow(destinationSheet, appendedRow, CONFIG);
      }

      if (config.postTransferActions && config.postTransferActions.sort && appendedRow > 2) {
        try {
          withRetry(() => SpreadsheetApp.flush(), { functionName: 'SpreadsheetApp.flush' });
          const { sortColumn, sortAscending } = config.postTransferActions;
          const range = destinationSheet.getRange(2, 1, appendedRow - 1, destinationSheet.getMaxColumns());
          withRetry(() => range.sort({ column: sortColumn, ascending: !!sortAscending }), { functionName: `${config.transferName}:sortDestination` });
        } catch (sortError) {
          handleError(new DependencyError(`${config.transferName} completed, but post-transfer sort failed.`, sortError), {
            correlationId, functionName: "executeTransfer:postTransferSort", spreadsheet: ss,
            extra: { transferName: config.transferName }
          }, CONFIG);
        }
      }

      logAudit(ss, {
        correlationId, action: config.transferName, sourceSheet: sourceSheet.getName(), sourceRow: editedRow,
        projectName, details: `Appended to ${config.destinationSheetName} row ${appendedRow}`, result: "success"
      }, CONFIG);
    }

  } catch (error) {
    handleError(error, {
      correlationId, functionName: "executeTransfer", spreadsheet: ss,
      extra: { transferName: config.transferName, sourceSheet: sourceSheet.getName(), editedRow, projectName }
    }, CONFIG);
    logAudit(ss, {
      correlationId, action: config.transferName, sourceSheet: sourceSheet.getName(), sourceRow: editedRow,
      projectName, result: "error", errorMessage: `${error.name}: ${error.message}`
    }, CONFIG);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Performs a robust duplicate check in the destination sheet using an SFID-first strategy.
 * This helper function for `executeTransfer` is critical for data integrity. It uses two main strategies:
 * 1.  **SFID Check (Primary):** If a non-empty `sfid` is provided, it performs a fast, exact-match search on the destination SFID column.
 * 2.  **Compound Key Check (Fallback):** If no SFID is available, it constructs a unique key from the Project Name plus any additional columns
 *     defined in `compoundKeySourceCols`. This maintains compatibility with legacy data that may not have an SFID.
 *     **Note on Bug Fix:** The compound key generation previously used `formatValueForKey`, which incorrectly treated date-like strings
 *     as dates. It now uses `normalizeString` to ensure all parts of the key are compared as simple strings.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destinationSheet The sheet object where duplicates will be checked.
 * @param {string|null} sfid The Salesforce ID from the source row. This is the preferred unique identifier.
 * @param {string} projectName The project name from the source row, used for the fallback check.
 * @param {any[]} sourceRowData The array of values from the source row, used to build the compound key.
 * @param {number} sourceReadWidth The number of columns that were read from the source row.
 * @param {object} dupConfig The configuration object for the duplicate check, passed from `executeTransfer`.
 * @param {string} correlationId A unique ID for tracing the entire operation.
 * @returns {boolean} Returns `true` if a duplicate is found, otherwise `false`.
 * @throws {ConfigurationError} If the configuration for the duplicate check is invalid (e.g., missing columns).
 */
function isDuplicateInDestination(destinationSheet, sfid, projectName, sourceRowData, sourceReadWidth, dupConfig, correlationId) {
  // Strategy 1: SFID is the primary, definitive check.
  if (sfid && dupConfig.sfidDestCol) {
    const foundRow = withRetry(() => findRowByValue(destinationSheet, sfid, dupConfig.sfidDestCol), {
      functionName: "isDuplicateInDestination:findRowByValue",
      correlationId: correlationId
    });
    return foundRow !== -1;
  }

  // Strategy 2: Fallback to Project Name + Compound Key if no SFID is present.
  if (!projectName) {
    return false; // Cannot perform fallback check without a project name.
  }

  const destProjectNameCol = dupConfig.projectNameDestCol;
  const lastDestRow = destinationSheet.getLastRow();
  if (lastDestRow < 2) return false; // No data rows exist

  const sep = dupConfig.keySeparator || "|";

  // Create pairs for the compound key, ensuring consistent order
  const keyPairs = [];
  if (dupConfig.compoundKeySourceCols && dupConfig.compoundKeyDestCols && dupConfig.compoundKeySourceCols.length === dupConfig.compoundKeyDestCols.length) {
    for (let i = 0; i < dupConfig.compoundKeySourceCols.length; i++) {
      keyPairs.push({ source: dupConfig.compoundKeySourceCols[i], dest: dupConfig.compoundKeyDestCols[i] });
    }
    keyPairs.sort((a, b) => a.source - b.source);
  }

  // 1. Build the fallback key to check against from the source data
  let keyToCheck = normalizeString(projectName);
  for (const pair of keyPairs) {
    const val = (pair.source <= sourceReadWidth) ? sourceRowData[pair.source - 1] : undefined;
    keyToCheck += sep + normalizeString(val);
  }

  // 2. Determine columns needed from the destination for comparison
  let colsToRead = [destProjectNameCol];
  if (keyPairs.length > 0) {
    colsToRead = uniqueArray(colsToRead.concat(keyPairs.map(p => p.dest)));
  }
  const minCol = Math.min(...colsToRead);
  const maxCol = Math.max(...colsToRead);
  const readWidth = maxCol - minCol + 1;

  if (maxCol > destinationSheet.getMaxColumns()) {
    throw new ConfigurationError("Duplicate check failed: destination sheet missing expected columns for compound key.");
  }

  // 3. Read destination data in a batch, with retry
  const vals = withRetry(() => destinationSheet.getRange(2, minCol, lastDestRow - 1, readWidth).getValues(), {
    functionName: "isDuplicateInDestination:readDestinationData",
    correlationId: correlationId
  });

  const projIdx = destProjectNameCol - minCol;

  // 4. Scan destination data for the key
  for (const row of vals) {
    if (projIdx >= row.length) continue;
    let existingKey = normalizeString(row[projIdx]);
    if (!existingKey) continue;

    // Build the key from the destination row using the same sorted order
    for (const pair of keyPairs) {
      const destColIndex = pair.dest - minCol;
      const v = (destColIndex < row.length) ? row[destColIndex] : "";
      existingKey += sep + normalizeString(v);
    }

    if (existingKey === keyToCheck) return true;
  }

  return false;
}

/**
 * Locates a duplicate row in the destination sheet, returning its row number.
 * This function is an adaptation of `isDuplicateInDestination`, but instead of a boolean, it returns the
 * 1-based row index of the first duplicate found, or -1 if no duplicate is found. It uses the same
 * robust SFID-first or compound-key-fallback strategies.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destinationSheet The sheet object where duplicates will be checked.
 * @param {string|null} sfid The Salesforce ID from the source row.
 * @param {string} projectName The project name from the source row.
 * @param {any[]} sourceRowData The array of values from the source row.
 * @param {number} sourceReadWidth The number of columns read from the source row.
 * @param {object} dupConfig The configuration for the duplicate check.
 * @param {string} correlationId A unique ID for tracing the operation.
 * @returns {number} The 1-based row index of the duplicate, or -1 if not found.
 */
function findDuplicateRow(destinationSheet, sfid, projectName, sourceRowData, sourceReadWidth, dupConfig, correlationId) {
  // Strategy 1: SFID is the primary, definitive check.
  if (sfid && dupConfig.sfidDestCol) {
    return withRetry(() => findRowByValue(destinationSheet, sfid, dupConfig.sfidDestCol), {
      functionName: "findDuplicateRow:findRowBySfid",
      correlationId: correlationId
    });
  }

  // Strategy 2: Fallback to Project Name + Compound Key if no SFID is present.
  if (!projectName) {
    return -1; // Cannot perform fallback check without a project name.
  }

  const destProjectNameCol = dupConfig.projectNameDestCol;
  const lastDestRow = destinationSheet.getLastRow();
  if (lastDestRow < 2) return -1; // No data rows exist

  const sep = dupConfig.keySeparator || "|";

  // Create pairs for the compound key, ensuring consistent order
  const keyPairs = [];
  if (dupConfig.compoundKeySourceCols && dupConfig.compoundKeyDestCols && dupConfig.compoundKeySourceCols.length === dupConfig.compoundKeyDestCols.length) {
    for (let i = 0; i < dupConfig.compoundKeySourceCols.length; i++) {
      keyPairs.push({ source: dupConfig.compoundKeySourceCols[i], dest: dupConfig.compoundKeyDestCols[i] });
    }
    keyPairs.sort((a, b) => a.source - b.source);
  }

  // 1. Build the fallback key to check against from the source data
  let keyToCheck = normalizeString(projectName);
  for (const pair of keyPairs) {
    const val = (pair.source <= sourceReadWidth) ? sourceRowData[pair.source - 1] : undefined;
    keyToCheck += sep + normalizeString(val);
  }

  // 2. Determine columns needed from the destination for comparison
  let colsToRead = [destProjectNameCol];
  if (keyPairs.length > 0) {
    colsToRead = uniqueArray(colsToRead.concat(keyPairs.map(p => p.dest)));
  }
  const minCol = Math.min(...colsToRead);
  const maxCol = Math.max(...colsToRead);
  const readWidth = maxCol - minCol + 1;

  if (maxCol > destinationSheet.getMaxColumns()) {
    throw new ConfigurationError("Duplicate check failed: destination sheet missing expected columns for compound key.");
  }

  // 3. Read destination data in a batch, with retry
  const vals = withRetry(() => destinationSheet.getRange(2, minCol, lastDestRow - 1, readWidth).getValues(), {
    functionName: "findDuplicateRow:readDestinationData",
    correlationId: correlationId
  });

  const projIdx = destProjectNameCol - minCol;

  // 4. Scan destination data for the key
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    if (projIdx >= row.length) continue;
    let existingKey = normalizeString(row[projIdx]);
    if (!existingKey) continue;

    // Build the key from the destination row using the same sorted order
    for (const pair of keyPairs) {
      const destColIndex = pair.dest - minCol;
      const v = (destColIndex < row.length) ? row[destColIndex] : "";
      existingKey += sep + normalizeString(v);
    }

    if (existingKey === keyToCheck) {
      return i + 2; // Return 1-based row index
    }
  }

  return -1;
}

/**
 * Updates a specific row in the destination sheet with new data.
 * This helper function is called by `executeTransfer` when a sync operation is required.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The destination sheet object.
 * @param {number} rowIndex The 1-based row index to update.
 * @param {any[]} newRowData An array of values to write to the row.
 * @param {object} config The transfer configuration object, used for logging.
 * @param {string} correlationId A unique ID for tracing the operation.
 */
function updateRowInDestination(sheet, rowIndex, newRowData, config, correlationId) {
  const mapping = config.destinationColumnMapping || {};
  const destCols = Object.values(mapping).map(Number).filter(Number.isFinite);
  if (!destCols.length) return;
  const maxDestCol = Math.max(...destCols);

  // Read current row (only up to highest mapped col)
  const range = sheet.getRange(rowIndex, 1, 1, maxDestCol);
  const existing = withRetry(() => range.getValues(), {
    functionName: `${config.transferName}:readRowBeforeUpdate`,
    correlationId
  })[0];

  // Merge: only mapped destination columns are replaced
  const merged = existing.slice();
  for (const destCol of destCols) {
    const i = destCol - 1;
    const v = newRowData[i];
    if (typeof v !== "undefined") merged[i] = v;
  }

  withRetry(() => range.setValues([merged]), {
    functionName: `${config.transferName}:updateRow`,
    correlationId
  });
}