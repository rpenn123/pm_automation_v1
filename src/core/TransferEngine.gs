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
  let appendedRow = -1;
  const sourceSheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const ss = e.source;
  let projectName = ""; // Initialize early for use in error logging

  try {
    // Attempt to acquire lock with retry logic. Throws TransientError on failure.
    withRetry(() => {
      lockAcquired = lock.tryLock(2500);
      if (!lockAcquired) throw new Error("Lock not acquired within the time limit.");
    }, { functionName: `${config.transferName}:acquireLock`, maxRetries: 2, initialDelayMs: 500 });

    // Validate destination sheet
    const destinationSheet = ss.getSheetByName(config.destinationSheetName);
    if (!destinationSheet) {
      throw new ConfigurationError(`Destination sheet "${config.destinationSheetName}" not found`);
    }

    // Read source data efficiently, or use pre-read data if available
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

      // Read the necessary part of the row in a single batch, with retry
      sourceRowData = withRetry(
        () => sourceSheet.getRange(editedRow, 1, 1, readWidth).getValues()[0],
        { functionName: `${config.transferName}:readSourceRow` }
      );
    }

    // Identify SFID and Project Name for logging and duplicate checks
    const { sfidSourceCol, projectNameSourceCol } = config.duplicateCheckConfig || {};
    let sfid = sfidSourceCol && sfidSourceCol <= readWidth ? sourceRowData[sfidSourceCol - 1] : null;
    projectName = projectNameSourceCol && projectNameSourceCol <= readWidth ? sourceRowData[projectNameSourceCol - 1] : "";
    sfid = sfid ? String(sfid).trim() : null;
    projectName = projectName ? String(projectName).trim() : "";
    
    // A record must have at least an SFID or a Project Name to be processed. Throws ValidationError on failure.
    if (!sfid && !projectName) {
      throw new ValidationError(`Row ${editedRow} is missing both SFID and Project Name.`);
    }

    // Perform Duplicate Check
    if (config.duplicateCheckConfig && config.duplicateCheckConfig.checkEnabled !== false) {
      if (isDuplicateInDestination(destinationSheet, sfid, projectName, sourceRowData, readWidth, config.duplicateCheckConfig, correlationId)) {
        const logIdentifier = sfid ? `SFID ${sfid}` : `project "${projectName}"`;
        logAudit(ss, {
          correlationId: correlationId,
          action: config.transferName,
          sourceSheet: sourceSheet.getName(),
          sourceRow: editedRow,
          sfid: sfid,
          projectName: projectName,
          details: `Duplicate detected for ${logIdentifier}.`,
          result: "skipped-duplicate"
        }, CONFIG);
        return; // This is an expected outcome, not an error.
      }
    }

    // Build the destination row
    const mapping = config.destinationColumnMapping || {};
    const maxMappedCol = getMaxValueInObject(mapping);
    const destLastCol = Math.max(destinationSheet.getMaxColumns(), maxMappedCol);
    const newRow = new Array(destLastCol).fill("");

    for (const sourceColStr in mapping) {
      if (!Object.prototype.hasOwnProperty.call(mapping, sourceColStr)) continue;
      
      const sourceCol = Number(sourceColStr);
      const destCol = mapping[sourceColStr];

      if (sourceCol <= readWidth) {
        const value = sourceRowData[sourceCol - 1]; 
        newRow[destCol - 1] = (value !== null && value !== undefined) ? value : "";
      } else {
        // This is a configuration issue, but may not be critical. Log as a warning.
        handleError(new ConfigurationError(`Source col ${sourceCol} not available in read data. Skipped mapping.`), {
            correlationId, functionName: "executeTransfer", spreadsheet: ss,
            extra: { transferName: config.transferName, sourceSheet: sourceSheet.getName(), editedRow }
        }, CONFIG);
      }
    }

    // Append the row, with retry
    withRetry(() => destinationSheet.appendRow(newRow), { functionName: `${config.transferName}:appendRow` });
    appendedRow = destinationSheet.getLastRow();
    
    // Update Last Edit tracking
    if (config.lastEditTrackedSheets && config.lastEditTrackedSheets.includes(config.destinationSheetName)) {
        updateLastEditForRow(destinationSheet, appendedRow, CONFIG);
    }

    // Post Transfer Actions (e.g., Sorting)
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

    // Success Logging
    logAudit(ss, {
      correlationId: correlationId,
      action: config.transferName,
      sourceSheet: sourceSheet.getName(),
      sourceRow: editedRow,
      projectName: projectName,
      details: `Appended to ${config.destinationSheetName} row ${appendedRow}`,
      result: "success"
    }, CONFIG);

  } catch (error) {
    // Centralized error handling
    handleError(error, {
        correlationId,
        functionName: "executeTransfer",
        spreadsheet: ss,
        extra: { transferName: config.transferName, sourceSheet: sourceSheet.getName(), editedRow, projectName }
    }, CONFIG);

    // Audit the failure
    logAudit(ss, {
      correlationId: correlationId,
      action: config.transferName,
      sourceSheet: sourceSheet.getName(),
      sourceRow: editedRow,
      projectName: projectName,
      result: "error",
      errorMessage: `${error.name}: ${error.message}`
    }, CONFIG);

  } finally {
    // Ensure the lock is always released
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Performs a robust duplicate check in the destination sheet using an SFID-first strategy.
 * This helper function for `executeTransfer` is critical for data integrity. It uses two main strategies:
 * 1.  **SFID Check (Primary):** If a non-empty `sfid` is provided, it performs a fast, exact-match search on the destination SFID column.
 * 2.  **Compound Key Check (Fallback):** If no SFID is available, it constructs a unique key from the Project Name plus any additional columns
 *     defined in `compoundKeySourceCols`. This maintains compatibility with legacy data that may not have an SFID.
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
  let keyToCheck = formatValueForKey(projectName);
  for (const pair of keyPairs) {
    const val = (pair.source <= sourceReadWidth) ? sourceRowData[pair.source - 1] : undefined;
    keyToCheck += sep + formatValueForKey(val);
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
    let existingKey = formatValueForKey(row[projIdx]);
    if (!existingKey) continue;

    // Build the key from the destination row using the same sorted order
    for (const pair of keyPairs) {
      const destColIndex = pair.dest - minCol;
      const v = (destColIndex < row.length) ? row[destColIndex] : "";
      existingKey += sep + formatValueForKey(v);
    }

    if (existingKey === keyToCheck) return true;
  }

  return false;
}