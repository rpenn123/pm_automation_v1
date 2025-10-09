/**
 * @OnlyCurrentDoc
 *
 * TransferEngine.gs
 *
 * This script provides a generic, reusable engine for transferring data between sheets based on a configuration object.
 * It is designed to be the robust backend for all data transfer operations, handling locking, duplicate checking,
 * flexible data mapping, and post-transfer actions like sorting.
 *
 * @version 1.0.0
 * @release 2024-07-29
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
 * @returns {void} This function does not return a value.
 */
function executeTransfer(e, config, preReadSourceRowData) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  let appendedRow = -1;
  const sourceSheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const ss = e.source;
  let projectName = ""; // Initialize early for use in error logging if possible

  try {
    // Attempt to acquire lock to prevent concurrent executions during data transfer
    lockAcquired = lock.tryLock(2500); // Slightly longer lock for transfers
    if (!lockAcquired) {
      Logger.log(`${config.transferName}: Lock not acquired. Skipping.`);
      logAudit(ss, {
        action: config.transferName,
        sourceSheet: sourceSheet.getName(),
        sourceRow: editedRow,
        result: "skipped-no-lock"
      }, CONFIG);
      return;
    }

    // Validate destination sheet
    const destinationSheet = ss.getSheetByName(config.destinationSheetName);
    if (!destinationSheet) {
      throw new Error(`Destination sheet "${config.destinationSheetName}" not found`);
    }

    // Read source data efficiently, or use pre-read data if available
    let sourceRowData = preReadSourceRowData;
    let readWidth;

    if (sourceRowData) {
      readWidth = sourceRowData.length;
    } else {
      const mappedSourceCols = Object.keys(config.destinationColumnMapping || {}).map(Number);
      const compoundKeyCols = (config.duplicateCheckConfig && config.duplicateCheckConfig.compoundKeySourceCols) || [];
      const maxSourceColNeeded = Math.max(...(config.sourceColumnsNeeded || []), ...mappedSourceCols, ...compoundKeyCols);
      const actualLastSourceCol = sourceSheet.getMaxColumns();
      readWidth = Math.min(maxSourceColNeeded, actualLastSourceCol);
      // Read the necessary part of the row in a single batch
      sourceRowData = sourceSheet.getRange(editedRow, 1, 1, readWidth).getValues()[0];
    }

    // Identify SFID and Project Name for logging and duplicate checks
    const { sfidSourceCol, projectNameSourceCol } = config.duplicateCheckConfig || {};
    let sfid = sfidSourceCol && sfidSourceCol <= readWidth ? sourceRowData[sfidSourceCol - 1] : null;
    projectName = projectNameSourceCol && projectNameSourceCol <= readWidth ? sourceRowData[projectNameSourceCol - 1] : "";
    sfid = sfid ? String(sfid).trim() : null;
    projectName = projectName ? String(projectName).trim() : "";
    
    // A record must have at least an SFID or a Project Name to be processed
    if (!sfid && !projectName) {
      Logger.log(`${config.transferName}: Skipping row ${editedRow}; missing both SFID and Project Name.`);
      logAudit(ss, {
        action: config.transferName,
        sourceSheet: sourceSheet.getName(),
        sourceRow: editedRow,
        details: "Missing both SFID and Project Name",
        result: "skipped"
      }, CONFIG);
      return;
    }

    // Perform Duplicate Check
    if (config.duplicateCheckConfig && config.duplicateCheckConfig.checkEnabled !== false) {
      if (isDuplicateInDestination(destinationSheet, sfid, projectName, sourceRowData, readWidth, config.duplicateCheckConfig)) {
        const logIdentifier = sfid ? `SFID ${sfid}` : `project "${projectName}"`;
        Logger.log(`${config.transferName}: Duplicate detected for ${logIdentifier}.`);
        logAudit(ss, {
          action: config.transferName,
          sourceSheet: sourceSheet.getName(),
          sourceRow: editedRow,
          sfid: sfid,
          projectName: projectName,
          details: "Duplicate",
          result: "skipped-duplicate"
        }, CONFIG);
        return;
      }
    }

    // Build the destination row
    const mapping = config.destinationColumnMapping || {};
    const maxMappedCol = getMaxValueInObject(mapping);
    // Determine the width of the new row
    const destLastCol = Math.max(destinationSheet.getMaxColumns(), maxMappedCol);
    const newRow = new Array(destLastCol).fill(""); // Initialize with empty strings

    for (const sourceColStr in mapping) {
      if (!Object.prototype.hasOwnProperty.call(mapping, sourceColStr)) continue;
      
      const sourceCol = Number(sourceColStr); // 1-indexed
      const destCol = mapping[sourceColStr];  // 1-indexed

      if (sourceCol <= readWidth) {
        // Access source array 0-indexed, place in destination array 0-indexed
        const value = sourceRowData[sourceCol - 1]; 
        newRow[destCol - 1] = (value !== null && value !== undefined) ? value : "";
      } else {
        Logger.log(`${config.transferName}: Warning: Source col ${sourceCol} not available in read data. Skipped mapping.`);
      }
    }

    // Append the row
    destinationSheet.appendRow(newRow);
    appendedRow = destinationSheet.getLastRow();
    
    // Update Last Edit tracking on the destination sheet if applicable
    if (config.lastEditTrackedSheets && config.lastEditTrackedSheets.includes(config.destinationSheetName)) {
        updateLastEditForRow(destinationSheet, appendedRow, CONFIG);
    }

    // Post Transfer Actions (e.g., Sorting)
    if (config.postTransferActions && config.postTransferActions.sort && appendedRow > 1) {
      SpreadsheetApp.flush(); // Ensure data is written before sorting
      try {
        const { sortColumn, sortAscending } = config.postTransferActions;
        // Get range excluding the header
        const range = destinationSheet.getRange(2, 1, appendedRow - 1, destinationSheet.getMaxColumns());
        range.sort({ column: sortColumn, ascending: !!sortAscending });
      } catch (sortError) {
        // Notify about sort failure, but the transfer itself succeeded
        notifyError(`${config.transferName} completed, but sorting failed`, sortError, ss, CONFIG);
      }
    }

    // Success Logging
    logAudit(ss, {
      action: config.transferName,
      sourceSheet: sourceSheet.getName(),
      sourceRow: editedRow,
      projectName: projectName,
      details: `Appended to ${config.destinationSheetName} row ${appendedRow}`,
      result: "success"
    }, CONFIG);

  } catch (error) {
    Logger.log(`${config.transferName} Error: ${error}\n${error.stack}`);
    notifyError(`${config.transferName} failed`, error, ss, CONFIG);
    logAudit(ss, {
      action: config.transferName,
      sourceSheet: sourceSheet.getName(),
      sourceRow: editedRow,
      projectName: projectName, // Use the initialized project name if available
      result: "error",
      errorMessage: String(error)
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
 * @param {number} [dupConfig.sfidDestCol] The 1-based column index for SFIDs in the destination sheet. Required for SFID check.
 * @param {number} dupConfig.projectNameDestCol The 1-based column index for project names in the destination sheet.
 * @param {number[]} [dupConfig.compoundKeySourceCols] An array of 1-based source column indices to build a compound key.
 * @param {number[]} [dupConfig.compoundKeyDestCols] The corresponding destination column indices for the compound key.
 * @param {string} [dupConfig.keySeparator="|"] The separator used when concatenating values for the compound key.
 * @returns {boolean} Returns `true` if a duplicate is found, otherwise `false`.
 */
function isDuplicateInDestination(destinationSheet, sfid, projectName, sourceRowData, sourceReadWidth, dupConfig) {
  // Strategy 1: SFID is the primary, definitive check.
  if (sfid && dupConfig.sfidDestCol) {
    // Use the efficient, exact-match lookup utility. A match here is a definitive duplicate.
    return findRowByValue(destinationSheet, sfid, dupConfig.sfidDestCol) !== -1;
  }

  // Strategy 2: Fallback to Project Name + Compound Key if no SFID is present.
  // This maintains backward compatibility with legacy data.
  if (!projectName) {
    // Cannot perform fallback check without a project name.
    return false;
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
    Logger.log("Duplicate check warning: destination sheet missing expected columns for compound key. Check may be incomplete.");
  }

  // 3. Read destination data in a batch
  const range = destinationSheet.getRange(2, minCol, lastDestRow - 1, readWidth);
  const vals = range.getValues();
  const projIdx = destProjectNameCol - minCol; // Relative index for project name

  // 4. Scan destination data for the key
  for (const row of vals) {
    if (projIdx >= row.length) continue;
    let existingKey = formatValueForKey(row[projIdx]); // Use consistent formatting
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