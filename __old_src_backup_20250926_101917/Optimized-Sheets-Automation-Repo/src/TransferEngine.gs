/**
 * @OnlyCurrentDoc
 * TransferEngine.gs
 * Generic, reusable engine for transferring data between sheets based on configuration.
 * Handles locking, duplicate checking, data mapping, and post-transfer actions.
 */

/**
 * Executes a generic, configuration-driven data transfer from a source row to a destination sheet.
 * This function is the core of the transfer mechanism, designed to be called by trigger handlers in `Automations.gs`.
 * It handles script locking, data reading, duplicate checking, column mapping, and optional post-transfer actions like sorting.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object passed from the trigger.
 * @param {object} config The configuration object that defines the transfer.
 * @param {string} config.transferName A descriptive name for the transfer, used for logging.
 * @param {string} config.destinationSheetName The name of the sheet to which data will be transferred.
 * @param {number[]} config.sourceColumnsNeeded An array of 1-based column indices required from the source row.
 * @param {Object<number, number>} config.destinationColumnMapping An object mapping source column indices to destination column indices.
 * @param {object} [config.duplicateCheckConfig] Optional configuration for preventing duplicate entries.
 * @param {object} [config.postTransferActions] Optional configuration for actions to perform after a successful transfer, such as sorting.
 */
function executeTransfer(e, config) {
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
      });
      return;
    }

    // Validate destination sheet
    const destinationSheet = ss.getSheetByName(config.destinationSheetName);
    if (!destinationSheet) {
      throw new Error(`Destination sheet "${config.destinationSheetName}" not found`);
    }

    // Read source data efficiently
    const maxSourceColNeeded = Math.max(...config.sourceColumnsNeeded);
    const actualLastSourceCol = sourceSheet.getLastColumn();
    const readWidth = Math.min(maxSourceColNeeded, actualLastSourceCol);
    // Read the necessary part of the row in a single batch
    const sourceRowData = sourceSheet.getRange(editedRow, 1, 1, readWidth).getValues()[0];

    // Identify Project Name (Required for logging and duplicate checks)
    const FC_PN_COL = CONFIG.FORECASTING_COLS.PROJECT_NAME;
    const projectNameColIndex = (config.duplicateCheckConfig && config.duplicateCheckConfig.projectNameSourceCol) || FC_PN_COL;
    
    // Access array using 0-based index (ColIndex - 1)
    if (projectNameColIndex <= readWidth && sourceRowData[projectNameColIndex - 1]) {
        projectName = String(sourceRowData[projectNameColIndex - 1]).trim();
    }

    if (!projectName) {
      Logger.log(`${config.transferName}: Skipping row ${editedRow}; missing Project Name.`);
      logAudit(ss, {
        action: config.transferName,
        sourceSheet: sourceSheet.getName(),
        sourceRow: editedRow,
        details: "Missing Project Name",
        result: "skipped"
      });
      return;
    }

    // Perform Duplicate Check
    if (config.duplicateCheckConfig && config.duplicateCheckConfig.checkEnabled !== false) {
      if (isDuplicateInDestination(destinationSheet, projectName, sourceRowData, readWidth, config.duplicateCheckConfig)) {
        Logger.log(`${config.transferName}: Duplicate detected for "${projectName}".`);
        logAudit(ss, {
          action: config.transferName,
          sourceSheet: sourceSheet.getName(),
          sourceRow: editedRow,
          projectName: projectName,
          details: "Duplicate",
          result: "skipped-duplicate"
        });
        return;
      }
    }

    // Build the destination row
    const mapping = config.destinationColumnMapping || {};
    const maxMappedCol = getMaxValueInObject(mapping);
    // Determine the width of the new row
    const destLastCol = Math.max(destinationSheet.getLastColumn(), maxMappedCol);
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
    if (CONFIG.LAST_EDIT.TRACKED_SHEETS.includes(config.destinationSheetName)) {
        updateLastEditForRow(destinationSheet, appendedRow);
    }

    // Post Transfer Actions (e.g., Sorting)
    if (config.postTransferActions && config.postTransferActions.sort && appendedRow > 1) {
      SpreadsheetApp.flush(); // Ensure data is written before sorting
      try {
        const { sortColumn, sortAscending } = config.postTransferActions;
        // Get range excluding the header
        const range = destinationSheet.getRange(2, 1, appendedRow - 1, destinationSheet.getLastColumn());
        range.sort({ column: sortColumn, ascending: !!sortAscending });
      } catch (sortError) {
        // Notify about sort failure, but the transfer itself succeeded
        notifyError(`${config.transferName} completed, but sorting failed`, sortError, ss);
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
    });

  } catch (error) {
    Logger.log(`${config.transferName} Error: ${error}\n${error.stack}`);
    notifyError(`${config.transferName} failed`, error, ss);
    logAudit(ss, {
      action: config.transferName,
      sourceSheet: sourceSheet.getName(),
      sourceRow: editedRow,
      projectName: projectName, // Use the initialized project name if available
      result: "error",
      errorMessage: String(error)
    });
  } finally {
    // Ensure the lock is always released
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Checks for a duplicate entry in the destination sheet based on a flexible keying strategy.
 * It can check for duplicates based on a simple project name or a compound key (e.g., Project Name + Deadline).
 * This function builds a unique key from the source row data and efficiently scans the destination sheet for a matching key.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} destinationSheet The sheet to scan for duplicates.
 * @param {string} projectName The project name from the source row.
 * @param {any[]} sourceRowData The array of values from the source row.
 * @param {number} sourceReadWidth The number of columns read from the source sheet to ensure no out-of-bounds access.
 * @param {object} dupConfig The configuration for the duplicate check.
 * @param {number} dupConfig.projectNameDestCol The 1-based column index of the project name in the destination sheet.
 * @param {number[]} [dupConfig.compoundKeySourceCols] Optional array of 1-based source column indices for a compound key.
 * @param {number[]} [dupConfig.compoundKeyDestCols] Optional array of 1-based destination column indices for a compound key.
 * @param {string} [dupConfig.keySeparator="|"] The separator character used for building compound keys.
 * @returns {boolean} True if a duplicate key is found in the destination sheet, otherwise false.
 */
function isDuplicateInDestination(destinationSheet, projectName, sourceRowData, sourceReadWidth, dupConfig) {
  const destProjectNameCol = dupConfig.projectNameDestCol;
  const lastDestRow = destinationSheet.getLastRow();
  if (lastDestRow < 2) return false; // No data rows exist

  // 1. Build the key we are checking against (from the source data)
  let keyToCheck = projectName.trim().toLowerCase();
  const sep = dupConfig.keySeparator || "|";

  // Add compound keys if configured and lengths match
  if (dupConfig.compoundKeySourceCols && dupConfig.compoundKeyDestCols && dupConfig.compoundKeySourceCols.length === dupConfig.compoundKeyDestCols.length) {
    for (let i = 0; i < dupConfig.compoundKeySourceCols.length; i++) {
      const sourceCol = dupConfig.compoundKeySourceCols[i];
      // Access source data array 0-indexed (Col - 1)
      const val = (sourceCol <= sourceReadWidth) ? sourceRowData[sourceCol - 1] : undefined;
      keyToCheck += sep + formatValueForKey(val);
    }
  }

  // 2. Determine the columns needed from the destination sheet for comparison
  let cols = [destProjectNameCol];
  if (dupConfig.compoundKeyDestCols && dupConfig.compoundKeyDestCols.length) {
    cols = uniqueArray(cols.concat(dupConfig.compoundKeyDestCols));
  }
  const minCol = Math.min(...cols);
  const maxCol = Math.max(...cols);
  const readWidth = maxCol - minCol + 1;

  if (maxCol > destinationSheet.getLastColumn()) {
    Logger.log("Duplicate check warning: destination sheet missing expected columns for compound key. Check may be incomplete.");
  }

  // 3. Read destination data efficiently in a batch
  const range = destinationSheet.getRange(2, minCol, lastDestRow - 1, readWidth);
  const vals = range.getValues();

  // Calculate relative indices for accessing the read data array (relative to minCol)
  const projIdx = destProjectNameCol - minCol;
  const cmpIdxs = (dupConfig.compoundKeyDestCols || []).map(col => col - minCol);

  // 4. Scan destination data for the key
  for (const row of vals) {
    if (projIdx >= row.length) continue; // Handle ragged rows
    
    let existingKey = row[projIdx] ? String(row[projIdx]).trim().toLowerCase() : "";
    if (!existingKey) continue;

    // Build the key from the destination row
    for (const idx of cmpIdxs) {
      const v = (idx < row.length) ? row[idx] : "";
      existingKey += sep + formatValueForKey(v);
    }

    // Compare keys
    if (existingKey === keyToCheck) return true;
  }

  return false;
}