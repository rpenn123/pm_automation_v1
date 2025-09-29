/**
 * @OnlyCurrentDoc
 * Automations.gs
 * Main onEdit trigger entry point, synchronization logic, and transfer definitions.
 */

/**
 * The main `onEdit` function, which serves as the entry point for all sheet automations.
 * It is designed to be installed as an installable trigger via `Setup.gs`.
 * The function includes performance optimizations to exit early for irrelevant edits.
 * It uses a rule-based routing system to delegate the event to the appropriate handler
 * based on the sheet and column that were edited.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object passed by the onEdit trigger.
 * @returns {void}
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  // Performance optimizations: exit early for multi-cell edits, header edits, or non-value changes.
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  if (e.range.getRow() <= 1) return;
  // Prevent loops or unnecessary processing if the normalized value hasn't changed
  if (normalizeForComparison(e.value) === normalizeForComparison(e.oldValue)) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();

  // 1. Always update Last Edit tracking if applicable
  // Check if the sheet name is in the configured tracked sheets
  if (CONFIG.LAST_EDIT.TRACKED_SHEETS.includes(sheetName)) {
    // Errors during Last Edit update are handled within the service function
    updateLastEditForRow(sheet, editedRow);
  }

  // 2. Route the edit event to the appropriate handler
  const { FORECASTING, UPCOMING } = CONFIG.SHEETS;
  const FC = CONFIG.FORECASTING_COLS;
  const UP = CONFIG.UPCOMING_COLS;
  const STATUS = CONFIG.STATUS_STRINGS;

  // Define routing rules (order matters for combined actions like Sync+Transfer)
  const rules = [
    // --- Sync Rules ---
    { 
      // Forecasting Progress Edit: Sync to Upcoming AND potentially trigger Framing
      sheet: FORECASTING, 
      col: FC.PROGRESS, 
      handler: handleSyncAndPotentialFramingTransfer 
    },
    { 
      // Upcoming Progress Edit: Sync back to Forecasting
      sheet: UPCOMING,    
      col: UP.PROGRESS, 
      handler: triggerSyncToForecasting 
    },

    // --- Transfer Rules (Specific Value Checks) ---
    { 
      // Forecasting Permits "approved": Transfer to Upcoming
      sheet: FORECASTING, 
      col: FC.PERMITS,
      valueCheck: (val) => normalizeString(val) === STATUS.PERMIT_APPROVED.toLowerCase(),
      handler: triggerUpcomingTransfer 
    },
    { 
      // Forecasting Delivered TRUE: Transfer to Inventory
      sheet: FORECASTING, 
      col: FC.DELIVERED,
      valueCheck: (val) => isTrueLike(val),
      handler: triggerInventoryTransfer 
    },

    // Fallback framing trigger (ensures framing runs if "In Progress" is set, even if the combined handler somehow missed it)
    { 
      sheet: FORECASTING, 
      col: FC.PROGRESS,
      valueCheck: (val) => normalizeString(val) === STATUS.IN_PROGRESS.toLowerCase(),
      handler: triggerFramingTransfer 
    }
  ];

  // Execute the first matching rule
  for (const rule of rules) {
    if (sheetName === rule.sheet && editedCol === rule.col) {
      if (!rule.valueCheck || rule.valueCheck(e.value)) {
        try {
          rule.handler(e);
        } catch (error) {
          // Centralized error handling for failed handlers
          Logger.log(`Handler error for ${rule.sheet} Col ${rule.col}: ${error}\n${error.stack}`);
          notifyError(`Handler failed for ${rule.sheet} Col ${rule.col}`, error, e.source);
          logAudit(e.source, {
            action: "HandlerError",
            sourceSheet: sheetName,
            sourceRow: editedRow,
            details: `Col ${editedCol}`,
            result: "error",
            errorMessage: String(error)
          });
        }
        // Stop after the first successful handler execution
        return;
      }
    }
  }
}

// =================================================================
// ==================== SYNC HANDLERS ==============================
// =================================================================

/**
 * Handles an edit to the 'Progress' column in the 'Forecasting' sheet.
 * This is a composite handler that performs two actions:
 * 1. Synchronizes the new 'Progress' value to the corresponding project in the 'Upcoming' sheet.
 * 2. If the new value is "In Progress", it triggers a data transfer to the 'Framing' sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function handleSyncAndPotentialFramingTransfer(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const FC = CONFIG.FORECASTING_COLS;

  // Read both SFID and Project Name from the source row
  const sfid = sheet.getRange(editedRow, FC.SFID).getValue();
  const projectName = sheet.getRange(editedRow, FC.PROJECT_NAME).getValue();
  const newValue = e.value;

  // 1. Synchronization (ensure at least one identifier exists)
  if (sfid || projectName) {
    syncProgressToUpcoming(sfid, projectName, newValue, e.source, e);
  }
  
  // 2. Conditional Transfer
  if (normalizeString(newValue) === CONFIG.STATUS_STRINGS.IN_PROGRESS.toLowerCase()) {
    triggerFramingTransfer(e);
  }
}

/**
 * Handles an edit to the 'Progress' column in the 'Upcoming' sheet.
 * This handler synchronizes the new 'Progress' value back to the
 * corresponding project in the 'Forecasting' sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function triggerSyncToForecasting(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const UP = CONFIG.UPCOMING_COLS;

  // Read both SFID and Project Name from the source row
  const sfid = sheet.getRange(editedRow, UP.SFID).getValue();
  const projectName = sheet.getRange(editedRow, UP.PROJECT_NAME).getValue();
  
  if (sfid || projectName) {
    syncProgressToForecasting(sfid, projectName, e.value, e.source, e);
  }
}

/**
 * Synchronizes 'Progress' from 'Forecasting' to 'Upcoming' using an SFID-first strategy.
 * It uses a script lock and a value check to prevent infinite loops.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project to sync (fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event object for logging context.
 */
function syncProgressToUpcoming(sfid, projectName, newValue, ss, eCtx) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { UPCOMING } = CONFIG.SHEETS;
  const UP = CONFIG.UPCOMING_COLS;
  const actionName = "SyncFtoU";
  const logIdentifier = sfid ? `SFID ${sfid}` : `"${projectName}"`;

  try {
    lockAcquired = lock.tryLock(2000);
    if (!lockAcquired) {
      Logger.log(`${actionName}: Lock not acquired for ${logIdentifier}. Skipping.`);
      logAudit(ss, { action: `${actionName}-SkippedNoLock`, sourceSheet: eCtx.range.getSheet().getName(), sourceRow: eCtx.range.getRow(), sfid: sfid, projectName: projectName, result: "skipped" });
      return;
    }

    const upcomingSheet = ss.getSheetByName(UPCOMING);
    if (!upcomingSheet) throw new Error(`Destination sheet "${UPCOMING}" not found`);

    // Use the new SFID-first lookup function
    const row = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);
    
    if (row !== -1) {
      const targetCell = upcomingSheet.getRange(row, UP.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      if (currentValueStr !== newValueStr) {
        targetCell.setValue(newValue);
        updateLastEditForRow(upcomingSheet, row);
        SpreadsheetApp.flush();
        logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress ${currentValueStr} -> ${newValueStr}`, result: "updated" });
      } else {
        logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid: sfid, projectName: projectName, details: "No change", result: "noop" });
      }
    } else {
      logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sfid: sfid, projectName: projectName, details: "Project not found", result: "miss" });
    }
  } catch (error) {
    Logger.log(`${actionName} error for ${logIdentifier}: ${error}\n${error.stack}`);
    notifyError(`${actionName} failed for project ${logIdentifier}`, error, ss);
    logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sfid: sfid, projectName: projectName, result: "error", errorMessage: String(error) });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Synchronizes 'Progress' from 'Upcoming' back to 'Forecasting' using an SFID-first strategy.
 * This completes the two-way sync, preventing infinite loops.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project to sync (fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event object for logging context.
 */
function syncProgressToForecasting(sfid, projectName, newValue, ss, eCtx) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { FORECASTING } = CONFIG.SHEETS;
  const FC = CONFIG.FORECASTING_COLS;
  const actionName = "SyncUtoF";
  const logIdentifier = sfid ? `SFID ${sfid}` : `"${projectName}"`;

  try {
    lockAcquired = lock.tryLock(2000);
    if (!lockAcquired) {
      Logger.log(`${actionName}: Lock not acquired for ${logIdentifier}. Skipping.`);
      logAudit(ss, { action: `${actionName}-SkippedNoLock`, sourceSheet: eCtx.range.getSheet().getName(), sourceRow: eCtx.range.getRow(), sfid: sfid, projectName: projectName, result: "skipped" });
      return;
    }

    const forecastingSheet = ss.getSheetByName(FORECASTING);
    if (!forecastingSheet) throw new Error(`Destination sheet "${FORECASTING}" not found`);

    // Use the new SFID-first lookup function
    const row = findRowByBestIdentifier(forecastingSheet, sfid, FC.SFID, projectName, FC.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = forecastingSheet.getRange(row, FC.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      if (currentValueStr !== newValueStr) {
        targetCell.setValue(newValue);
        updateLastEditForRow(forecastingSheet, row);
        SpreadsheetApp.flush();
        logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress ${currentValueStr} -> ${newValueStr}`, result: "updated" });
      } else {
        logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid: sfid, projectName: projectName, details: "No change", result: "noop" });
      }
    } else {
      logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sfid: sfid, projectName: projectName, details: "Project not found", result: "miss" });
    }
  } catch (error) {
    Logger.log(`${actionName} error for ${logIdentifier}: ${error}\n${error.stack}`);
    notifyError(`${actionName} failed for project ${logIdentifier}`, error, ss);
    logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sfid: sfid, projectName: projectName, result: "error", errorMessage: String(error) });
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// =================================================================
// ==================== TRANSFER DEFINITIONS =======================
// =================================================================
// These functions define the configuration for specific transfers and pass them to the generic TransferEngine.

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Upcoming'.
 * This transfer is initiated when the 'Permits' column in 'Forecasting' is set to "approved".
 * It constructs a configuration object and passes it to the `executeTransfer` engine.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function triggerUpcomingTransfer(e) {
  const FC = CONFIG.FORECASTING_COLS;
  const UP = CONFIG.UPCOMING_COLS;

  const config = {
    transferName: "Upcoming Transfer (Permits)",
    destinationSheetName: CONFIG.SHEETS.UPCOMING,
    // List of required source column indices (1-based)
    sourceColumnsNeeded: [
      FC.SFID, FC.PROJECT_NAME, FC.EQUIPMENT, FC.PROGRESS,
      FC.PERMITS, FC.DEADLINE, FC.LOCATION
    ],
    // Mapping: { [Source_COL_1]: Dest_COL_1 }
    destinationColumnMapping: createMapping([
      [FC.SFID,         UP.SFID],
      [FC.PROJECT_NAME, UP.PROJECT_NAME],
      [FC.DEADLINE,     UP.DEADLINE],
      [FC.PROGRESS,     UP.PROGRESS],
      [FC.EQUIPMENT,    UP.EQUIPMENT],
      [FC.PERMITS,      UP.PERMITS],
      [FC.LOCATION,     UP.LOCATION]
      // Construction Start, Install date, Construction, Notes intentionally left blank
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      // SFID is now the primary check, with Project Name as fallback
      sfidSourceCol: FC.SFID,
      sfidDestCol: UP.SFID,
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: UP.PROJECT_NAME
    },
    postTransferActions: {
      sort: true,
      sortColumn: UP.DEADLINE,
      sortAscending: false // Sort by deadline descending
    }
  };
  executeTransfer(e, config);
}

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Inventory_Elevators'.
 * This transfer is initiated when the 'Delivered' column in 'Forecasting' is checked (TRUE).
 * It constructs a configuration object and passes it to the `executeTransfer` engine.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function triggerInventoryTransfer(e) {
  const FC = CONFIG.FORECASTING_COLS;
  const INV = CONFIG.INVENTORY_COLS;

  const config = {
    transferName: "Inventory Transfer (Delivered)",
    destinationSheetName: CONFIG.SHEETS.INVENTORY,
    sourceColumnsNeeded: [ FC.PROJECT_NAME, FC.DETAILS, FC.EQUIPMENT, FC.PROGRESS ],
    destinationColumnMapping: createMapping([
      [FC.PROJECT_NAME, INV.PROJECT_NAME],
      [FC.PROGRESS,     INV.PROGRESS],
      [FC.EQUIPMENT,    INV.EQUIPMENT],
      [FC.DETAILS,      INV.DETAILS]
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: INV.PROJECT_NAME
    }
  };
  executeTransfer(e, config);
}

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Framing'.
 * This transfer is initiated when the 'Progress' column in 'Forecasting' is set to "In Progress".
 * It uses a compound key for duplicate checking to allow the same project to be added
 * if its deadline changes.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function triggerFramingTransfer(e) {
  const FC = CONFIG.FORECASTING_COLS;
  const FR = CONFIG.FRAMING_COLS;

  const config = {
    transferName: "Framing Transfer (In Progress)",
    destinationSheetName: CONFIG.SHEETS.FRAMING,
    sourceColumnsNeeded: [ FC.SFID, FC.PROJECT_NAME, FC.EQUIPMENT, FC.ARCHITECT, FC.DEADLINE ],
    destinationColumnMapping: createMapping([
      [FC.SFID,         FR.SFID],
      [FC.PROJECT_NAME, FR.PROJECT_NAME],
      [FC.EQUIPMENT,    FR.EQUIPMENT],
      [FC.ARCHITECT,    FR.ARCHITECT],
      [FC.DEADLINE,     FR.DEADLINE]
    ]),
    // Use SFID as primary duplicate check, falling back to compound key (Project Name + Deadline)
    duplicateCheckConfig: {
      checkEnabled: true,
      sfidSourceCol: FC.SFID,
      sfidDestCol: FR.SFID,
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: FR.PROJECT_NAME,
      compoundKeySourceCols: [FC.DEADLINE],
      compoundKeyDestCols: [FR.DEADLINE],
      keySeparator: "|"
    }
  };
  executeTransfer(e, config);
}