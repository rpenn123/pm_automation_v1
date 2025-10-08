/**
 * @OnlyCurrentDoc
 * Automations.gs
 * This file contains the main `onEdit` trigger entry point, the core data synchronization logic,
 * and the definitions for all automated data transfers between sheets.
 */

/**
 * The main `onEdit` function, serving as the central hub for all sheet automations.
 * This function is installed as an installable trigger by `Setup.gs`.
 *
 * **Execution Flow:**
 * 1.  **Guard Clauses:** Performs several checks to exit early for irrelevant edits.
 * 2.  **Batch Row Read:** Reads the entire edited row in a single operation for performance.
 * 3.  **Last Edit Tracking:** Calls the `LastEditService` to update the timestamp for the modified row.
 * 4.  **Rule-Based Routing:** Uses a `rules` array to find the appropriate handler based on the edited sheet and column.
 * 5.  **Handler Execution:** Executes the first matching handler, passing the pre-read row data to it.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object passed by the onEdit trigger.
 * @returns {void}
 */
function onEdit(e) {
  if (!e || !e.range) return;

  // Performance optimizations: exit early for multi-cell edits, header edits, or non-value changes.
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  if (e.range.getRow() <= 1) return;
  if (normalizeForComparison(e.value) === normalizeForComparison(e.oldValue)) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();

  // The only function that should access the global CONFIG.
  const config = CONFIG;

  // 1. Batch-read the entire edited row once for performance.
  const sourceRowData = sheet.getRange(editedRow, 1, 1, sheet.getMaxColumns()).getValues()[0];

  // 2. Always update Last Edit tracking if applicable. This runs independently of other rules.
  if (config.LAST_EDIT.TRACKED_SHEETS.includes(sheetName)) {
    updateLastEditForRow(sheet, editedRow, config);
    // Also log this edit to the audit trail for accountability.
    try {
      const a1Notation = e.range.getA1Notation ? e.range.getA1Notation() : 'unknown';
      logAudit(e.source, {
        action: 'Row Edit',
        sourceSheet: sheetName,
        sourceRow: editedRow,
        details: `Cell ${a1Notation} updated. New value: "${e.value}"`,
        result: 'success'
      }, config);
    } catch (logError) {
      Logger.log(`Failed to write audit log for edit on ${sheetName} R${editedRow}: ${logError}`);
    }
  }

  // 3. Route the edit event to the appropriate handler based on a set of rules.
  const { FORECASTING, UPCOMING } = config.SHEETS;
  const FC = config.FORECASTING_COLS;
  const UP = config.UPCOMING_COLS;
  const STATUS = config.STATUS_STRINGS;

  const rules = [
    { sheet: FORECASTING, col: FC.PROGRESS, handler: handleSyncAndPotentialFramingTransfer },
    { sheet: UPCOMING, col: UP.PROGRESS, handler: triggerSyncToForecasting },
    { sheet: FORECASTING, col: FC.PERMITS, valueCheck: (val) => normalizeString(val) === STATUS.PERMIT_APPROVED.toLowerCase(), handler: triggerUpcomingTransfer },
    { sheet: FORECASTING, col: FC.DELIVERED, valueCheck: (val) => isTrueLike(val), handler: triggerInventoryTransfer },
  ];

  // Execute the first matching rule.
  for (const rule of rules) {
    if (sheetName === rule.sheet && editedCol === rule.col) {
      if (!rule.valueCheck || rule.valueCheck(e.value)) {
        try {
          // Pass the pre-read row data and the config object to the handler.
          rule.handler(e, sourceRowData, config);
        } catch (error) {
          Logger.log(`Handler error for ${rule.sheet} Col ${rule.col}: ${error}\n${error.stack}`);
          notifyError(`Handler failed for ${rule.sheet} Col ${rule.col}`, error, e.source, config);
          logAudit(e.source, {
            action: "HandlerError",
            sourceSheet: sheetName,
            sourceRow: editedRow,
            details: `Col ${editedCol}`,
            result: "error",
            errorMessage: String(error)
          }, config);
        }
        return; // Stop after the first matching rule.
      }
    }
  }
}

// =================================================================
// ==================== SYNC HANDLERS ==============================
// =================================================================

/**
 * Handles an edit to 'Progress' in 'Forecasting'. It always syncs the value to 'Upcoming'
 * and conditionally triggers a transfer to 'Framing' if the new value is "In Progress".
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function handleSyncAndPotentialFramingTransfer(e, sourceRowData, config) {
  const FC = config.FORECASTING_COLS;

  // Use the pre-read data instead of making new I/O calls.
  const sfid = sourceRowData[FC.SFID - 1];
  const projectName = sourceRowData[FC.PROJECT_NAME - 1];
  const newValue = e.value;

  // 1. Synchronization.
  if (sfid || projectName) {
    syncProgressToUpcoming(sfid, projectName, newValue, e.source, e, config);
  }

  // 2. Conditional Transfer to Framing.
  if (normalizeString(newValue) === config.STATUS_STRINGS.IN_PROGRESS.toLowerCase()) {
    triggerFramingTransfer(e, sourceRowData, config);
  }
}

/**
 * Handles an edit to 'Progress' in 'Upcoming', syncing the value back to 'Forecasting'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function triggerSyncToForecasting(e, sourceRowData, config) {
  const UP = config.UPCOMING_COLS;

  // Use the pre-read data.
  const sfid = sourceRowData[UP.SFID - 1];
  const projectName = sourceRowData[UP.PROJECT_NAME - 1];

  if (sfid || projectName) {
    syncProgressToForecasting(sfid, projectName, e.value, e.source, e, config);
  }
}

/**
 * Synchronizes the 'Progress' value from 'Forecasting' to 'Upcoming'.
 * Uses a script lock to prevent race conditions.
 *
 * @param {string} sfid The Salesforce ID of the project.
 * @param {string} projectName The name of the project (fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event for logging.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function syncProgressToUpcoming(sfid, projectName, newValue, ss, eCtx, config) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { UPCOMING } = config.SHEETS;
  const UP = config.UPCOMING_COLS;
  const actionName = "SyncFtoU";
  const logIdentifier = sfid ? `SFID ${sfid}` : `"${projectName}"`;

  try {
    lockAcquired = acquireLockWithRetry(lock);
    if (!lockAcquired) {
      Logger.log(`${actionName}: Lock not acquired for ${logIdentifier} after multiple retries. Skipping.`);
      logAudit(ss, { action: `${actionName}-SkippedNoLock`, sourceSheet: eCtx.range.getSheet().getName(), sourceRow: eCtx.range.getRow(), sfid: sfid, projectName: projectName, result: "skipped" }, config);
      return;
    }

    const upcomingSheet = ss.getSheetByName(UPCOMING);
    if (!upcomingSheet) throw new Error(`Destination sheet "${UPCOMING}" not found`);

    const row = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = upcomingSheet.getRange(row, UP.PROGRESS);
      if (normalizeForComparison(targetCell.getValue()) !== normalizeForComparison(newValue)) {
        targetCell.setValue(newValue);
        updateLastEditForRow(upcomingSheet, row, config);
        logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress -> ${newValue}`, result: "updated" }, config);
      } else {
        logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid: sfid, projectName: projectName, details: "No change", result: "noop" }, config);
      }
    } else {
      logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sfid: sfid, projectName: projectName, details: "Project not found", result: "miss" }, config);
    }
  } catch (error) {
    Logger.log(`${actionName} error for ${logIdentifier}: ${error}\n${error.stack}`);
    notifyError(`${actionName} failed for project ${logIdentifier}`, error, ss, config);
    logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sfid: sfid, projectName: projectName, result: "error", errorMessage: String(error) }, config);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Synchronizes the 'Progress' value from 'Upcoming' back to 'Forecasting'.
 * Uses a script lock to prevent race conditions.
 *
 * @param {string} sfid The Salesforce ID of the project.
 * @param {string} projectName The name of the project (fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event for logging.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function syncProgressToForecasting(sfid, projectName, newValue, ss, eCtx, config) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { FORECASTING } = config.SHEETS;
  const FC = config.FORECASTING_COLS;
  const actionName = "SyncUtoF";
  const logIdentifier = sfid ? `SFID ${sfid}` : `"${projectName}"`;

  try {
    lockAcquired = acquireLockWithRetry(lock);
    if (!lockAcquired) {
      Logger.log(`${actionName}: Lock not acquired for ${logIdentifier} after multiple retries. Skipping.`);
      logAudit(ss, { action: `${actionName}-SkippedNoLock`, sourceSheet: eCtx.range.getSheet().getName(), sourceRow: eCtx.range.getRow(), sfid: sfid, projectName: projectName, result: "skipped" }, config);
      return;
    }

    const forecastingSheet = ss.getSheetByName(FORECASTING);
    if (!forecastingSheet) throw new Error(`Destination sheet "${FORECASTING}" not found`);

    const row = findRowByBestIdentifier(forecastingSheet, sfid, FC.SFID, projectName, FC.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = forecastingSheet.getRange(row, FC.PROGRESS);
      if (normalizeForComparison(targetCell.getValue()) !== normalizeForComparison(newValue)) {
        targetCell.setValue(newValue);
        updateLastEditForRow(forecastingSheet, row, config);
        logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress -> ${newValue}`, result: "updated" }, config);
      } else {
        logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid: sfid, projectName: projectName, details: "No change", result: "noop" }, config);
      }
    } else {
      logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sfid: sfid, projectName: projectName, details: "Project not found", result: "miss" }, config);
    }
  } catch (error) {
    Logger.log(`${actionName} error for ${logIdentifier}: ${error}\n${error.stack}`);
    notifyError(`${actionName} failed for project ${logIdentifier}`, error, ss, config);
    logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sfid: sfid, projectName: projectName, result: "error", errorMessage: String(error) }, config);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// =================================================================
// ==================== TRANSFER DEFINITIONS =======================
// =================================================================

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Upcoming'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function triggerUpcomingTransfer(e, sourceRowData, config) {
  const FC = config.FORECASTING_COLS;
  const UP = config.UPCOMING_COLS;

  const transferConfig = {
    transferName: "Upcoming Transfer (Permits)",
    destinationSheetName: config.SHEETS.UPCOMING,
    lastEditTrackedSheets: config.LAST_EDIT.TRACKED_SHEETS,
    destinationColumnMapping: createMapping([
      [FC.SFID, UP.SFID], [FC.PROJECT_NAME, UP.PROJECT_NAME], [FC.DEADLINE, UP.DEADLINE],
      [FC.PROGRESS, UP.PROGRESS], [FC.EQUIPMENT, UP.EQUIPMENT], [FC.PERMITS, UP.PERMITS],
      [FC.LOCATION, UP.LOCATION]
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      sfidSourceCol: FC.SFID, sfidDestCol: UP.SFID,
      projectNameSourceCol: FC.PROJECT_NAME, projectNameDestCol: UP.PROJECT_NAME
    },
    postTransferActions: {
      sort: true, sortColumn: UP.DEADLINE, sortAscending: false
    }
  };
  executeTransfer(e, transferConfig, sourceRowData);
}

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Inventory_Elevators'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function triggerInventoryTransfer(e, sourceRowData, config) {
  const FC = config.FORECASTING_COLS;
  const INV = config.INVENTORY_COLS;

  const transferConfig = {
    transferName: "Inventory Transfer (Delivered)",
    destinationSheetName: config.SHEETS.INVENTORY,
    lastEditTrackedSheets: config.LAST_EDIT.TRACKED_SHEETS,
    destinationColumnMapping: createMapping([
      [FC.PROJECT_NAME, INV.PROJECT_NAME], [FC.PROGRESS, INV.PROGRESS],
      [FC.EQUIPMENT, INV.EQUIPMENT], [FC.DETAILS, INV.DETAILS]
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      projectNameSourceCol: FC.PROJECT_NAME, projectNameDestCol: INV.PROJECT_NAME
    }
  };
  executeTransfer(e, transferConfig, sourceRowData);
}

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Framing'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @param {object} config The global configuration object.
 * @returns {void}
 */
function triggerFramingTransfer(e, sourceRowData, config) {
  const FC = config.FORECASTING_COLS;
  const FR = config.FRAMING_COLS;

  const transferConfig = {
    transferName: "Framing Transfer (In Progress)",
    destinationSheetName: config.SHEETS.FRAMING,
    lastEditTrackedSheets: config.LAST_EDIT.TRACKED_SHEETS,
    destinationColumnMapping: createMapping([
      [FC.SFID, FR.SFID], [FC.PROJECT_NAME, FR.PROJECT_NAME], [FC.EQUIPMENT, FR.EQUIPMENT],
      [FC.ARCHITECT, FR.ARCHITECT], [FC.DEADLINE, FR.DEADLINE]
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      sfidSourceCol: FC.SFID, sfidDestCol: FR.SFID,
      projectNameSourceCol: FC.PROJECT_NAME, projectNameDestCol: FR.PROJECT_NAME,
      compoundKeySourceCols: [FC.DEADLINE], compoundKeyDestCols: [FR.DEADLINE],
      keySeparator: "|"
    }
  };
  executeTransfer(e, transferConfig, sourceRowData);
}