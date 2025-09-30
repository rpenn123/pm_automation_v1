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

  // 1. Batch-read the entire edited row once for performance.
  const sourceRowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];

  // 2. Always update Last Edit tracking if applicable. This runs independently of other rules.
  if (CONFIG.LAST_EDIT.TRACKED_SHEETS.includes(sheetName)) {
    updateLastEditForRow(sheet, editedRow);
  }

  // 3. Route the edit event to the appropriate handler based on a set of rules.
  const { FORECASTING, UPCOMING } = CONFIG.SHEETS;
  const FC = CONFIG.FORECASTING_COLS;
  const UP = CONFIG.UPCOMING_COLS;
  const STATUS = CONFIG.STATUS_STRINGS;

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
          // Pass the pre-read row data to the handler.
          rule.handler(e, sourceRowData);
        } catch (error) {
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
 * @returns {void}
 */
function handleSyncAndPotentialFramingTransfer(e, sourceRowData) {
  const FC = CONFIG.FORECASTING_COLS;

  // Use the pre-read data instead of making new I/O calls.
  const sfid = sourceRowData[FC.SFID - 1];
  const projectName = sourceRowData[FC.PROJECT_NAME - 1];
  const newValue = e.value;

  // 1. Synchronization.
  if (sfid || projectName) {
    syncProgressToUpcoming(sfid, projectName, newValue, e.source, e);
  }

  // 2. Conditional Transfer to Framing.
  if (normalizeString(newValue) === CONFIG.STATUS_STRINGS.IN_PROGRESS.toLowerCase()) {
    triggerFramingTransfer(e, sourceRowData);
  }
}

/**
 * Handles an edit to 'Progress' in 'Upcoming', syncing the value back to 'Forecasting'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @returns {void}
 */
function triggerSyncToForecasting(e, sourceRowData) {
  const UP = CONFIG.UPCOMING_COLS;

  // Use the pre-read data.
  const sfid = sourceRowData[UP.SFID - 1];
  const projectName = sourceRowData[UP.PROJECT_NAME - 1];

  if (sfid || projectName) {
    syncProgressToForecasting(sfid, projectName, e.value, e.source, e);
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
 * @returns {void}
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

    const row = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = upcomingSheet.getRange(row, UP.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      // Only write if the value is different to avoid re-triggering onEdit.
      if (currentValueStr !== newValueStr) {
        targetCell.setValue(newValue);
        updateLastEditForRow(upcomingSheet, row);
        logAudit(ss, { action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress -> ${newValue}`, result: "updated" });
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
 * Synchronizes the 'Progress' value from 'Upcoming' back to 'Forecasting'.
 * Uses a script lock to prevent race conditions.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project to sync (used as a fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event object for detailed logging context.
 * @returns {void}
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

    const row = findRowByBestIdentifier(forecastingSheet, sfid, FC.SFID, projectName, FC.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = forecastingSheet.getRange(row, FC.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      // Only write if the value is different.
      if (currentValueStr !== newValueStr) {
        targetCell.setValue(newValue);
        updateLastEditForRow(forecastingSheet, row);
        logAudit(ss, { action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid: sfid, projectName: projectName, details: `Progress -> ${newValue}`, result: "updated" });
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

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Upcoming'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @returns {void}
 */
function triggerUpcomingTransfer(e, sourceRowData) {
  const FC = CONFIG.FORECASTING_COLS;
  const UP = CONFIG.UPCOMING_COLS;

  const config = {
    transferName: "Upcoming Transfer (Permits)",
    destinationSheetName: CONFIG.SHEETS.UPCOMING,
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
  executeTransfer(e, config, sourceRowData);
}

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Inventory_Elevators'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @returns {void}
 */
function triggerInventoryTransfer(e, sourceRowData) {
  const FC = CONFIG.FORECASTING_COLS;
  const INV = CONFIG.INVENTORY_COLS;

  const config = {
    transferName: "Inventory Transfer (Delivered)",
    destinationSheetName: CONFIG.SHEETS.INVENTORY,
    destinationColumnMapping: createMapping([
      [FC.PROJECT_NAME, INV.PROJECT_NAME], [FC.PROGRESS, INV.PROGRESS],
      [FC.EQUIPMENT, INV.EQUIPMENT], [FC.DETAILS, INV.DETAILS]
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      projectNameSourceCol: FC.PROJECT_NAME, projectNameDestCol: INV.PROJECT_NAME
    }
  };
  executeTransfer(e, config, sourceRowData);
}

/**
 * Defines and triggers the transfer from 'Forecasting' to 'Framing'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @param {any[]} sourceRowData The pre-read data from the edited row.
 * @returns {void}
 */
function triggerFramingTransfer(e, sourceRowData) {
  const FC = CONFIG.FORECASTING_COLS;
  const FR = CONFIG.FRAMING_COLS;

  const config = {
    transferName: "Framing Transfer (In Progress)",
    destinationSheetName: CONFIG.SHEETS.FRAMING,
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
  executeTransfer(e, config, sourceRowData);
}