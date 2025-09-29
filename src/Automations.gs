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
 * 1.  **Guard Clauses:** Performs several checks to exit early for irrelevant edits, such as multi-cell edits,
 *     header row edits, or edits that don't actually change the cell's value. This is a critical
 *     performance optimization.
 * 2.  **Last Edit Tracking:** It first calls the `LastEditService` to update the "Last Edit" timestamp
 *     for the modified row, if the sheet is configured for tracking.
 * 3.  **Rule-Based Routing:** It then uses a `rules` array to find the appropriate handler function
 *     based on which sheet and column were edited. This makes the logic clean and extensible.
 * 4.  **Handler Execution:** The first matching rule's handler is executed within a try-catch block
 *     to ensure that an error in one automation doesn't halt others.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object passed by the onEdit trigger, containing details about the cell edit.
 * @returns {void}
 */
function onEdit(e) {
  if (!e || !e.range) return;

  // Performance optimizations: exit early for multi-cell edits, header edits, or non-value changes.
  if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
  if (e.range.getRow() <= 1) return;
  // Crucial check to prevent infinite loops from programmatic edits that re-trigger onEdit.
  if (normalizeForComparison(e.value) === normalizeForComparison(e.oldValue)) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const editedCol = e.range.getColumn();
  const editedRow = e.range.getRow();

  // 1. Always update Last Edit tracking if applicable. This runs independently of other rules.
  if (CONFIG.LAST_EDIT.TRACKED_SHEETS.includes(sheetName)) {
    // Errors during Last Edit update are handled within the service function itself.
    updateLastEditForRow(sheet, editedRow);
  }

  // 2. Route the edit event to the appropriate handler based on a set of rules.
  const { FORECASTING, UPCOMING } = CONFIG.SHEETS;
  const FC = CONFIG.FORECASTING_COLS;
  const UP = CONFIG.UPCOMING_COLS;
  const STATUS = CONFIG.STATUS_STRINGS;

  // Define routing rules. Order can matter if multiple actions could be triggered by the same edit.
  const rules = [
    // --- Sync Rules ---
    {
      // Forecasting Progress Edit: Syncs to Upcoming AND can trigger a transfer to Framing.
      sheet: FORECASTING,
      col: FC.PROGRESS,
      handler: handleSyncAndPotentialFramingTransfer
    },
    {
      // Upcoming Progress Edit: Syncs the change back to the Forecasting sheet.
      sheet: UPCOMING,
      col: UP.PROGRESS,
      handler: triggerSyncToForecasting
    },

    // --- Transfer Rules (triggered by specific cell values) ---
    {
      // Forecasting Permits "approved": Transfers the project to the Upcoming sheet.
      sheet: FORECASTING,
      col: FC.PERMITS,
      valueCheck: (val) => normalizeString(val) === STATUS.PERMIT_APPROVED.toLowerCase(),
      handler: triggerUpcomingTransfer
    },
    {
      // Forecasting Delivered set to TRUE: Transfers the project to the Inventory sheet.
      sheet: FORECASTING,
      col: FC.DELIVERED,
      valueCheck: (val) => isTrueLike(val),
      handler: triggerInventoryTransfer
    },
  ];

  // Execute the first matching rule.
  for (const rule of rules) {
    if (sheetName === rule.sheet && editedCol === rule.col) {
      // If the rule has a valueCheck, ensure it passes before executing the handler.
      if (!rule.valueCheck || rule.valueCheck(e.value)) {
        try {
          rule.handler(e);
        } catch (error) {
          // Centralized error handling for any failure within a handler.
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
        // Stop after the first successful handler execution to prevent multiple actions on one edit.
        return;
      }
    }
  }
}

// =================================================================
// ==================== SYNC HANDLERS ==============================
// =================================================================

/**
 * Handles an edit to the 'Progress' column in 'Forecasting'. This is a composite handler that
 * orchestrates two potential actions from a single event:
 * 1.  **Always Sync:** It synchronizes the new 'Progress' value to the corresponding project
 *     in the 'Upcoming' sheet to keep them aligned.
 * 2.  **Conditional Transfer:** If the new value is specifically "In Progress", it also triggers
 *     a data transfer to the 'Framing' sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function handleSyncAndPotentialFramingTransfer(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const FC = CONFIG.FORECASTING_COLS;

  // Read both SFID and Project Name from the source row for robust matching.
  const sfid = sheet.getRange(editedRow, FC.SFID).getValue();
  const projectName = sheet.getrange(editedRow, FC.PROJECT_NAME).getValue();
  const newValue = e.value;

  // 1. Synchronization (ensure at least one identifier exists).
  if (sfid || projectName) {
    syncProgressToUpcoming(sfid, projectName, newValue, e.source, e);
  }

  // 2. Conditional Transfer to Framing.
  if (normalizeString(newValue) === CONFIG.STATUS_STRINGS.IN_PROGRESS.toLowerCase()) {
    triggerFramingTransfer(e);
  }
}

/**
 * Handles an edit to the 'Progress' column in the 'Upcoming' sheet.
 * This handler completes the two-way data sync by propagating the 'Progress'
 * value back to the corresponding project in the 'Forecasting' sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The onEdit event object.
 * @returns {void}
 */
function triggerSyncToForecasting(e) {
  const sheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const UP = CONFIG.UPCOMING_COLS;

  // Read both SFID and Project Name for robust matching.
  const sfid = sheet.getRange(editedRow, UP.SFID).getValue();
  const projectName = sheet.getRange(editedRow, UP.PROJECT_NAME).getValue();

  if (sfid || projectName) {
    syncProgressToForecasting(sfid, projectName, e.value, e.source, e);
  }
}

/**
 * Synchronizes the 'Progress' value from 'Forecasting' to 'Upcoming'.
 * This function uses a script lock to prevent race conditions and infinite loops that could
 * occur if both sync functions were to run simultaneously. It finds the target row using an
 * SFID-first strategy and only writes the new value if it has actually changed, preventing
 * unnecessary `onEdit` trigger events.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project to sync (used as a fallback).
 * @param {*} newValue The new value of the 'Progress' cell.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The parent spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original onEdit event object, passed for detailed logging context.
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

    // Use the robust SFID-first lookup function.
    const row = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = upcomingSheet.getRange(row, UP.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      // Only write if the value is different to avoid re-triggering onEdit.
      if (currentValueStr !== newValueStr) {
        targetCell.setValue(newValue);
        updateLastEditForRow(upcomingSheet, row);
        SpreadsheetApp.flush(); // Ensure the write is committed before the script ends.
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
 * Synchronizes the 'Progress' value from 'Upcoming' back to 'Forecasting'.
 * This is the counterpart to `syncProgressToUpcoming` and completes the two-way sync.
 * It uses the same locking and value-checking mechanisms to ensure data integrity and
 * prevent infinite loops.
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

    // Use the robust SFID-first lookup function.
    const row = findRowByBestIdentifier(forecastingSheet, sfid, FC.SFID, projectName, FC.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = forecastingSheet.getRange(row, FC.PROGRESS);
      const currentValueStr = normalizeForComparison(targetCell.getValue());
      const newValueStr = normalizeForComparison(newValue);

      // Only write if the value is different.
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
// These functions act as configuration providers for the generic TransferEngine.

/**
 * Defines and triggers the transfer of a project from 'Forecasting' to 'Upcoming'.
 * This function acts as a configuration provider for the `TransferEngine`. It is initiated when
 * the 'Permits' column in 'Forecasting' is set to "approved". It defines the source/destination
 * mapping and specifies that the destination sheet should be sorted by deadline after the transfer.
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
    sourceColumnsNeeded: [
      FC.SFID, FC.PROJECT_NAME, FC.EQUIPMENT, FC.PROGRESS,
      FC.PERMITS, FC.DEADLINE, FC.LOCATION
    ],
    destinationColumnMapping: createMapping([
      [FC.SFID,         UP.SFID],
      [FC.PROJECT_NAME, UP.PROJECT_NAME],
      [FC.DEADLINE,     UP.DEADLINE],
      [FC.PROGRESS,     UP.PROGRESS],
      [FC.EQUIPMENT,    UP.EQUIPMENT],
      [FC.PERMITS,      UP.PERMITS],
      [FC.LOCATION,     UP.LOCATION]
      // Note: Some columns are intentionally left blank in the destination.
    ]),
    duplicateCheckConfig: {
      checkEnabled: true,
      sfidSourceCol: FC.SFID,
      sfidDestCol: UP.SFID,
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: UP.PROJECT_NAME
    },
    postTransferActions: {
      sort: true,
      sortColumn: UP.DEADLINE,
      sortAscending: false // Sort by deadline descending to show nearest deadlines first.
    }
  };
  executeTransfer(e, config);
}

/**
 * Defines and triggers the transfer of a project from 'Forecasting' to 'Inventory_Elevators'.
 * This function acts as a configuration provider for the `TransferEngine`. It is initiated when
 * the 'Delivered' column in 'Forecasting' is set to `TRUE`. It defines the specific columns
 * needed for an inventory record.
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
      // This transfer primarily relies on Project Name for duplicate checks as SFID may not be relevant.
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: INV.PROJECT_NAME
    }
  };
  executeTransfer(e, config);
}

/**
 * Defines and triggers the transfer of a project from 'Forecasting' to 'Framing'.
 * This function acts as a configuration provider for the `TransferEngine`. It's initiated when
 * 'Progress' is set to "In Progress". It uses a compound key (Project Name + Deadline) for
 * duplicate checking. This is a key design choice: it allows the *same project* to be added
 * multiple times if its deadline changes, as each instance represents a distinct framing task.
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
    // Use SFID as primary duplicate check, falling back to a compound key for legacy data.
    duplicateCheckConfig: {
      checkEnabled: true,
      sfidSourceCol: FC.SFID,
      sfidDestCol: FR.SFID,
      projectNameSourceCol: FC.PROJECT_NAME,
      projectNameDestCol: FR.PROJECT_NAME,
      compoundKeySourceCols: [FC.DEADLINE], // The deadline is part of the unique key.
      compoundKeyDestCols: [FR.DEADLINE],
      keySeparator: "|"
    }
  };
  executeTransfer(e, config);
}