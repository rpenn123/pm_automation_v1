/**
 * @OnlyCurrentDoc
 *
 * Automations.gs
 *
 * This script contains the primary `onEdit` trigger, which serves as the central hub for all sheet automations.
 * It orchestrates data synchronization between sheets and triggers transfers based on specific cell edits.
 * The logic is rule-based, making it extensible for future automations.
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

/**
 * The main `onEdit` function, serving as the central hub for all sheet automations.
 * This function is installed as an installable trigger by `Setup.gs` and responds to any user edit.
 *
 * **Execution Flow:**
 * 1.  **Guard Clauses:** Performs several checks to exit early for irrelevant edits (e.g., multi-cell edits, header edits, or non-value changes).
 * 2.  **Batch Row Read:** Reads the entire edited row in a single operation to optimize performance.
 * 3.  **Last Edit Tracking:** Calls `updateLastEditForRow` to timestamp the modification for data lifecycle management.
 * 4.  **Audit Logging:** Records the basic edit details to the audit log for accountability.
 * 5.  **Rule-Based Routing:** Iterates through a `rules` array to find a handler that matches the edited sheet and column.
 * 6.  **Handler Execution:** Executes the first matching handler, passing the event object and pre-read row data to it.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The event object passed by the `onEdit` trigger, containing details about the cell edit.
 * @returns {void} This function does not return a value.
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
  const correlationId = Utilities.getUuid(); // For tracing
  const config = CONFIG;

  try {
    // 1. Batch-read the entire edited row once for performance, with retry.
    const sourceRowData = withRetry(
      () => sheet.getRange(editedRow, 1, 1, sheet.getMaxColumns()).getValues()[0],
      { functionName: "onEdit:readSourceRow", correlationId: correlationId }
    );

    // 2. Always update Last Edit tracking if applicable.
    if (config.LAST_EDIT.TRACKED_SHEETS.includes(sheetName)) {
      updateLastEditForRow(sheet, editedRow, config);
      const a1Notation = e.range.getA1Notation ? e.range.getA1Notation() : 'unknown';
      logAudit(e.source, {
        correlationId: correlationId,
        action: 'Row Edit',
        sourceSheet: sheetName,
        sourceRow: editedRow,
        details: `Cell ${a1Notation} updated. New value: "${e.value}"`,
        result: 'success'
      }, config);
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

    for (const rule of rules) {
      if (sheetName === rule.sheet && editedCol === rule.col) {
        if (!rule.valueCheck || rule.valueCheck(e.value)) {
          // Pass all necessary context to the handler.
          rule.handler(e, sourceRowData, config, correlationId);
          return; // Stop after the first matching rule.
        }
      }
    }
  } catch (error) {
    handleError(error, {
      correlationId: correlationId,
      functionName: "onEdit",
      spreadsheet: e.source,
      extra: { sheetName: sheetName, editedCol: editedCol, editedRow: editedRow }
    }, config);
  }
}

// =================================================================
// ==================== SYNC HANDLERS ==============================
// =================================================================

/**
 * Handles an edit to the 'Progress' column in the 'Forecasting' sheet. It performs two actions:
 * 1.  Always syncs the new 'Progress' value to the corresponding row in the 'Upcoming' sheet.
 * 2.  Conditionally triggers a transfer to the 'Framing' sheet if the new value is "In Progress".
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function handleSyncAndPotentialFramingTransfer(e, sourceRowData, config, correlationId) {
  const FC = config.FORECASTING_COLS;
  const sfid = sourceRowData[FC.SFID - 1];
  const projectName = sourceRowData[FC.PROJECT_NAME - 1];
  const newValue = e.value;

  // 1. Synchronization.
  if (sfid || projectName) {
    syncProgressToUpcoming(sfid, projectName, newValue, e.source, e, config, correlationId);
  }

  // 2. Conditional actions based on the new status.
  const normalizedValue = normalizeString(newValue);
  const STATUS = config.STATUS_STRINGS;

  if (normalizedValue === STATUS.IN_PROGRESS.toLowerCase()) {
    triggerFramingTransfer(e, sourceRowData, config, correlationId);
  } else if (normalizedValue === STATUS.INSPECTIONS.toLowerCase()) {
    triggerInspectionEmail(e, sourceRowData, config, correlationId);
  }
}

/**
 * Handles an edit to the 'Progress' column in the 'Upcoming' sheet, syncing the new value
 * back to the corresponding row in the 'Forecasting' sheet.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Upcoming' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function triggerSyncToForecasting(e, sourceRowData, config, correlationId) {
  const UP = config.UPCOMING_COLS;
  const sfid = sourceRowData[UP.SFID - 1];
  const projectName = sourceRowData[UP.PROJECT_NAME - 1];

  if (sfid || projectName) {
    syncProgressToForecasting(sfid, projectName, e.value, e.source, e, config, correlationId);
  }
}

/**
 * Synchronizes the 'Progress' value from the 'Forecasting' sheet to the 'Upcoming' sheet.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project, used as a fallback if SFID is not available.
 * @param {any} newValue The new value of the 'Progress' cell to be set in the 'Upcoming' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original `onEdit` event context, used for logging purposes.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function syncProgressToUpcoming(sfid, projectName, newValue, ss, eCtx, config, correlationId) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { UPCOMING } = config.SHEETS;
  const UP = config.UPCOMING_COLS;
  const actionName = "SyncFtoU";

  try {
    lockAcquired = lock.tryLock(1000); // Try to acquire lock for 1 second

    if (!lockAcquired) {
      // If lock is busy, log it and exit gracefully.
      logAudit(ss, {
        correlationId: correlationId,
        action: actionName,
        details: "Sync skipped: lock was busy.",
        result: "noop"
      }, config);
      return;
    }

    const upcomingSheet = ss.getSheetByName(UPCOMING);
    if (!upcomingSheet) throw new ConfigurationError(`Destination sheet "${UPCOMING}" not found`);

    const row = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = upcomingSheet.getRange(row, UP.PROGRESS);
      if (normalizeForComparison(targetCell.getValue()) !== normalizeForComparison(newValue)) {
        withRetry(() => targetCell.setValue(newValue), { functionName: `${actionName}:setValue`, correlationId: correlationId });
        updateLastEditForRow(upcomingSheet, row, config);
        logAudit(ss, { correlationId, action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid, projectName, details: `Progress -> ${newValue}`, result: "updated" }, config);
      } else {
        logAudit(ss, { correlationId, action: actionName, sourceSheet: UPCOMING, sourceRow: row, sfid, projectName, details: "No change", result: "noop" }, config);
      }
    } else {
      logAudit(ss, { correlationId, action: actionName, sourceSheet: UPCOMING, sfid, projectName, details: "Project not found", result: "miss" }, config);
    }
  } catch (error) {
    handleError(error, {
      correlationId: correlationId,
      functionName: actionName,
      spreadsheet: ss,
      extra: { sfid, projectName }
    }, config);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

/**
 * Synchronizes the 'Progress' value from the 'Upcoming' sheet back to the 'Forecasting' sheet.
 *
 * @param {string} sfid The Salesforce ID of the project to sync.
 * @param {string} projectName The name of the project, used as a fallback if SFID is not available.
 * @param {any} newValue The new value of the 'Progress' cell to be set in the 'Forecasting' sheet.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet object.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} eCtx The original `onEdit` event context, used for logging purposes.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function syncProgressToForecasting(sfid, projectName, newValue, ss, eCtx, config, correlationId) {
  const lock = LockService.getScriptLock();
  let lockAcquired = false;
  const { FORECASTING } = config.SHEETS;
  const FC = config.FORECASTING_COLS;
  const actionName = "SyncUtoF";

  try {
    lockAcquired = lock.tryLock(1000); // Try to acquire lock for 1 second

    if (!lockAcquired) {
      // If lock is busy, log it and exit gracefully.
      logAudit(ss, {
        correlationId: correlationId,
        action: actionName,
        details: "Sync skipped: lock was busy.",
        result: "noop"
      }, config);
      return;
    }

    const forecastingSheet = ss.getSheetByName(FORECASTING);
    if (!forecastingSheet) throw new ConfigurationError(`Destination sheet "${FORECASTING}" not found`);

    const row = findRowByBestIdentifier(forecastingSheet, sfid, FC.SFID, projectName, FC.PROJECT_NAME);

    if (row !== -1) {
      const targetCell = forecastingSheet.getRange(row, FC.PROGRESS);
      if (normalizeForComparison(targetCell.getValue()) !== normalizeForComparison(newValue)) {
        withRetry(() => targetCell.setValue(newValue), { functionName: `${actionName}:setValue`, correlationId: correlationId });
        updateLastEditForRow(forecastingSheet, row, config);
        logAudit(ss, { correlationId, action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid, projectName, details: `Progress -> ${newValue}`, result: "updated" }, config);
      } else {
        logAudit(ss, { correlationId, action: actionName, sourceSheet: FORECASTING, sourceRow: row, sfid, projectName, details: "No change", result: "noop" }, config);
      }
    } else {
      logAudit(ss, { correlationId, action: actionName, sourceSheet: FORECASTING, sfid, projectName, details: "Project not found", result: "miss" }, config);
    }
  } catch (error) {
    handleError(error, {
      correlationId: correlationId,
      functionName: actionName,
      spreadsheet: ss,
      extra: { sfid, projectName }
    }, config);
  } finally {
    if (lockAcquired) lock.releaseLock();
  }
}

// =================================================================
// ==================== TRANSFER DEFINITIONS =======================
// =================================================================

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Upcoming'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function triggerUpcomingTransfer(e, sourceRowData, config, correlationId) {
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
  executeTransfer(e, transferConfig, sourceRowData, correlationId);
}

/**
 * Defines and triggers an email notification when a project's progress is set to 'Inspections'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function triggerInspectionEmail(e, sourceRowData, config, correlationId) {
  const FC = config.FORECASTING_COLS;
  const UP = config.UPCOMING_COLS;
  const ss = e.source;
  const actionName = "InspectionEmail";

  try {
    const projectName = sourceRowData[FC.PROJECT_NAME - 1];
    const equipment = sourceRowData[FC.EQUIPMENT - 1];
    const location = sourceRowData[FC.LOCATION - 1];
    const sfid = sourceRowData[FC.SFID - 1];

    // Construction status is on the 'Upcoming' sheet, so we need to look it up.
    const upcomingSheet = ss.getSheetByName(config.SHEETS.UPCOMING);
    let constructionStatus = "N/A"; // Default value

    if (upcomingSheet) {
      const upcomingRow = findRowByBestIdentifier(upcomingSheet, sfid, UP.SFID, projectName, UP.PROJECT_NAME);
      if (upcomingRow !== -1) {
        // Correctly get the value from the identified row and column
        constructionStatus = upcomingSheet.getRange(upcomingRow, UP.CONSTRUCTION).getValue();
      } else {
        logAudit(ss, {
          correlationId,
          action: actionName,
          details: `Could not find matching project in '${config.SHEETS.UPCOMING}' to look up Construction status.`,
          result: "warning"
        }, config);
      }
    } else {
        logAudit(ss, {
          correlationId,
          action: actionName,
          details: `Sheet '${config.SHEETS.UPCOMING}' not found for Construction status lookup.`,
          result: "warning"
        }, config);
    }

    const subject = `Re: Inspection Update | ${projectName}`;
    const body = `
Project: ${projectName}
Status (Progress): Ready for Inspections
Equipment: ${equipment}
Construction: ${constructionStatus}
Address: ${location}
    `.trim();

    MailApp.sendEmail({
      to: "pm@mobility123.com",
      subject: subject,
      body: body,
    });

    logAudit(ss, {
      correlationId: correlationId,
      action: actionName,
      sourceSheet: e.range.getSheet().getName(),
      sourceRow: e.range.getRow(),
      sfid: sfid,
      projectName: projectName,
      details: "Inspection email sent successfully.",
      result: "success"
    }, config);

  } catch (error) {
    handleError(error, {
      correlationId: correlationId,
      functionName: actionName,
      spreadsheet: ss,
      extra: {
        sourceRow: e.range.getRow()
      }
    }, config);
  }
}

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Inventory_Elevators'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function triggerInventoryTransfer(e, sourceRowData, config, correlationId) {
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
  executeTransfer(e, transferConfig, sourceRowData, correlationId);
}

/**
 * Defines and triggers the transfer of a project row from 'Forecasting' to 'Framing'.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e The `onEdit` event object from the trigger.
 * @param {any[]} sourceRowData The pre-read data from the entire edited row in the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId A unique ID for tracing the entire operation.
 */
function triggerFramingTransfer(e, sourceRowData, config, correlationId) {
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
      projectNameSourceCol: FC.PROJECT_NAME, projectNameDestCol: FR.PROJECT_NAME
    },
    syncOnDuplicate: true
  };
  executeTransfer(e, transferConfig, sourceRowData, correlationId);
}