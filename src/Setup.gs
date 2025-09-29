/**
 * @OnlyCurrentDoc
 * Setup.gs
 * Handles UI menu creation (`onOpen`) and one-time installation routines for project features.
 * This script provides the primary user interface for manual script operations.
 */

/**
 * An `onOpen` simple trigger that runs automatically when the spreadsheet is opened.
 * It creates a custom "🚀 Project Actions" menu in the Google Sheets UI, providing users
 * with easy access to the script's main functions without needing to open the script editor.
 *
 * @param {object} e The event object passed by the `onOpen` simple trigger (not used directly, but required by the signature).
 * @returns {void}
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Project Actions')
    .addItem('Update Dashboard Now', 'updateDashboard_wrapper')
    .addSeparator()
    .addItem('Run Full Setup (Install Triggers & Logging)', 'setup')
    .addItem('Initialize Last Edit Formulas (Optional)', 'initializeLastEditFormulas_wrapper')
    .addToUi();
}

/**
 * A wrapper function to call `updateDashboard` from the custom menu.
 * This is a best practice in Apps Script, as calling a function directly from the UI
 * can sometimes lead to context or permission issues. This wrapper ensures the function
 * executes in the correct global context.
 *
 * @returns {void}
 */
function updateDashboard_wrapper() {
  updateDashboard();
}

/**
 * A wrapper function to call `initializeLastEditFormulas` from the custom menu.
 * It provides a clear success message to the user via a UI alert upon completion,
 * confirming that the backfill operation has finished.
 *
 * @returns {void}
 */
function initializeLastEditFormulas_wrapper() {
  initializeLastEditFormulas();
   SpreadsheetApp.getUi().alert("Initialization complete. Formulas applied to existing rows.");
}


/**
 * The main, one-time setup routine for the entire project, executed from the custom menu.
 * This function is critical for new deployments or for repairing a broken installation. It performs:
 * 1. **Trigger Installation:** Idempotently installs the `onEdit` and `onOpen` triggers required for automations.
 * 2. **Column Creation:** Ensures "Last Edit" tracking columns are present on all configured sheets.
 * 3. **Logging Initialization:** Sets up the external logging system, which may require user authorization
 *    on the first run to create and manage a separate log spreadsheet.
 * It provides clear user feedback via UI alerts for both success and failure scenarios.
 *
 * @returns {void}
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. Install triggers
    installOnEditTrigger(ss);
    installOnOpenTriggerForLogs(ss);

    // 2. Create Last Edit columns
    ensureAllLastEditColumns(ss);
    Logger.log("Last Edit columns ensured.");

    // 3. Initialize external logging
    // This requires authorization scopes for external spreadsheets (Drive/Sheets)
    const logSS = getOrCreateLogSpreadsheet();
    ensureMonthlyLogSheet(logSS);
    Logger.log("Logging initialized.");

    ui.alert("✅ Setup Complete", "onEdit trigger installed, Last Edit columns created, and logging initialized.", ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`Setup failed: ${error}\n${error.stack}`);
    notifyError("Project Setup Routine Failed", error, ss);
    ui.alert("❌ Setup Failed", `An error occurred during setup. Please check the logs or the notification email.\nError: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Idempotently installs the required installable `onEdit` trigger for the spreadsheet.
 * "Idempotent" means that running this function multiple times will not create duplicate triggers.
 * It first scans all existing project triggers to see if one for the `onEdit` function already
 * exists. If not, it creates it. This is crucial for preventing automations from running
 * multiple times on a single edit.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet to which the trigger will be attached.
 * @returns {void}
 */
function installOnEditTrigger(ss) {
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => 
    t.getHandlerFunction() === "onEdit" && t.getEventType() === ScriptApp.EventType.ON_EDIT
  );
  
  if (!exists) {
    // Requires authorization to create triggers
    ScriptApp.newTrigger("onEdit").forSpreadsheet(ss).onEdit().create();
    Logger.log("Installable onEdit trigger created.");
  } else {
    Logger.log("Installable onEdit trigger already exists.");
  }
}

/**
 * Idempotently installs an `onOpen` trigger for the log sorting function (`sortLogSheetsOnOpen`).
 * This function ensures that the external log spreadsheet is automatically sorted every time
 * the main project spreadsheet is opened, keeping the latest logs at the top for easy viewing.
 * It checks for a pre-existing trigger to avoid creating duplicates.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet to which the `onOpen` trigger will be attached.
 * @returns {void}
 */
function installOnOpenTriggerForLogs(ss) {
  const triggers = ScriptApp.getProjectTriggers();
  const functionName = "sortLogSheetsOnOpen";

  // Check if a trigger for this function already exists
  const exists = triggers.some(t =>
    t.getHandlerFunction() === functionName && t.getEventType() === ScriptApp.EventType.ON_OPEN
  );

  if (!exists) {
    // Create the trigger for the specified spreadsheet
    ScriptApp.newTrigger(functionName).forSpreadsheet(ss).onOpen().create();
    Logger.log(`Installable onOpen trigger for "${functionName}" created.`);
  } else {
    Logger.log(`Installable onOpen trigger for "${functionName}" already exists.`);
  }
}