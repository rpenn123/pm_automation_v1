/**
 * @OnlyCurrentDoc
 * Setup.gs
 * Handles UI menu creation (onOpen) and one-time installation routines.
 */

/**
 * A simple trigger that runs when the spreadsheet is opened. It creates a custom menu
 * in the UI for users to interact with the script's main functions.
 *
 * @param {object} e The event object passed by the `onOpen` simple trigger.
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
 * This is a best practice in Apps Script to ensure the global scope is correctly
 * handled when a function is called from a UI element.
 * @returns {void}
 */
function updateDashboard_wrapper() {
  updateDashboard();
}

/**
 * A wrapper function to call `initializeLastEditFormulas` from the custom menu.
 * It also displays an alert to the user upon completion.
 * @returns {void}
 */
function initializeLastEditFormulas_wrapper() {
  initializeLastEditFormulas();
   SpreadsheetApp.getUi().alert("Initialization complete. Formulas applied to existing rows.");
}


/**
 * The main, one-time setup routine for the entire project. This function is called from the custom menu
 * and performs the following critical setup steps:
 * 1. Installs the installable `onEdit` trigger required for all automations.
 * 2. Ensures the "Last Edit" tracking columns are present on all configured sheets.
 * 3. Initializes the external logging system by creating the log spreadsheet and the current month's log sheet.
 * It provides user feedback via UI alerts for both success and failure.
 * @returns {void}
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    // 1. Installable onEdit trigger
    installOnEditTrigger(ss);

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
 * It first checks if a trigger for the `onEdit` function already exists. If not, it creates one.
 * This prevents the creation of duplicate triggers.
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