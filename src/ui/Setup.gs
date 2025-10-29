/**
 * @OnlyCurrentDoc
 *
 * Setup.gs
 *
 * Handles UI menu creation (`onOpen`) and one-time installation routines for project features.
 * This script provides the primary user interface for running manual script operations,
 * such as initial setup, dashboard updates, and configuration.
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

/**
 * An `onOpen` simple trigger that runs automatically when the spreadsheet is opened.
 * It creates a custom "🚀 Project Actions" menu in the Google Sheets UI and also triggers the
 * sorting of external log sheets, consolidating all `onOpen` actions into a single function.
 *
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e The event object passed by the `onOpen` simple trigger.
 * @returns {void} This function does not return a value.
 */
function onOpen(e) {
  // Run UI setup
  SpreadsheetApp.getUi()
    .createMenu('🚀 Project Actions')
    .addItem('Update Dashboard Now', 'updateDashboard_wrapper')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚙️ Setup & Configuration')
      .addItem('Run Full Setup (Install Triggers)', 'setup')
      .addItem('Set Error Notification Email', 'setErrorNotificationEmail_wrapper')
      .addSeparator()
      .addItem('Initialize Last Edit Formulas (Optional)', 'initializeLastEditFormulas_wrapper')
      .addSeparator()
      .addItem('Remove Old Triggers', 'cleanupOldTriggers')
    )
    .addSubMenu(SpreadsheetApp.getUi().createMenu("Maintenance")
      .addItem("Dedupe Upcoming (SFID → Name|Location)", "dedupeUpcomingBySfidOrNameLoc_wrapper"))
    .addToUi();

}

function dedupeUpcomingBySfidOrNameLoc_wrapper() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Dedupe Upcoming",
    "This will remove older duplicates in 'Upcoming'. Proceed?",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;
  dedupeUpcomingBySfidOrNameLoc();
  ui.alert("Done", "Deduplication completed.", ui.ButtonSet.OK);
}

/**
 * A wrapper function to call `updateDashboard` from the custom menu.
 * This is an Apps Script best practice, as calling a function directly from the UI can sometimes
 * lead to context or permission issues. This wrapper ensures the function executes in the correct global context.
 *
 * @returns {void} This function does not return a value.
 */
function updateDashboard_wrapper() {
  updateDashboard();
}

/**
 * A wrapper function to call `initializeLastEditFormulas` from the custom menu.
 * It provides a clear success message to the user via a UI alert upon completion, confirming
 * that the formula backfill operation has finished.
 *
 * @returns {void} This function does not return a value.
 */
function initializeLastEditFormulas_wrapper() {
  initializeLastEditFormulas();
   SpreadsheetApp.getUi().alert("Initialization complete. Formulas applied to existing rows.");
}


/**
 * The main, one-time setup routine for the entire project, executed from the custom menu.
 * This function is critical for new deployments or for repairing a broken installation. It performs:
 * 1. **Trigger Installation:** Idempotently installs the `onEdit` trigger required for all automations.
 * 2. **Column Creation:** Ensures "Last Edit" tracking columns are present on all configured sheets.
 * 3. **Logging Initialization:** Sets up the external logging system, creating the log sheet if needed.
 * It provides clear user feedback via UI alerts for both success and failure.
 *
 * @returns {void} This function does not return a value.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const config = CONFIG; // Top-level entry point

  try {
    // 1. Install onEdit trigger
    installOnEditTrigger(ss);

    // 2. Create Last Edit columns
    ensureAllLastEditColumns(ss, config);
    Logger.log("Last Edit columns ensured.");

    // 3. Initialize external logging
    ensureMonthlyLogSheet(ss);
    Logger.log("Logging initialized.");

    ui.alert("✅ Setup Complete", "onEdit trigger installed, Last Edit columns created, and logging initialized.", ui.ButtonSet.OK);

  } catch (error) {
    Logger.log(`Setup failed: ${error}\n${error.stack}`);
    notifyError("Project Setup Routine Failed", error, ss, config);
    ui.alert("❌ Setup Failed", `An error occurred during setup. Please check the logs or the notification email.\nError: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Idempotently installs the required installable `onEdit` trigger for the spreadsheet.
 * "Idempotent" means that running this function multiple times will not create duplicate triggers.
 * It first scans all existing project triggers to see if one for the `onEdit` function already exists.
 * If not, it creates it. This is crucial for preventing automations from running multiple times on a single edit.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet object to which the trigger will be attached.
 * @returns {void} This function does not return a value.
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
 * Prompts the user to set or update the email address for error notifications.
 * This function is called from the custom menu. It validates the user's input to ensure it is a
 * properly formatted email address before saving it to Script Properties. It provides clear feedback
 * to the user on both success and failure.
 *
 * @returns {void} This function does not return a value.
 */
function setErrorNotificationEmail_wrapper() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const config = CONFIG; // Top-level entry point
  const currentEmail = props.getProperty(config.LOGGING.ERROR_EMAIL_PROP) || "Not set";

  const response = ui.prompt(
    "Set Error Notification Email",
    `Enter the email address where error alerts should be sent.\n\nCurrently set to: ${currentEmail}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() === ui.Button.OK) {
    const newEmail = response.getResponseText().trim();
    // Simple regex for email validation
    if (/^\S+@\S+\.\S+$/.test(newEmail)) {
      props.setProperty(config.LOGGING.ERROR_EMAIL_PROP, newEmail);
      ui.alert("✅ Success", `Error notification email has been set to: ${newEmail}`, ui.ButtonSet.OK);
    } else {
      ui.alert("❌ Invalid Email", "The email address you entered is not valid. Please try again.", ui.ButtonSet.OK);
    }
  }
}

/**
 * A temporary helper function to remove the old, unused trigger
 * that is causing the "Script function not found" error.
 */
function cleanupOldTriggers() {
  const allTriggers = ScriptApp.getProjectTriggers();
  let triggerDeleted = false;

  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === "sortLogSheetsOnOpen") {
      try {
        ScriptApp.deleteTrigger(trigger);
        Logger.log("Successfully deleted the old 'sortLogSheetsOnOpen' trigger.");
        triggerDeleted = true;
      } catch (e) {
        Logger.log(`Failed to delete trigger: ${e.message}`);
      }
    }
  }

  const ui = SpreadsheetApp.getUi();
  if (triggerDeleted) {
    ui.alert("Success!", "The old, unused 'sortLogSheetsOnOpen' trigger has been removed. The error message you saw should now be gone.", ui.ButtonSet.OK);
  } else {
    ui.alert("Info", "No trigger for 'sortLogSheetsOnOpen' was found. It may have already been removed.", ui.ButtonSet.OK);
  }
}