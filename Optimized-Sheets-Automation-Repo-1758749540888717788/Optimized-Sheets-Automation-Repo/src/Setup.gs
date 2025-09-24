/**
 * @OnlyCurrentDoc
 * Setup.gs
 * Handles UI menu creation (onOpen) and one-time installation routines.
 */

/**
 * Creates the custom menu when the spreadsheet opens.
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

// Wrappers are used to ensure the global scope is correctly initialized when called from the menu.
function updateDashboard_wrapper() {
  updateDashboard();
}

function initializeLastEditFormulas_wrapper() {
  initializeLastEditFormulas();
   SpreadsheetApp.getUi().alert("Initialization complete. Formulas applied to existing rows.");
}


/**
 * One-time setup routine.
 * - Ensures installable onEdit trigger exists.
 * - Ensures Last Edit columns exist on key sheets.
 * - Ensures external log workbook and current month tab (with fallback).
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

/** Installs the required onEdit trigger if it doesn't exist. */
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