/**
 * @OnlyCurrentDoc
 * LoggerService.gs
 * Handles error notifications via email and the external monthly audit log system.
 */

/**
 * Email notification on critical errors.
 * Centralized function used across the entire application.
 */
function notifyError(subjectDetails, error, ss) {
  const email = CONFIG.ERROR_NOTIFICATION_EMAIL;
  const appName = CONFIG.APP_NAME;

  // Validate email format
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) {
    Logger.log(`Error email skipped due to invalid email address: "${email}"`);
    return;
  }

  try {
    const subject = `Critical Error in ${appName}: ${subjectDetails}`;
    let body = `A critical error occurred in the ${appName} script.\n\n`;
    body += `Timestamp: ${new Date().toISOString()}\n`;

    // Determine the spreadsheet context
    const activeSS = ss || SpreadsheetApp.getActiveSpreadsheet();
    if (activeSS && activeSS.getId) {
      body += `Spreadsheet: ${activeSS.getName()} (${activeSS.getId()})\n\n`;
    } else {
      body += `Spreadsheet: Unknown (could not determine active spreadsheet)\n\n`;
    }

    // Format error details
    body += `Error Message: ${(error && error.message) ? error.message : String(error)}\n\n`;
    if (error && error.stack) {
      body += `Stack Trace:\n${error.stack}\n`;
    }

    // Send email (requires authorization scope)
    MailApp.sendEmail({ to: email, subject: subject, body: body, noReply: true });
    Logger.log(`Sent error email to ${email}`);
  } catch (mailError) {
    Logger.log(`CRITICAL: Failed to send error email: ${mailError}`);
  }
}

/**
 * Find or create the external log spreadsheet and store its ID in script properties.
 * If creation/open fails (e.g., due to missing scopes), falls back to the active spreadsheet.
 */
function getOrCreateLogSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(CONFIG.LOGGING.SPREADSHEET_ID_PROP);

  // 1. Try opening the stored ID.
  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (e) {
      Logger.log(`Stored log spreadsheet ID invalid or inaccessible: ${e}`);
      // fall through
    }
  }

  // 2. Try creating a new external log spreadsheet (requires Drive/Sheets scopes).
  try {
    const newSS = SpreadsheetApp.create(CONFIG.LOGGING.SPREADSHEET_NAME);
    props.setProperty(CONFIG.LOGGING.SPREADSHEET_ID_PROP, newSS.getId());
    return newSS;
  } catch (e2) {
    // 3. Fallback: use the active workbook.
    const active = SpreadsheetApp.getActiveSpreadsheet();
    try {
      // Notify about the fallback, but don't let notification failure stop the process.
      // Use a property to prevent spamming notifications on every log attempt if fallback is active.
      if (!props.getProperty("FALLBACK_ACTIVE_NOTIFIED")) {
         notifyError("Could not create external log workbook. Falling back to internal logging.", e2, active);
         props.setProperty("FALLBACK_ACTIVE_NOTIFIED", "true");
      }
    } catch (ignore) {}
    Logger.log("FALLBACK: Using active spreadsheet for monthly logs.");
    return active;
  }
}

/** Ensure a monthly tab exists in the provided log workbook. */
function ensureMonthlyLogSheet(logSS, monthKey) {
  // Use the padded month key for standardized log sheet names
  const key = monthKey || getMonthKeyPadded();
  let sh = logSS.getSheetByName(key);
  if (!sh) {
    sh = logSS.insertSheet(key);
    // Set headers and freeze the first row
    sh.getRange(1, 1, 1, 11).setValues([[
      "Timestamp", "User", "Action",
      "SourceSpreadsheetName", "SourceSpreadsheetId",
      "SourceSheet", "SourceRow", "ProjectName",
      "Details", "Result", "ErrorMessage"
    ]]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Write an audit entry to the external monthly log (or fallback). */
function logAudit(sourceSS, entry) {
  try {
    const logSS = getOrCreateLogSpreadsheet();
    const sheet = ensureMonthlyLogSheet(logSS);
    // Safely get the active user's email (requires authorization scope)
    const user = Session.getActiveUser() ? Session.getActiveUser().getEmail() : "unknown";

    const newRow = [
      new Date(),
      user,
      entry.action || "",
      (sourceSS && sourceSS.getName) ? sourceSS.getName() : "",
      (sourceSS && sourceSS.getId) ? sourceSS.getId() : "",
      entry.sourceSheet || "",
      entry.sourceRow || "",
      entry.projectName || "",
      entry.details || "",
      entry.result || "",
      entry.errorMessage || ""
    ];

    // Append row efficiently
    sheet.appendRow(newRow);

  } catch (e) {
    Logger.log(`CRITICAL: Audit logging failure: ${e}`);
    // If the logging system itself fails, attempt to notify.
    notifyError("Audit logging system failed critically", e, sourceSS);
  }
}