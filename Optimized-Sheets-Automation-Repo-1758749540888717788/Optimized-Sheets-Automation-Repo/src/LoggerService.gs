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
    const subject = `[Error] ${appName}: ${subjectDetails}`;
    const timestamp = new Date();

    // Determine the spreadsheet context
    const activeSS = ss || SpreadsheetApp.getActiveSpreadsheet();
    const ssName = (activeSS && activeSS.getName) ? activeSS.getName() : "Unknown";
    const ssId = (activeSS && activeSS.getId) ? activeSS.getId() : "N/A";
    const ssUrl = (activeSS && activeSS.getUrl) ? activeSS.getUrl() : "#";

    // Format error details for HTML and plain text
    const errorMessage = (error && error.message) ? error.message : String(error);
    const stackTrace = (error && error.stack) ? error.stack : "No stack trace available.";

    // --- Create HTML Body for better readability ---
    let htmlBody = `
      <p>A critical error occurred in the <strong>${appName}</strong> script.</p>
      <hr>
      <p><strong>Timestamp:</strong> ${timestamp.toUTCString()}</p>
      <p><strong>Spreadsheet:</strong> <a href="${ssUrl}">${ssName}</a> (ID: ${ssId})</p>
      <p><strong>Error Message:</strong></p>
      <p style="color: red; font-family: monospace; background-color: #f5f5f5; padding: 10px; border-radius: 4px;">${errorMessage}</p>
      <p><strong>Stack Trace:</strong></p>
      <pre style="font-family: monospace; background-color: #f5f5f5; padding: 10px; border-radius: 4px;">${stackTrace}</pre>
    `;

    // --- Create Plain Text Fallback Body ---
    let body = `A critical error occurred in the ${appName} script.\n\n` +
               `Timestamp: ${timestamp.toISOString()}\n` +
               `Spreadsheet: ${ssName} (${ssId})\n` +
               `URL: ${ssUrl}\n\n` +
               `Error Message: ${errorMessage}\n\n` +
               `Stack Trace:\n${stackTrace}\n`;

    // Send email with HTML body and plain text fallback
    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
      noReply: true
    });
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