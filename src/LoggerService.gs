/**
 * @OnlyCurrentDoc
 * LoggerService.gs
 * Handles error notifications via email and the external monthly audit log system.
 */

/**
 * Sends a formatted email notification when a critical error occurs.
 * This function constructs a detailed HTML email with the error message, stack trace,
 * and spreadsheet context. It includes a plain text fallback for email clients
 * that do not support HTML.
 *
 * @param {string} subjectDetails A brief description of the error context (e.g., "Dashboard Update Failed").
 * @param {Error} error The JavaScript Error object that was caught.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] The spreadsheet where the error occurred. Defaults to the active spreadsheet if not provided.
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
 * Retrieves, creates, or falls back to the designated log spreadsheet.
 * The function first tries to open the spreadsheet using an ID stored in Script Properties.
 * If that fails, it attempts to create a new spreadsheet in the user's Drive.
 * If creation also fails (e.g., due to insufficient permissions), it defaults to using
 * the currently active spreadsheet as a last resort, ensuring that logging can always proceed.
 *
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The log spreadsheet object.
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

/**
 * Ensures that a sheet for the specified month exists in the log spreadsheet.
 * If the sheet doesn't exist, it creates and formats it with a header row.
 * Sheet names are based on a "YYYY-MM" key.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} logSS The spreadsheet where logs are stored.
 * @param {string} [monthKey] The month key (e.g., "2024-07"). Defaults to the current month.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet for the specified month.
 */
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

/**
 * Writes a detailed audit entry to the appropriate monthly log sheet.
 * This function orchestrates getting the log spreadsheet and the correct monthly sheet,
 * then appends a new row with the provided audit information.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSS The spreadsheet where the audited action occurred.
 * @param {object} entry An object containing the details of the log entry.
 * @param {string} entry.action The name of the action being logged (e.g., "SyncFtoU").
 * @param {string} [entry.sourceSheet] The name of the sheet where the action was initiated.
 * @param {number} [entry.sourceRow] The row number related to the action.
 * @param {string} [entry.projectName] The project name involved in the action.
 * @param {string} [entry.details] A description of what happened.
 * @param {string} [entry.result] The outcome of the action (e.g., "success", "skipped").
 * @param {string} [entry.errorMessage] Any error message if the action failed.
 */
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

    // Sort the sheet to keep the newest entries at the top
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) { // Check if there is data to sort (beyond the header)
      const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
      range.sort({ column: 1, ascending: false }); // Sort by timestamp descending
    }

  } catch (e) {
    Logger.log(`CRITICAL: Audit logging failure: ${e}`);
    // If the logging system itself fails, attempt to notify.
    notifyError("Audit logging system failed critically", e, sourceSS);
  }
}