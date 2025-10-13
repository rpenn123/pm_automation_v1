/**
 * @OnlyCurrentDoc
 *
 * LoggerService.gs
 *
 * Handles critical error notifications via email and manages the external monthly audit log system.
 * This service ensures that errors are reported to administrators and that all significant
 * actions are recorded for accountability and debugging.
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

/**
 * Sends a formatted email notification when a critical error occurs.
 * This function constructs a detailed HTML email with the error message, stack trace, and spreadsheet context.
 * It includes a plain text fallback for email clients that do not support HTML. It will not send an email
 * if no valid recipient is configured in Script Properties.
 *
 * @param {string} subjectDetails A brief description of the error context (e.g., "Dashboard Update Failed").
 * @param {Error} error The JavaScript `Error` object that was caught.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss] The spreadsheet where the error occurred. Defaults to the active spreadsheet if not provided.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function notifyError(subjectDetails, error, ss, config) {
  const props = PropertiesService.getScriptProperties();
  const email = props.getProperty(config.LOGGING.ERROR_EMAIL_PROP);
  const appName = config.APP_NAME;

  // Validate email format
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) {
    Logger.log(`Error email skipped: No valid notification email has been set. Please run the setup menu to configure it.`);
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
    // If sending the email fails, we now use the centralized handler.
    handleError(new DependencyError("Failed to send error notification email.", mailError), {
      correlationId: "N/A",
      functionName: "notifyError",
      spreadsheet: ss
    }, config);
  }
}

/**
 * Ensures that a sheet for the specified month exists in the log spreadsheet.
 * If the sheet doesn't exist, it creates and formats it with a frozen header row,
 * now including a `CorrelationId` column for traceability.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} logSS The spreadsheet where logs are stored.
 * @param {string} [monthKey] The month key to use (e.g., "2024-07"). Defaults to the current month if not provided.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object for the specified month.
 */
function ensureMonthlyLogSheet(sourceSS, monthKey) {
  // Use the padded month key for standardized log sheet names
  const key = `Logs ${monthKey || getMonthKeyPadded()}`;
  let sh = sourceSS.getSheetByName(key);
  if (!sh) {
    sh = sourceSS.insertSheet(key);
    // Set headers and freeze the first row, now with CorrelationId.
    sh.getRange(1, 1, 1, 12).setValues([[
      "Timestamp", "CorrelationId", "User", "Action",
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
 * This function now requires a `correlationId` and includes it in the log entry,
 * enhancing traceability across all logged actions.
 *
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSS The spreadsheet where the audited action occurred.
 * @param {object} entry An object containing the details of the log entry.
 * @param {string} entry.correlationId A unique ID to trace the entire operation.
 * @param {string} entry.action The name of the action being logged (e.g., "SyncFtoU").
 * @param {string} [entry.sourceSheet] The name of the sheet where the action was initiated.
 * @param {number} [entry.sourceRow] The row number related to the action.
 * @param {string} [entry.projectName] The project name involved in the action.
 * @param {string} [entry.details] A description of what happened.
 * @param {string} [entry.result] The outcome of the action (e.g., "success", "skipped", "error").
 * @param {string} [entry.errorMessage] Any error message if the action failed.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function logAudit(sourceSS, entry, config) {
  if (!entry.correlationId) {
    // Enforce correlationId for traceability.
    Logger.log("CRITICAL: logAudit called without a correlationId. Entry: " + JSON.stringify(entry));
    return;
  }

  try {
    Logger.log(`DEBUG: logAudit started. Correlation ID: ${entry.correlationId}`);
    Logger.log(`DEBUG: sourceSS object is present: ${!!sourceSS}`);
    Logger.log(`DEBUG: entry object: ${JSON.stringify(entry, null, 2)}`);

    const sheet = ensureMonthlyLogSheet(sourceSS);
    const user = Session.getActiveUser() ? Session.getActiveUser().getEmail() : "unknown";

    const newRow = [
      new Date(),
      entry.correlationId,
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

    Logger.log(`DEBUG: newRow content: ${JSON.stringify(newRow, null, 2)}`);
    sheet.appendRow(newRow);
    Logger.log("DEBUG: appendRow successfully completed.");

  } catch (e) {
    // If the logging system itself fails, use the centralized handler.
    Logger.log(`DEBUG: Caught error in logAudit. Raw Error: ${e.message}\nStack: ${e.stack}`);
    handleError(new DependencyError("Audit logging system failed critically.", e), {
      correlationId: entry.correlationId,
      functionName: "logAudit",
      spreadsheet: sourceSS,
      extra: { originalEntry: entry }
    }, config);
  }
}
