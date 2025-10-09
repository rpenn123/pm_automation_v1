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
 * Retrieves or creates the designated log spreadsheet.
 * This function is resilient, attempting to open by ID, then create, but will throw
 * a `DependencyError` if it cannot secure a log spreadsheet, removing the fallback to the active sheet.
 *
 * @param {object} config The global configuration object (`CONFIG`).
 * @param {string} correlationId The correlation ID for tracing.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The log spreadsheet object.
 * @throws {DependencyError} If the log spreadsheet cannot be opened or created.
 */
function getOrCreateLogSpreadsheet(config, correlationId) {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(config.LOGGING.SPREADSHEET_ID_PROP);

  // 1. Try opening the stored ID.
  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (e) {
      // The stored ID is invalid. Log this as a warning, clear the property, and proceed to create a new one.
      Logger.log(`Stored log spreadsheet ID '${storedId}' is invalid or inaccessible. A new log sheet will be created. Error: ${e.message}`);
      props.deleteProperty(config.LOGGING.SPREADSHEET_ID_PROP);
    }
  }

  // 2. Try creating a new external log spreadsheet.
  try {
    const newSS = SpreadsheetApp.create(config.LOGGING.SPREADSHEET_NAME);
    props.setProperty(config.LOGGING.SPREADSHEET_ID_PROP, newSS.getId());
    return newSS;
  } catch (e2) {
    // If creation fails, this is a critical, unrecoverable error for logging.
    throw new DependencyError("Failed to create a new log spreadsheet.", e2);
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
function ensureMonthlyLogSheet(logSS, monthKey) {
  // Use the padded month key for standardized log sheet names
  const key = monthKey || getMonthKeyPadded();
  let sh = logSS.getSheetByName(key);
  if (!sh) {
    sh = logSS.insertSheet(key);
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
    const logSS = getOrCreateLogSpreadsheet(config, entry.correlationId);
    const sheet = ensureMonthlyLogSheet(logSS);
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

    sheet.appendRow(newRow);

    // Sorting is now handled by a separate trigger to reduce latency in the main execution path.
    // See `sortLogSheetsOnOpen`.

  } catch (e) {
    // If the logging system itself fails, use the centralized handler.
    handleError(new DependencyError("Audit logging system failed critically.", e), {
      correlationId: entry.correlationId,
      functionName: "logAudit",
      spreadsheet: sourceSS,
      extra: { originalEntry: entry }
    }, config);
  }
}

/**
 * Sorts all monthly log sheets within the designated log spreadsheet.
 * This function is designed to be run by an `onOpen` trigger in the log spreadsheet itself. It iterates through all sheets,
 * finds any that match the "YYYY-MM" log sheet format, and sorts them by timestamp (column 1) in descending order.
 * This action ensures that the latest logs are always at the top and easy to review.
 *
 * @returns {void} This function does not return a value.
 */
function sortLogSheetsOnOpen() {
  try {
    const logSS = getOrCreateLogSpreadsheet(CONFIG);
    if (!logSS) {
      Logger.log("Auto-Sort: Could not retrieve the log spreadsheet. Aborting sort.");
      return;
    }

    const sheets = logSS.getSheets();
    const monthKeyRegex = /^\d{4}-\d{2}$/; // Regex to identify "YYYY-MM" sheet names

    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      // Check if the sheet name matches the monthly log format
      if (monthKeyRegex.test(sheetName)) {
        const lastRow = sheet.getLastRow();

        // Only sort if there's more than just a header row
        if (lastRow > 1) {
          // Define the range to be sorted, excluding the header row
          const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
          // Sort by the first column (Timestamp) in descending order (newest first)
          range.sort({ column: 1, ascending: false });
          Logger.log(`Auto-Sort: Successfully sorted sheet "${sheetName}".`);
        }
      }
    });
  } catch (e) {
    // We avoid calling notifyError here to prevent a potential infinite loop if the error
    // is related to accessing the spreadsheet, which could trigger more logging.
    Logger.log(`CRITICAL: Auto-Sort for log sheets failed. Error: ${e.message}`);
  }
}