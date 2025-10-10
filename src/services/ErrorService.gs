/**
 * @OnlyCurrentDoc
 *
 * ErrorService.gs
 *
 * A centralized service for handling, logging, and reporting errors in a standardized way.
 * This service implements the strategy defined in the Error Handling RFC (docs/ERRORS.md).
 *
 * @version 1.1.0
 * @release 2025-10-08
 */

// =================================================================
// ==================== CUSTOM ERROR CLASSES =======================
// =================================================================

/**
 * Base error class for all custom application errors, providing a common foundation for naming and stack tracing.
 * Note: Google Apps Script's Rhino runtime does not fully support ES6 classes,
 * so this project uses constructor functions and prototype inheritance for custom errors.
 *
 * @param {string} message The primary error message.
 * @param {Error} [cause] The underlying error that caused this one, used for chaining.
 * @constructor
 */
function BaseError(message, cause) {
  this.name = this.constructor.name;
  this.message = message;
  this.stack = (new Error()).stack;
  this.cause = cause;
}
BaseError.prototype = Object.create(Error.prototype);
BaseError.prototype.constructor = BaseError;

/**
 * Represents an error due to invalid or missing input data (e.g., a required SFID is missing).
 * These errors are generally not retryable as they indicate a data problem that needs manual correction.
 *
 * @param {string} message The validation-specific error message.
 * @param {Error} [cause] The underlying error, if any.
 * @constructor
 */
function ValidationError(message, cause) {
  BaseError.call(this, message, cause);
}
ValidationError.prototype = Object.create(BaseError.prototype);
ValidationError.prototype.constructor = ValidationError;

/**
 * Represents an error originating from an external service or dependency (e.g., Google Sheets API, LockService).
 * These may or may not be retryable. For explicitly retryable dependency errors, use `TransientError`.
 *
 * @param {string} message The dependency-related error message.
 * @param {Error} [cause] The original error from the external service.
 * @constructor
 */
function DependencyError(message, cause) {
  BaseError.call(this, message, cause);
}
DependencyError.prototype = Object.create(BaseError.prototype);
DependencyError.prototype.constructor = DependencyError;

/**
 * A sub-type of `DependencyError` that is known to be temporary and suitable for retrying.
 * This is used for issues like API rate limits, network timeouts, or temporary lock contention.
 *
 * @param {string} message The transient error message.
 * @param {Error} [cause] The original, underlying transient error.
 * @constructor
 */
function TransientError(message, cause) {
  DependencyError.call(this, message, cause);
}
TransientError.prototype = Object.create(DependencyError.prototype);
TransientError.prototype.constructor = TransientError;

/**
 * Represents an error in the application's configuration (e.g., a required sheet name in `Config.gs` is incorrect).
 * These errors are not retryable and indicate a problem that requires a code or configuration fix.
 *
 * @param {string} message The configuration-specific error message.
 * @constructor
 */
function ConfigurationError(message) {
  BaseError.call(this, message);
}
ConfigurationError.prototype = Object.create(BaseError.prototype);
ConfigurationError.prototype.constructor = ConfigurationError;


// =================================================================
// ==================== CENTRALIZED ERROR HANDLER ==================
// =================================================================

/**
 * Centralized function for handling all caught errors throughout the application.
 * It logs the error in a structured JSON format to the script's execution logs and sends an email
 * notification via `notifyError` for critical issues. Non-critical errors, like `ValidationError`,
 * are logged as warnings and do not trigger notifications.
 *
 * @param {Error} error The error object that was caught.
 * @param {object} context An object containing contextual information about the error's origin.
 * @param {string} context.correlationId A unique ID for tracing the entire operation end-to-end.
 * @param {string} context.functionName The name of the function where the error was caught.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [context.spreadsheet] The spreadsheet where the error occurred.
 * @param {object} [context.extra] A free-form object for any other relevant details (e.g., sheetName, row, sfid).
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function handleError(error, context, config) {
  try {
    const isCritical = !(error instanceof ValidationError); // Example: Don't notify for simple validation errors.

    // 1. Log the error in a structured format.
    const logEntry = {
      timestamp: new Date().toISOString(),
      severity: isCritical ? "ERROR" : "WARN",
      message: error.message,
      correlationId: context.correlationId,
      error: {
        type: error.name || "Error",
        message: error.message,
        stack: error.stack,
        cause: error.cause ? { type: error.cause.name, message: error.cause.message, stack: error.cause.stack } : null
      },
      context: {
        appName: config.APP_NAME,
        functionName: context.functionName,
        spreadsheetId: context.spreadsheet ? context.spreadsheet.getId() : "N/A",
        extra: context.extra || {}
      }
    };

    // Use Logger.log for structured logging. In a real scenario, this would go to a proper logging service.
    Logger.log(JSON.stringify(logEntry));

    // 2. Send notification for critical errors.
    if (isCritical) {
      // This reuses the existing notifyError logic but could be replaced.
      const subject = `[${error.name || 'Error'}] ${context.functionName} failed`;
      notifyError(subject, error, context.spreadsheet, config);
    }

  } catch (loggingError) {
    // If the error handling itself fails, log to the basic logger.
    Logger.log(`CRITICAL: Error in handleError function: ${loggingError}. Original error: ${error.message}`);
  }
}

// =================================================================
// ==================== I/O RESILIENCE UTILITIES ===================
// =================================================================

/**
 * Wraps a function call with a robust retry mechanism, featuring exponential backoff and jitter.
 * This utility is essential for making I/O operations (like API calls to Google Sheets) more resilient
 * to transient issues such as network flakes or temporary API unavailability. It will not retry on
 * non-transient errors like `ValidationError` or `ConfigurationError`.
 *
 * @param {function(): any} fn The function to execute. It should return a value or throw an error on failure.
 * @param {object} options Configuration options for the retry behavior.
 * @param {string} options.functionName A descriptive name for the operation, used in log messages.
 * @param {number} [options.maxRetries=3] The maximum number of times to retry the function.
 * @param {number} [options.initialDelayMs=200] The base delay in milliseconds for the first retry. Subsequent retries use exponential backoff.
 * @returns {any} The return value of the wrapped function (`fn`) if it succeeds.
 * @throws {DependencyError} If the function continues to fail after all retry attempts have been exhausted.
 * @throws {ValidationError|ConfigurationError} If the function throws a non-retryable error, it is re-thrown immediately without a retry attempt.
 */
function withRetry(fn, options) {
  const { functionName, maxRetries = 3, initialDelayMs = 200 } = options;
  let attempt = 0;

  while (attempt < maxRetries) {
    try {
      return fn();
    } catch (e) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new DependencyError(`'${functionName}' failed after ${maxRetries} attempts.`, e);
      }

      // Don't retry on non-transient errors
      if (e instanceof ValidationError || e instanceof ConfigurationError) {
        throw e;
      }

      // Calculate delay with exponential backoff and jitter
      const backoff = Math.pow(2, attempt) * initialDelayMs;
      const jitter = backoff * 0.2 * Math.random();
      const delay = backoff + jitter;

      Logger.log(`Attempt ${attempt} for '${functionName}' failed. Retrying in ${Math.round(delay)}ms. Error: ${e.message}`);
      Utilities.sleep(delay);
    }
  }
}