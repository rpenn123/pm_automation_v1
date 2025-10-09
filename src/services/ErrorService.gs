/**
 * @OnlyCurrentDoc
 *
 * ErrorService.gs
 *
 * A centralized service for handling, logging, and reporting errors in a standardized way.
 * This service implements the strategy defined in the Error Handling RFC (docs/ERRORS.md).
 *
 * @version 1.0.0
 * @release 2025-10-08
 */

// =================================================================
// ==================== CUSTOM ERROR CLASSES =======================
// =================================================================

/**
 * Base error class for all custom application errors.
 * Note: Google Apps Script's Rhino runtime does not support ES6 classes,
 * so we use constructor functions and prototype inheritance.
 *
 * @param {string} message The error message.
 * @param {Error} [cause] The original error that caused this one.
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
 * Represents an error due to invalid input data. Not retryable.
 * @param {string} message The error message.
 * @param {Error} [cause] The original error.
 */
function ValidationError(message, cause) {
  BaseError.call(this, message, cause);
}
ValidationError.prototype = Object.create(BaseError.prototype);
ValidationError.prototype.constructor = ValidationError;

/**
 * Represents an error from an external service (e.g., Google Sheets API).
 * @param {string} message The error message.
 * @param {Error} [cause] The original error.
 */
function DependencyError(message, cause) {
  BaseError.call(this, message, cause);
}
DependencyError.prototype = Object.create(BaseError.prototype);
DependencyError.prototype.constructor = DependencyError;

/**
 * A sub-type of DependencyError that is known to be temporary and retryable.
 * @param {string} message The error message.
 * @param {Error} [cause] The original error.
 */
function TransientError(message, cause) {
  DependencyError.call(this, message, cause);
}
TransientError.prototype = Object.create(DependencyError.prototype);
TransientError.prototype.constructor = TransientError;

/**
 * Represents an error in the application's configuration. Not retryable.
 * @param {string} message The error message.
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
 * Centralized function for handling all caught errors.
 * It logs the error in a structured format and can send notifications for critical issues.
 *
 * @param {Error} error The error object that was caught.
 * @param {object} context An object containing contextual information.
 * @param {string} context.correlationId A unique ID for tracing the operation.
 * @param {string} context.functionName The name of the function where the error occurred.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [context.spreadsheet] The spreadsheet object.
 * @param {object} [context.extra] Any other relevant details (e.g., sheetName, row, sfid).
 * @param {object} config The global configuration object (`CONFIG`).
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
 * Wraps a function call with a retry mechanism featuring exponential backoff and jitter.
 * This is designed for I/O operations that might fail due to transient issues.
 *
 * @param {function} fn The function to execute and retry on failure.
 * @param {object} options Configuration for the retry mechanism.
 * @param {string} options.functionName A name for the operation, for logging purposes.
 * @param {number} [options.maxRetries=3] The maximum number of retries.
 * @param {number} [options.initialDelayMs=200] The initial delay before the first retry.
 * @returns {any} The return value of the wrapped function if successful.
 * @throws {DependencyError} If the function fails after all retries.
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