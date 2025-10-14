/**
 * Test Suite for Error Handling
 *
 * @version 1.0.0
 * @release 2025-10-08
 */

function runErrorHandlingTests() {
    console.log("Running test suite: Error Handling");

    // Mock necessary global objects and functions for this test suite
    const originalLogger = global.Logger;
    const originalUtilities = global.Utilities;
    const originalPropertiesService = global.PropertiesService;
    const originalMailApp = global.MailApp;
    const originalSpreadsheetApp = global.SpreadsheetApp;
    const originalNotifyError = global.notifyError;
    const originalConfig = global.CONFIG;

    // Setup mocks
    global.Logger = { log: (message) => {} }; // Suppress logs for most tests
    global.Utilities = { ...global.Utilities, sleep: (ms) => {} };
    global.PropertiesService = { getScriptProperties: () => ({ getProperty: (key) => 'test.email@example.com' }) };
    global.MailApp = { sendEmail: (options) => {} };
    global.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getId: () => 'mock-spreadsheet-id', getName: () => 'MockSpreadsheet', getUrl: () => 'http://mock.url' })};
    global.notifyError = () => {}; // Mock notifyError to isolate handleError

    try {
        testWithRetry();
        testHandleError();
        testCustomErrorClasses();
    } finally {
        // Restore original globals
        global.Logger = originalLogger;
        global.Utilities = originalUtilities;
        global.PropertiesService = originalPropertiesService;
        global.MailApp = originalMailApp;
        global.SpreadsheetApp = originalSpreadsheetApp;
        global.notifyError = originalNotifyError;
        global.CONFIG = originalConfig;
    }
}

/**
 * Test function for the `withRetry` utility.
 */
function testWithRetry() {
    let attempts = 0;
    const failingFunction = () => {
        attempts++;
        throw new Error("Transient failure");
    };

    try {
        withRetry(failingFunction, { functionName: "testRetry", maxRetries: 3 });
    } catch (e) {
        if (e.name !== 'DependencyError') {
            throw new Error(`testWithRetry Failed: Expected DependencyError, but got ${e.name}`);
        }
        if (attempts !== 3) {
            throw new Error(`testWithRetry Failed: Expected 3 attempts, but got ${attempts}`);
        }
        console.log("  -> Passed: testWithRetry");
        return;
    }
    throw new Error("testWithRetry Failed: Expected withRetry to throw an error, but it did not.");
}

/**
 * Test function for the `handleError` function.
 */
function testHandleError() {
    let loggedMessage = "";
    const tempLogger = { log: (message) => { loggedMessage = message; } };
    const originalLogger = global.Logger;
    global.Logger = tempLogger;

    try {
        const error = new ValidationError("Invalid input.", new Error("Root cause"));
        const context = {
            correlationId: "test-corr-id",
            functionName: "testFunction",
            spreadsheet: global.SpreadsheetApp.getActiveSpreadsheet(),
            extra: { someValue: "test" }
        };

        handleError(error, context, global.CONFIG);

        const logData = JSON.parse(loggedMessage);

        if (logData.severity !== 'WARN' || logData.error.type !== 'ValidationError' || logData.correlationId !== 'test-corr-id') {
            throw new Error("Logged data does not match expected format.");
        }
        console.log("  -> Passed: testHandleError");
    } catch(e) {
        throw new Error(`testHandleError Failed: ${e.message}`);
    }
    finally {
        global.Logger = originalLogger; // Restore original logger
    }
}

/**
 * Test function to verify custom error class constructors.
 */
function testCustomErrorClasses() {
    const cause = new Error("The root cause");
    const validationError = new ValidationError("Validation failed", cause);
    const dependencyError = new DependencyError("Dependency failed", cause);
    const transientError = new TransientError("Transient failure", cause);
    const configError = new ConfigurationError("Config is wrong");

    if (validationError.name !== 'ValidationError' || !validationError.cause || validationError.cause.message !== 'The root cause') {
        throw new Error('ValidationError constructor failed.');
    }
    if (dependencyError.name !== 'DependencyError' || !dependencyError.cause || dependencyError.cause.message !== 'The root cause') {
        throw new Error('DependencyError constructor failed.');
    }
    if (transientError.name !== 'TransientError' || !transientError.cause || transientError.cause.message !== 'The root cause') {
        throw new Error('TransientError constructor failed.');
    }
    if (configError.name !== 'ConfigurationError' || configError.cause !== undefined) {
        throw new Error('ConfigurationError constructor failed.');
    }
    console.log("  -> Passed: testCustomErrorClasses");
}