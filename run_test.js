const fs = require('fs');

// =================================================================
// =================== MOCK GLOBAL OBJECTS =========================
// =================================================================

// Mock the global objects that the .gs files expect to exist.
global.Logger = {
    log: (message) => console.log(`[Logger] ${message}`)
};

global.PropertiesService = {
    getScriptProperties: () => ({
        getProperty: () => null,
        setProperty: () => {},
    })
};

global.MailApp = {
    sendEmail: (options) => {
        console.log(`[MailApp] Email sent: ${JSON.stringify(options)}`);
    }
};

global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
        // This is a simplified mock. Tests that need more specific behavior
        // will define their own mocks.
        getSheetByName: () => ({
            getRange: () => ({
                getValues: () => [[]],
                getValue: () => '',
            }),
            getMaxColumns: () => 10,
        })
    }),
};

global.notifyError = (subject, error, ss) => {
    const errorMessage = error ? error.message : subject;
    console.error(`[notifyError] Subject: ${subject}`);
    if (error) {
        console.error(`[notifyError] Error: ${errorMessage}`);
    }
    // Throw an error that the test runner can catch to fail the test.
    throw new Error(errorMessage);
};

global.Utilities = {
  formatDate: (date, timeZone, format) => {
    if (!date || !(date instanceof Date)) return "";
    if (format === "yyyy-MM-dd") {
      const d = new Date(date);
      const year = d.getUTCFullYear();
      const month = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    if (format === "yyyy-MM-dd'T'HH:mm:ss") {
        const pad = (num) => String(num).padStart(2, '0');
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}'T'${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
    }
    return date.toString();
  },
  sleep: (ms) => { /* no-op for tests */ },
  getUuid: () => 'mock-uuid-' + Math.random().toString(36).substring(2)
};

global.Session = {
  getScriptTimeZone: () => "America/New_York",
  getActiveUser: () => ({
      getEmail: () => 'test.user@example.com'
  })
};

global.LockService = {
    getScriptLock: () => ({
        tryLock: () => true,
        releaseLock: () => {},
    }),
};

// More robust mock for Jest functions to support implementation and restoration.
const jest = {
    spyOn: (obj, funcName) => {
        const originalFunc = obj[funcName];
        let implementation = originalFunc;

        const spy = {
            mock: {
                calls: []
            },
            mockImplementation: function(fn) {
                implementation = fn;
                return this;
            },
            mockRestore: function() {
                obj[funcName] = originalFunc;
            }
        };

        obj[funcName] = (...args) => {
            spy.mock.calls.push(args);
            return implementation.apply(obj, args);
        };

        return spy;
    },
    fn: (implementation) => {
        const mock = (...args) => {
            mock.mock.calls.push(args);
            if (implementation) {
                return implementation(...args);
            }
        };
        mock.mock = {
            calls: []
        };
        return mock;
    }
};
global.jest = jest;


// =================================================================
// ======================= SCRIPT LOADER ===========================
// =================================================================

// Load the script content from the .gs files
const utilitiesGs = fs.readFileSync('src/core/Utilities.gs', 'utf8');
let configGs = fs.readFileSync('src/Config.gs', 'utf8');
let dashboardGs = fs.readFileSync('src/ui/Dashboard.gs', 'utf8');
let lastEditServiceGs = fs.readFileSync('src/services/LastEditService.gs', 'utf8');
let loggerServiceGs = fs.readFileSync('src/services/LoggerService.gs', 'utf8');
let errorServiceGs = fs.readFileSync('src/services/ErrorService.gs', 'utf8');
const automationsGs = fs.readFileSync('src/core/Automations.gs', 'utf8');
let transferEngineGs = fs.readFileSync('src/core/TransferEngine.gs', 'utf8');

// Load test files
const testGs = fs.readFileSync('tests/bugfix-robust-find-test.gs', 'utf8');
const existingTestGs = fs.readFileSync('tests/test_Utilities.gs', 'utf8');
const chartTitleTestGs = fs.readFileSync('tests/chart_title.test.gs', 'utf8');
const dashboardTestGs = fs.readFileSync('tests/test_Dashboard.gs', 'utf8');
const auditTestGs = fs.readFileSync('tests/test_AuditLogging.gs', 'utf8');
const hoverNotesTestGs = fs.readFileSync('tests/test_Dashboard_HoverNotes.gs', 'utf8');
const transferEngineTestGs = fs.readFileSync('tests/test_TransferEngine.gs', 'utf8');
const transferEngineSortTestGs = fs.readFileSync('tests/test_TransferEngine_Sort.gs', 'utf8');
const findRowByValueTestGs = fs.readFileSync('tests/test_findRowByValue.gs', 'utf8');
const errorHandlingTestGs = fs.readFileSync('tests/test_ErrorHandling.gs', 'utf8');
const transferEngineReadWidthTestGs = fs.readFileSync('tests/test_TransferEngine_ReadWidth.gs', 'utf8');
const dateAsNameBugfixTestGs = fs.readFileSync('tests/bugfix/DateAsName.test.gs', 'utf8');

// Make CONFIG global for tests
configGs = configGs.replace('const CONFIG =', 'global.CONFIG =');
// Make logAudit and notifyError global for spying
loggerServiceGs = loggerServiceGs.replace('function logAudit(', 'global.logAudit = function logAudit(');
loggerServiceGs = loggerServiceGs.replace('function notifyError(', 'global.notifyError = function notifyError(');
// Make readForecastingData global for mocking in tests
dashboardGs = dashboardGs.replace('function readForecastingData(', 'global.readForecastingData = function readForecastingData(');
// Make updateLastEditForRow global for mocking in tests
lastEditServiceGs = lastEditServiceGs.replace('function updateLastEditForRow(', 'global.updateLastEditForRow = function updateLastEditForRow(');

// Use 'eval' to make the functions available in the current scope.
eval(utilitiesGs);
eval(configGs);
eval(dashboardGs);
eval(lastEditServiceGs);
eval(loggerServiceGs);
eval(errorServiceGs); // Must be loaded before other services that use it
eval(automationsGs);
transferEngineGs = transferEngineGs.replace('function executeTransfer(', 'global.executeTransfer = function executeTransfer(');
transferEngineGs = transferEngineGs.replace('function isDuplicateInDestination(', 'global.isDuplicateInDestination = function isDuplicateInDestination(');
eval(transferEngineGs); // TransferEngine is needed by Automations

eval(testGs);
eval(existingTestGs);
eval(chartTitleTestGs);
eval(dashboardTestGs);
eval(auditTestGs);
eval(hoverNotesTestGs);
eval(transferEngineTestGs);
eval(transferEngineSortTestGs);
eval(findRowByValueTestGs);
eval(errorHandlingTestGs);
eval(transferEngineReadWidthTestGs);
eval(dateAsNameBugfixTestGs);

// =================================================================
// ======================= TEST EXECUTION ==========================
// =================================================================

// Run the tests
try {
    console.log("\n--- Running Date-As-Name Bugfix tests ---");
    runDateAsNameBugfixTests();
    console.log("\n--- Running new bugfix test ---");
    runRobustFindTest();
    console.log("\n--- Running existing utility tests ---");
    runUtilityTests();
    console.log("\n--- Running chart title tests ---");
    runChartTitleTests();
    console.log("\n--- Running dashboard tests ---");
    test_nonCompleteProjectWithPastDeadline_isCountedAsOverdue();
    console.log("\n--- Running audit logging tests ---");
    runAuditLoggingTests();
    console.log("\n--- Running Dashboard Hover Notes test ---");
    test_Dashboard_HoverNotes();
    console.log("\n--- Running Transfer Engine tests ---");
    runTransferEngineTests();
    console.log("\n--- Running Transfer Engine Sort tests ---");
    runTransferEngineSortTests();
    console.log("\n--- Running findRowByValue tests ---");
    runFindRowByValueTests();
    console.log("\n--- Running Error Handling tests ---");
    runErrorHandlingTests();
    console.log("\n--- Running Transfer Engine Read Width tests ---");
    runTransferEngineReadWidthTests();
    console.log("\nTest execution finished successfully.");
} catch (e) {
    console.error("\nTest failed:", e.message);
    process.exit(1); // Exit with an error code
}