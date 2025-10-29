const fs = require('fs');
const path = require('path');

// =================================================================
// =================== MOCK GLOBAL OBJECTS =========================
// =================================================================

// Load the mock framework
const mockFramework = fs.readFileSync('tests/mockFramework.js', 'utf8');
eval(mockFramework);

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
    flush: () => {},
    BandingTheme: { LIGHT_GREY: 'LIGHT_GREY' },
    BorderStyle: { SOLID_THIN: 'SOLID_THIN' }
};

global.Charts = {
  ChartHiddenDimensionStrategy: {
    IGNORE_ROWS: 'IGNORE_ROWS'
  }
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

// =================================================================
// =================== NEW TEST HARNESS ============================
// =================================================================

// Snap a pristine copy of CONFIG to restore before each suite.
const BASE_CONFIG_JSON = JSON.stringify(global.CONFIG || {});

/** Reset global.CONFIG to its original value */
function resetConfig() {
  global.CONFIG = JSON.parse(BASE_CONFIG_JSON);
  // After parsing, the date objects are strings, so we need to convert them back.
  if (global.CONFIG.DASHBOARD_DATES) {
    global.CONFIG.DASHBOARD_DATES.START = new Date(global.CONFIG.DASHBOARD_DATES.START);
    global.CONFIG.DASHBOARD_DATES.END = new Date(global.CONFIG.DASHBOARD_DATES.END);
  }
}

/** Deep freeze to detect any accidental writes during a suite */
function deepFreeze(o) {
  if (!o || typeof o !== 'object' || Object.isFrozen(o)) return o;
  Object.freeze(o);
  for (const k of Object.getOwnPropertyNames(o)) {
    const v = o[k];
    if (v && typeof v === 'object' && !Object.isFrozen(v)) deepFreeze(v);
  }
  return o;
}

/** Optional: diff helper to print changes to CONFIG after a suite */
function diffConfig(a, b, prefix = '') {
  const lines = [];
  const keys = new Set([...Object.keys(a || {}), ...Object.keys(b || {})]);
  for (const k of keys) {
    const pa = a ? a[k] : undefined;
    const pb = b ? b[k] : undefined;
    const pfx = prefix ? `${prefix}.${k}` : k;
    const ta = Object.prototype.toString.call(pa);
    const tb = Object.prototype.toString.call(pb);
    if (ta === '[object Object]' && tb === '[object Object]') {
      lines.push(...diffConfig(pa, pb, pfx));
    } else if (JSON.stringify(pa) !== JSON.stringify(pb)) {
      lines.push(`CONFIG changed at ${pfx}: ${JSON.stringify(pa)} -> ${JSON.stringify(pb)}`);
    }
  }
  return lines;
}

/**
 * Evaluate a test file inside an IIFE so top-level const/let (e.g. `const CONFIG = ...`)
 * do not leak into the shared VM. Also export run* test functions to global.
 */
function evalTestFile(absPath) {
  let code = fs.readFileSync(absPath, 'utf8');

  // Export any function declared as `function runXxx(...)` to global
  code = code.replace(
    /(^|\n)\s*function\s+(run[A-Za-z0-9_]+)\s*\(/g,
    (m, lead, name) => `${lead}global.${name} = function ${name}(`
  );

    // Export any function declared as `function testXxx(...)` to global
  code = code.replace(
    /(^|\n)\s*function\s+(test[A-Za-z0-9_]+)\s*\(/g,
    (m, lead, name) => `${lead}global.${name} = function ${name}(`
  );

  // Wrap in an IIFE to contain top-level bindings
  const wrapped = `(function(){\n${code}\n})();`;
  eval(wrapped);
}

/** Run a suite with an automatic CONFIG reset and freeze */
function runSuite(name, fn) {
  resetConfig();
  //deepFreeze(global.CONFIG); // Temporarily disabled to allow for mocking
  const before = JSON.parse(JSON.stringify(global.CONFIG));
  console.log(`\n--- Running ${name} tests ---`);
  try {
    fn();
  } finally {
    const after = JSON.parse(JSON.stringify(global.CONFIG));
    const changes = diffConfig(before, after);
    if (changes.length) {
      console.warn(`\n[WARN] ${name} modified CONFIG:\n` + changes.map(s => '  - ' + s).join('\n'));
    }
  }
}

// =================================================================
// ======================= TEST EXECUTION ==========================
// =================================================================

try {
    // Load all test files using the new sandboxed loader
    const testsDir = path.resolve(__dirname, 'tests');
    evalTestFile(path.join(testsDir, 'bugfix/test_findRowByValue_case_insensitivity.gs'));
    evalTestFile(path.join(testsDir, 'bugfix-robust-find-test.gs'));
    evalTestFile(path.join(testsDir, 'test_Utilities.gs'));
    evalTestFile(path.join(testsDir, 'chart_title.test.gs'));
    evalTestFile(path.join(testsDir, 'test_Dashboard.gs'));
    evalTestFile(path.join(testsDir, 'test_AuditLogging.gs'));
    evalTestFile(path.join(testsDir, 'test_Dashboard_HoverNotes.gs'));
    evalTestFile(path.join(testsDir, 'test_TransferEngine.gs'));
    evalTestFile(path.join(testsDir, 'test_TransferEngine_Sort.gs'));
    evalTestFile(path.join(testsDir, 'test_findRowByValue.gs'));
    evalTestFile(path.join(testsDir, 'test_ErrorHandling.gs'));
    evalTestFile(path.join(testsDir, 'test_TransferEngine_ReadWidth.gs'));
    evalTestFile(path.join(testsDir, 'bugfix/DateAsName.test.gs'));
    evalTestFile(path.join(testsDir, 'test_inspections_email.gs'));
    evalTestFile(path.join(testsDir, 'test_Upcoming_SyncOnDuplicate.gs'));
    evalTestFile(path.join(testsDir, 'test_Upcoming_FallbackCompoundKey.gs'));
    evalTestFile(path.join(testsDir, 'test_UpdateRow_NonDestructive.gs'));


    // Run suites with isolation
    runSuite('Date-As-Name Bugfix', () => global.runDateAsNameBugfixTests && runDateAsNameBugfixTests());
    runSuite('Robust Find', () => global.runRobustFindTest && runRobustFindTest());
    runSuite('Utility', () => global.runUtilityTests && runUtilityTests());
    runSuite('Chart Title', () => global.runChartTitleTests && runChartTitleTests());
    runSuite('Dashboard', () => global.test_nonCompleteProjectWithPastDeadline_isCountedAsOverdue && test_nonCompleteProjectWithPastDeadline_isCountedAsOverdue());
    runSuite('Audit Logging', () => global.runAuditLoggingTests && runAuditLoggingTests());
    runSuite('Dashboard Hover Notes', () => global.test_Dashboard_HoverNotes && test_Dashboard_HoverNotes());
    runSuite('Transfer Engine', () => global.runTransferEngineTests && runTransferEngineTests());
    runSuite('Transfer Engine Sort', () => global.runTransferEngineSortTests && runTransferEngineSortTests());
    runSuite('Find Row By Value', () => global.runFindRowByValueTests && runFindRowByValueTests());
    runSuite('Error Handling', () => global.runErrorHandlingTests && runErrorHandlingTests());
    runSuite('Transfer Engine Read Width', () => global.runTransferEngineReadWidthTests && runTransferEngineReadWidthTests());
    runSuite('Inspection Email', () => global.test_sendInspectionEmail_sendsCorrectEmail && test_sendInspectionEmail_sendsCorrectEmail());
    runSuite('Upcoming Sync on Duplicate', () => global.run_test_upcoming_sync_on_duplicate && run_test_upcoming_sync_on_duplicate());
    runSuite('Upcoming Fallback Compound Key', () => global.run_test_upcoming_fallback_compound_key && run_test_upcoming_fallback_compound_key());
    runSuite('Update Row Non-Destructive', () => global.run_test_update_row_non_destructive && run_test_update_row_non_destructive());
    runSuite('Bugfix', () => global.runBugfixTests && runBugfixTests());

    console.log("\nTest execution finished successfully.");
} catch (e) {
    console.error("\nTest failed:", e.message);
    process.exit(1); // Exit with an error code
}
