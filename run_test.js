const fs = require('fs');

// =================================================================
// =================== MOCK GLOBAL OBJECTS =========================
// =================================================================

// Mock the global objects that the .gs files expect to exist.
global.Logger = {
    log: (message) => console.log(`[Logger] ${message}`)
};

global.SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
        // Mock any methods on the Spreadsheet object if needed
    })
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

// This mock is needed by the code under test in Utilities.gs
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
  sleep: (ms) => { /* no-op for tests */ }
};

// This mock is also needed by Utilities.gs
global.Session = {
  getScriptTimeZone: () => "America/New_York",
};


// Mock Range object for TextFinder simulation (not needed for the fix, but for testing the old code)
global.MockRange = class {
    constructor(data, startRow) {
        this.data = data;
        this.startRow = startRow;
    }
    getValues() {
        return this.data;
    }
    // This method is called by the old buggy code.
    createTextFinder(text) {
        // This functionality is not needed for the fixed code, so we can return a dummy object.
        return {
            matchCase: () => this,
            matchEntireCell: () => this,
            findNext: () => null
        };
    }
};


// =================================================================
// ======================= SCRIPT LOADER ===========================
// =================================================================

// Load the script content from the .gs files
const utilitiesGs = fs.readFileSync('src/core/Utilities.gs', 'utf8');
let configGs = fs.readFileSync('src/Config.gs', 'utf8');
const dashboardGs = fs.readFileSync('src/ui/Dashboard.gs', 'utf8');
const testGs = fs.readFileSync('tests/bugfix-robust-find-test.gs', 'utf8');
const existingTestGs = fs.readFileSync('tests/test_Utilities.gs', 'utf8');
const chartTitleTestGs = fs.readFileSync('tests/chart_title.test.gs', 'utf8');

// Make CONFIG global for tests
configGs = configGs.replace('const CONFIG =', 'global.CONFIG =');

// Use 'eval' to make the functions available in the current scope.
eval(utilitiesGs);
eval(configGs);
eval(dashboardGs);
eval(testGs);
eval(existingTestGs);
eval(chartTitleTestGs);

// =================================================================
// ======================= TEST EXECUTION ==========================
// =================================================================

// Run the tests
try {
    console.log("--- Running new bugfix test ---");
    runRobustFindTest();
    console.log("\n--- Running existing utility tests ---");
    runUtilityTests();
    console.log("\n--- Running chart title tests ---");
    runChartTitleTests();
    console.log("\nTest execution finished successfully.");
} catch (e) {
    console.error("\nTest failed:", e.message);
    process.exit(1); // Exit with an error code
}