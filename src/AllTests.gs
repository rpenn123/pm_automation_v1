/**
 * @OnlyCurrentDoc
 * AllTests.gs
 * This file contains all unit and integration tests for the project.
 * It includes a master test runner that can be executed from a custom menu.
 *
 * NOTE ON SCOPE: In Google Apps Script, all .gs files in a project share a single
 * global scope. This means functions defined in other files (e.g., Utilities.gs,
 * TransferEngine.gs) are directly accessible here without any 'import' statements.
 * The test functions can therefore call the functions-under-test directly.
 */

/**
 * Adds a custom menu to the spreadsheet UI to run tests.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Admin Tools')
    .addItem('Run All Tests', 'runAllTests')
    .addToUi();
}

/**
 * A master test runner to execute all tests in the suite.
 * Logs a summary of the results.
 */
function runAllTests() {
  const allTests = [
    // Dashboard.gs
    testProcessForecastingData,
    // LastEditService.gs
    test_generateOptimizedRelativeTimeFormula,
    testUpdateLastEditForRow,
    // LoggerService.gs
    testNotifyError,
    testLogAudit,
    // TransferEngine.gs
    testExecuteTransfer_Success,
    testIsDuplicateInDestination_DateInProjectName,
    testIsDuplicateInDestination_TimezoneBug,
    testIsDuplicateInDestination_CompoundKey,
    testIsDuplicateInDestination_NoDuplicate,
    // Utilities.gs / BugFixes
    testFindRowByProjectName_DateInProjectName,
    testFindRowByValue_WhitespaceBug,
    testNormalizeForComparison,
    testNormalizeString,
    testIsTrueLike,
    testParseAndNormalizeDate,
    testFormatValueForKey,
    testGetMonthKeyPadded,
    testGetHeaderColumnIndex,
    testFindRowByBestIdentifier,
    testCreateMapping,
    testGetMaxValueInObject,
    testUniqueArray
  ];

  let passed = 0;
  let failed = 0;

  console.log("Running all tests...");

  allTests.forEach(function(test) {
    const testName = test.name;
    try {
      test();
      passed++;
    } catch (e) {
      failed++;
      console.error(`ERROR in test ${testName}: ${e.message}\n${e.stack || ''}`);
    }
  });

  console.log("\n----- TEST SUMMARY -----");
  console.log(`Total tests: ${allTests.length}`);
  console.log(`✅ Passed: ${passed}`);
  console.log(`❌ Failed: ${failed}`);
  console.log("----------------------");

  if (failed > 0) {
    throw new Error(`${failed} test(s) failed.`);
  }
}

// =================================================================
// =================== Dashboard.gs Tests ==========================
// =================================================================
function testProcessForecastingData() {
    const testName = "testProcessForecastingData";
    let assertions = 0;
    const originalConfig = (typeof CONFIG !== 'undefined') ? JSON.parse(JSON.stringify(CONFIG)) : undefined;

    try {
        this.CONFIG = {
            FORECASTING_COLS: { DEADLINE: 2, PROGRESS: 3, PERMITS: 4 },
            STATUS_STRINGS: { IN_PROGRESS: "In Progress", SCHEDULED: "Scheduled", PERMIT_APPROVED: "Approved" }
        };
        const today = new Date(); today.setHours(0,0,0,0);
        const futureDate = new Date(today.getTime() + (7 * 24 * 60 * 60 * 1000));
        const pastDate = new Date(today.getTime() - (7 * 24 * 60 * 60 * 1000));
        const forecastingValues = [
            ["Proj 1", futureDate, "Scheduled", ""],
            ["Proj 2", pastDate, "In Progress", ""],
            ["Proj 3", futureDate, "Completed", "Approved"],
            ["Proj 4", pastDate, "In Progress", "Approved"],
            ["Proj 5", null, "In Progress", ""],
            ["Proj 6", futureDate, "In Progress", ""],
            ["Proj 7", today, "In Progress", ""]
        ];
        const result = processForecastingData(forecastingValues);
        const expectedGrandTotals = [6, 2, 3, 2];
        if(JSON.stringify(result.grandTotals) === JSON.stringify(expectedGrandTotals)) assertions++;
        else console.error(`❌ ${testName}: FAILED on grandTotals. Expected ${JSON.stringify(expectedGrandTotals)}, got ${JSON.stringify(result.grandTotals)}`);
        if(result.allOverdueItems.length === 3) assertions++;
        else console.error(`❌ ${testName}: FAILED on allOverdueItems count. Expected 3, got ${result.allOverdueItems.length}`);
        if(result.missingDeadlinesCount === 1) assertions++;
        else console.error(`❌ ${testName}: FAILED on missingDeadlinesCount. Expected 1, got ${result.missingDeadlinesCount}`);
        if (assertions === 3) console.log(`✅ ${testName}: PASSED`);
        else throw new Error(`${testName}: One or more assertions failed.`);
    } finally {
        if (originalConfig) this.CONFIG = originalConfig;
    }
}

// =================================================================
// =================== LastEditService.gs Tests ====================
// =================================================================
function test_generateOptimizedRelativeTimeFormula() {
    const testName = "test_generateOptimizedRelativeTimeFormula";
    const formula = _generateOptimizedRelativeTimeFormula("Z5");
    if (formula.includes('LET(      diff, NOW() - Z5,')) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: Formula does not match expected structure.`);
}

function testUpdateLastEditForRow() {
    const testName = "testUpdateLastEditForRow";
    let assertions = 0;
    let formulaSet = "", valueSet = null;
    const mockSheet = { getRange: (r,c) => ({ getA1Notation: () => `Z${r}`, setValue: (v) => { valueSet = v; }, setFormula: (f) => { formulaSet = f; } }) };
    const originalEnsure = this.ensureLastEditColumns;
    this.ensureLastEditColumns = () => ({ tsCol: 26, relCol: 27 });
    try {
        updateLastEditForRow(mockSheet, 5);
        if (valueSet instanceof Date) assertions++; else console.error(`❌ ${testName}: FAILED - Timestamp was not a Date.`);
        if (formulaSet === _generateOptimizedRelativeTimeFormula("Z5")) assertions++; else console.error(`❌ ${testName}: FAILED - Formula not set correctly.`);
        if (assertions === 2) console.log(`✅ ${testName}: PASSED`);
        else throw new Error(`${testName}: One or more assertions failed.`);
    } finally {
        this.ensureLastEditColumns = originalEnsure;
    }
}

// =================================================================
// =================== LoggerService.gs Tests ======================
// =================================================================
function testNotifyError() {
    const testName = "testNotifyError";
    let assertions = 0;
    const mocks = { MailApp: this.MailApp, PropertiesService: this.PropertiesService, SpreadsheetApp: this.SpreadsheetApp, CONFIG: this.CONFIG };
    let emailSent = null;
    try {
        this.MailApp = { sendEmail: (options) => { emailSent = options; } };
        this.PropertiesService = { getScriptProperties: () => ({ getProperty: () => "test@example.com" }) };
        this.SpreadsheetApp = { getActiveSpreadsheet: () => ({ getName: () => "Test Sheet", getId: () => "id", getUrl: () => "#" }) };
        this.CONFIG = { APP_NAME: "Test App", LOGGING: { ERROR_EMAIL_PROP: "p" } };
        const error = new Error("Test error"); error.stack = "Test stack";
        notifyError("Test Subject", error);
        if (emailSent && emailSent.to === "test@example.com") assertions++; else console.error(`❌ ${testName}: Incorrect recipient.`);
        if (emailSent && emailSent.htmlBody.includes("Test error")) assertions++; else console.error(`❌ ${testName}: Body missing error message.`);
        if (assertions === 2) console.log(`✅ ${testName}: PASSED`); else throw new Error(`${testName}: Assertions failed.`);
    } finally {
        this.MailApp = mocks.MailApp; this.PropertiesService = mocks.PropertiesService; this.SpreadsheetApp = mocks.SpreadsheetApp; this.CONFIG = mocks.CONFIG;
    }
}

function testLogAudit() {
    const testName = "testLogAudit";
    const mocks = { SpreadsheetApp: this.SpreadsheetApp, Session: this.Session, PropertiesService: this.PropertiesService, getMonthKeyPadded: this.getMonthKeyPadded, CONFIG: this.CONFIG, notifyError: this.notifyError };
    let appendedRow = null;
    try {
        const mockSheet = { appendRow: (r) => { appendedRow = r; }, getLastRow: () => 2, getLastColumn: () => 1, getRange: () => ({ sort: () => {} }) };
        this.SpreadsheetApp = { openById: () => ({ getSheetByName: () => mockSheet, insertSheet: () => mockSheet }), getActiveSpreadsheet: () => ({}) };
        this.Session = { getActiveUser: () => ({ getEmail: () => "user@test.com" }) };
        this.PropertiesService = { getScriptProperties: () => ({ getProperty: () => "id" }) };
        this.getMonthKeyPadded = () => "2024-01";
        this.CONFIG = { LOGGING: { SPREADSHEET_ID_PROP: 'id' } };
        this.notifyError = () => {};
        logAudit({ getName: () => "Source", getId: () => "s_id" }, { action: "Test Action" });
        if (appendedRow && appendedRow[1] === "user@test.com") console.log(`✅ ${testName}: PASSED`);
        else throw new Error(`${testName}: Audit row not appended correctly.`);
    } finally {
        this.SpreadsheetApp = mocks.SpreadsheetApp; this.Session = mocks.Session; this.PropertiesService = mocks.PropertiesService; this.getMonthKeyPadded = mocks.getMonthKeyPadded; this.CONFIG = mocks.CONFIG; this.notifyError = mocks.notifyError;
    }
}

// =================================================================
// ============ TransferEngine.gs Integration Test =================
// =================================================================
function testExecuteTransfer_Success() {
  const testName = "testExecuteTransfer_Success";
  const mocks = { CONFIG: this.CONFIG, logAudit: this.logAudit, notifyError: this.notifyError, updateLastEditForRow: this.updateLastEditForRow, LockService: this.LockService };
  let appendedData = null;
  try {
      const mockDestSheet = { _data: [["Project Name"]], getName: () => "Destination", getLastRow: function() { return this._data.length; }, getLastColumn: () => 1, appendRow: function(r) { this._data.push(r); appendedData = r; }, getRange: function(r,c,nr,nc) { return { getValues: () => [[this._data[r-1][c-1]]] }; } };
      const mockSourceSheet = { _data: [[""],["Test Proj"]], getName: () => "Source", getLastColumn: () => 1, getRange: function(r,c,nr,nc) { return { getValues: () => [[this._data[r-1][c-1]]] }; } };
      const mockEvent = { range: { getRow: () => 2, getSheet: () => mockSourceSheet }, source: { getSheetByName: (n) => mockDestSheet } };
      const transferConfig = { transferName: "Test", destinationSheetName: "Destination", destinationColumnMapping: { 1: 1 }, duplicateCheckConfig: { projectNameSourceCol: 1, projectNameDestCol: 1 } };
      this.CONFIG = { LAST_EDIT: { TRACKED_SHEETS: ["Destination"] } };
      this.logAudit = () => {}; this.notifyError = () => {}; this.updateLastEditForRow = () => {}; this.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
      executeTransfer(mockEvent, transferConfig);
      if (appendedData && appendedData[0] === "Test Proj") console.log(`✅ ${testName}: PASSED`);
      else throw new Error(`${testName}: Failed. Appended data: ${JSON.stringify(appendedData)}`);
  } finally {
      this.CONFIG = mocks.CONFIG; this.logAudit = mocks.logAudit; this.notifyError = mocks.notifyError; this.updateLastEditForRow = mocks.updateLastEditForRow; this.LockService = mocks.LockService;
  }
}

// =================================================================
// ============ TransferEngine.gs & Utilities.gs Tests =============
// =================================================================

function testIsDuplicateInDestination_DateInProjectName() {
  const testName = "testIsDuplicateInDestination_DateInProjectName";
  const mockSheet = {
    _data: [["Header"], [new Date("2024-09-29T00:00:00.000Z"), "Some Value", new Date("2025-01-01T00:00:00.000Z")]],
    getLastRow: function() { return this._data.length; }, getLastColumn: function() { return this._data[0].length; },
    getRange: function(r, c, nr, nc) { return { getValues: () => this._data.slice(r - 1).map(row => row.slice(c-1, c-1+nc)) }; }
  };
  const sourceRowData = ["", "2024-09-29", "", "", new Date("2025-01-01T00:00:00.000Z")];
  const dupConfig = { projectNameDestCol: 1, compoundKeySourceCols: [5], compoundKeyDestCols: [3] };
  if (isDuplicateInDestination(mockSheet, null, "2024-09-29", sourceRowData, 5, dupConfig)) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED. Expected true.`);
}

function testIsDuplicateInDestination_TimezoneBug() {
  const testName = "testIsDuplicateInDestination_TimezoneBug";
  const mockSheet = {
    _data: [["Header"], [new Date("2024-10-31T02:00:00.000Z"), "Some Value"]],
    getLastRow: function() { return this._data.length; }, getLastColumn: function() { return this._data[0].length; },
    getRange: function(r, c, nr, nc) { return { getValues: () => this._data.slice(r - 1).map(row => row.slice(c-1, c-1+nc)) }; }
  };
  if (isDuplicateInDestination(mockSheet, null, "2024-10-31", ["2024-10-31"], 1, { projectNameDestCol: 1 })) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED. Expected true.`);
}

function testFindRowByProjectName_DateInProjectName() {
  const testName = "testFindRowByProjectName_DateInProjectName";
  const mockSheet = {
    _data: [["Project Name"], [new Date("2024-10-31T00:00:00.000Z")]],
    getLastRow: function() { return this._data.length; },
    getRange: function(r, c, nr, nc) { return { getValues: () => this._data.slice(r-1).map(row => row.slice(c-1, c-1+nc)), createTextFinder: () => ({ matchCase:()=>({matchEntireCell:()=>({findNext:()=>null})})}) }; }
  };
  if (findRowByProjectNameRobust(mockSheet, "2024-10-31", 1) === 2) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED. Expected row 2.`);
}

function testFindRowByValue_WhitespaceBug() {
  const testName = "testFindRowByValue_WhitespaceBug";
  const mockSheet = {
    _data: [["Header"], ["  12345  "]],
    getLastRow: function() { return this._data.length; },
    getRange: function(r, c, nr, nc) { return { getValues: () => this._data.slice(r - 1).map(row => row.slice(c-1, c-1+nc)) }; }
  };
  if (findRowByValue(mockSheet, "12345", 1) === 2) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED. Expected row 2.`);
}

function testNormalizeForComparison() {
  const testName = "testNormalizeForComparison";
  const testData = [
    { i: null, e: "" }, { i: undefined, e: "" }, { i: "  hello  ", e: "hello" }, { i: 123, e: "123" },
    { i: true, e: "true" }, { i: false, e: "false" }, { i: new Date("2024-01-01T12:00:00Z"), e: "2024-01-01T12:00:00" }
  ];
  const fails = testData.filter(t => normalizeForComparison(t.i) !== t.e);
  if (fails.length === 0) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED on inputs: ${JSON.stringify(fails.map(f=>f.i))}`);
}

function testNormalizeString() {
  const testName = "testNormalizeString";
  const testData = [{ i: "  Hello World  ", e: "hello world" }, { i: "ALL CAPS", e: "all caps" }, { i: null, e: "" }];
  const fails = testData.filter(t => normalizeString(t.i) !== t.e);
  if (fails.length === 0) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED on inputs: ${JSON.stringify(fails.map(f=>f.i))}`);
}

function testIsTrueLike() {
  const testName = "testIsTrueLike";
  const tVals = ["true", "TRUE", "Y", "1", true], fVals = ["false", "no", "0", null, false, ""];
  const tFails = tVals.filter(v => isTrueLike(v) !== true);
  const fFails = fVals.filter(v => isTrueLike(v) !== false);
  if (tFails.length === 0 && fFails.length === 0) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED on inputs: ${JSON.stringify(tFails.concat(fFails))}`);
}

function testParseAndNormalizeDate() {
    const testName = "testParseAndNormalizeDate";
    const d = new Date("2024-05-20T15:30:00Z");
    const expected = new Date("2024-05-20T00:00:00Z").getTime();
    const result = parseAndNormalizeDate(d);
    if (result && result.getTime() === expected) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testFormatValueForKey() {
    const testName = "testFormatValueForKey";
    const testData = [{ i: new Date("2024-07-15T10:00:00Z"), e: "2024-07-15" }, { i: " Some String ", e: "some string" }];
    const fails = testData.filter(t => formatValueForKey(t.i) !== t.e);
    if (fails.length === 0) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED on inputs: ${JSON.stringify(fails.map(f=>f.i))}`);
}

function testGetMonthKeyPadded() {
    const testName = "testGetMonthKeyPadded";
    if (getMonthKeyPadded(new Date(2024, 6, 15)) === "2024-07") console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testGetHeaderColumnIndex() {
    const testName = "testGetHeaderColumnIndex";
    const mockSheet = { getLastColumn: () => 3, getRange: () => ({ getValues: () => [["Name", "  Email  ", "ID"]] }) };
    if (getHeaderColumnIndex(mockSheet, "Email") === 2 && getHeaderColumnIndex(mockSheet, "id") === 3) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testFindRowByBestIdentifier() {
    const testName = "testFindRowByBestIdentifier";
    const mockSheet = {
        _data: [["SFID"],["SFID123"]],
        getLastRow: function() { return this._data.length; },
        getRange: function(r, c) { return { getValues: () => [[this._data[r-1][c-1]]] }; }
    };
    if (findRowByBestIdentifier(mockSheet, "SFID123", 1, "Project A", 2) === 2) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED. SFID lookup failed.`);
}

function testCreateMapping() {
    const testName = "testCreateMapping";
    if (JSON.stringify(createMapping([[1, 5], [2, 6]])) === JSON.stringify({ 1: 5, 2: 6 })) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testGetMaxValueInObject() {
    const testName = "testGetMaxValueInObject";
    if (getMaxValueInObject({ a: 10, b: 50, c: 20 }) === 50) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testUniqueArray() {
    const testName = "testUniqueArray";
    if (JSON.stringify(uniqueArray([1, 2, "a", 2, "b", "a", 3])) === JSON.stringify([1, 2, "a", "b", 3])) console.log(`✅ ${testName}: PASSED`);
    else throw new Error(`${testName}: FAILED`);
}

function testIsDuplicateInDestination_CompoundKey() {
  const testName = "testIsDuplicateInDestination_CompoundKey";
  const mockSheet = {
    _data: [["Header"],["Project X", new Date("2024-10-31T00:00:00.000Z"), "user@test.com"]],
    getLastRow: function() { return this._data.length; }, getLastColumn: () => 3,
    getRange: function(r, c, nr, nc) { return { getValues: () => this._data.slice(r-1).map(row => row.slice(c-1, c-1+nc)) }; }
  };
  const sourceRowData = ["", "Project X", new Date("2024-10-31"), "user@test.com"];
  const dupConfig = { projectNameDestCol: 1, compoundKeySourceCols: [3, 4], compoundKeyDestCols: [2, 3] };
  if (isDuplicateInDestination(mockSheet, null, "Project X", sourceRowData, 4, dupConfig)) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED`);
}

function testIsDuplicateInDestination_NoDuplicate() {
  const testName = "testIsDuplicateInDestination_NoDuplicate";
  const mockSheet = { _data: [["Header"],["Existing Project"]], getLastRow: () => 2, getLastColumn: () => 1, getRange: () => ({ getValues: () => [["Existing Project"]] }) };
  if (!isDuplicateInDestination(mockSheet, null, "New Project", ["New Project"], 1, { projectNameDestCol: 1 })) console.log(`✅ ${testName}: PASSED`);
  else throw new Error(`${testName}: FAILED`);
}