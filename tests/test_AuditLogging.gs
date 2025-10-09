/**
 * Test suite for verifying audit logging functionality.
 */

// A simple spy to track function calls
function createSpy(obj, funcName) {
    const originalFunc = obj[funcName];
    const spy = {
        called: false,
        args: [],
        callCount: 0,
        reset: () => {
            spy.called = false;
            spy.args = [];
            spy.callCount = 0;
        },
        restore: () => {
            obj[funcName] = originalFunc;
        }
    };

    obj[funcName] = (...args) => {
        spy.called = true;
        spy.args.push(args);
        spy.callCount++;
        return originalFunc ? originalFunc.apply(obj, args) : undefined;
    };

    return spy;
}


// Mock for global logAudit function
let logAuditSpy;

// Mock implementations for Google Apps Script services
const mockSpreadsheet = {
    getSheetByName: (name) => mockSheet,
    getName: () => 'Test Spreadsheet',
    getId: () => 'test-ss-id',
    getUrl: () => 'http://example.com/ss'
};

const mockSheet = {
    getName: () => 'Forecasting',
    getRange: (row, col, numRows, numCols) => ({
        getValues: () => [['']],
        getRow: () => 2,
        getColumn: () => 2,
        getSheet: () => mockSheet,
        getValue: () => 'new value',
        getA1Notation: () => 'B2',
        // Add methods needed by LastEditService
        setValue: () => {},
        setFormula: () => {}
    }),
    getMaxColumns: () => 10,
    hideColumns: () => {},
    insertColumnAfter: () => {},
    getLastColumn: () => 1,
    // Add getParent to link back to the spreadsheet
    getParent: () => mockSpreadsheet
};

const mockRange = {
    getRow: () => 2,
    getColumn: () => 2,
    getNumRows: () => 1,
    getNumColumns: () => 1,
    getSheet: () => mockSheet,
    getValue: () => 'new value',
};

// Global mocks
global.SpreadsheetApp = {
    getActiveSpreadsheet: () => mockSpreadsheet,
    // **Bug Fix**: Add the `create` mock to prevent test failures.
    // This was overriding the mock from run_test.js.
    create: (name) => {
        return {
            getId: () => 'mock-created-ss-id',
            getName: () => name,
            getSheets: () => [],
            getSheetByName: () => null,
            insertSheet: (sheetName) => ({
                getName: () => sheetName,
                getRange: () => ({
                    setValues: () => ({
                        setFontWeight: () => {}
                    })
                }),
                setFrozenRows: () => {},
                appendRow: () => {} // **Bug Fix**: Correctly placed on the sheet mock
            })
        };
    }
};

global.Session = {
    getActiveUser: () => ({
        getEmail: () => 'test@example.com'
    }),
};

global.LockService = {
    getScriptLock: () => ({
        tryLock: () => true,
        releaseLock: () => {}
    }),
};

// Mock CONFIG object
const CONFIG = {
    APP_NAME: "Test App",
    LAST_EDIT: {
        TRACKED_SHEETS: ['Forecasting', 'Upcoming'],
        AT_HEADER: "AT_HEADER",
        REL_HEADER: "REL_HEADER"
    },
    SHEETS: {
        FORECASTING: 'Forecasting',
        UPCOMING: 'Upcoming'
    },
    FORECASTING_COLS: {
        PROGRESS: 2,
    },
    UPCOMING_COLS: {
        PROGRESS: 2
    },
    STATUS_STRINGS: {
        IN_PROGRESS: "In Progress"
    }
};

/**
 * Test runner for audit logging.
 */
function runAuditLoggingTests() {
    console.log('Running tests for Audit Logging...');
    testOnEdit_ShouldLogAuditForTrackedSheets();
    console.log('All Audit Logging tests passed!');
}

/**
 * Test case: Verifies that a simple edit on a tracked sheet triggers an audit log entry.
 */
function testOnEdit_ShouldLogAuditForTrackedSheets() {
    console.log('Running test: testOnEdit_ShouldLogAuditForTrackedSheets');

    // Set up the spy on the global logAudit function
    logAuditSpy = createSpy(global, 'logAudit');

    // Mock the onEdit event object
    const e = {
        range: mockRange,
        value: 'new value',
        oldValue: 'old value',
        source: mockSpreadsheet
    };

    try {
        // Execute the onEdit function
        onEdit(e);

        // Assertions
        if (!logAuditSpy.called) {
            throw new Error('Test Failed: logAudit was not called.');
        }

        if (logAuditSpy.callCount !== 1) {
            throw new Error(`Test Failed: Expected logAudit to be called once, but it was called ${logAuditSpy.callCount} times.`);
        }

        const logArgs = logAuditSpy.args[0];
        const logEntry = logArgs[1]; // The 'entry' object is the second argument

        if (logEntry.action !== 'Row Edit') {
            throw new Error(`Test Failed: Expected action to be 'Row Edit', but got '${logEntry.action}'.`);
        }

        if (logEntry.sourceSheet !== 'Forecasting') {
            throw new Error(`Test Failed: Expected sourceSheet to be 'Forecasting', but got '${logEntry.sourceSheet}'.`);
        }

        if (logEntry.sourceRow !== 2) {
            throw new Error(`Test Failed: Expected sourceRow to be 2, but got '${logEntry.sourceRow}'.`);
        }

        if (!logEntry.details.includes('New value: "new value"')) {
            throw new Error(`Test Failed: Log details did not contain the expected new value. Got: '${logEntry.details}'.`);
        }

        console.log('Test passed: testOnEdit_ShouldLogAuditForTrackedSheets');
    } finally {
        // Clean up the spy
        logAuditSpy.restore();
    }
}