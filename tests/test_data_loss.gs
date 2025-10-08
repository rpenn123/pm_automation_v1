/**
 * Test case to verify the bug where editing a row with empty cells
 * causes data loss during a transfer operation.
 */
function runDataLossTest() {
    console.log("Running Data Loss Test...");

    // 1. Mock the necessary global objects and functions.
    let capturedSourceRowData = null;
    global.executeTransfer = (e, config, sourceRowData) => {
        capturedSourceRowData = sourceRowData;
    };
    global.updateLastEditForRow = () => {}; // Mocked for this test
    global.logAudit = () => {}; // Mocked for this test
    global.notifyError = () => {}; // Mocked for this test
    global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
    global.acquireLockWithRetry = () => true;
    // Mock the sync function to prevent it from erroring during this test.
    global.syncProgressToUpcoming = () => {};


    const MOCK_SHEET_DATA = [
        // Col 1 (A) to 6 (F) have some data.
        "SFID-123", "Project X", "Detail", "", "Equipment", "Some Progress",
        // Col 7 (G) is the edited column. Col 8 (H) and 9 (I) are empty.
        "In Progress", "", "",
        // Col 10 (J) is the 'Deadline' column, which is required for the transfer.
        "2025-12-31"
    ];

    const mockSpreadsheet = {
        getSheetByName: (name) => mockSheet,
        getName: () => 'Test Spreadsheet',
        getId: () => 'test-ss-id',
        getUrl: () => 'http://example.com/ss'
    };

    const mockSheet = {
        getName: () => "Forecasting",
        // This is the source of the bug. It returns the last column with content.
        getLastColumn: () => 7,
        // This is the correct method to use. It returns the total number of columns.
        getMaxColumns: () => 10,
        // Add mock for getLastRow to prevent `findRowByValue` from crashing
        getLastRow: () => 2,
        getRange: (row, col, numRows, numCols) => ({
            getValues: () => [MOCK_SHEET_DATA.slice(0, numCols)],
            getSheet: () => mockSheet,
            getRow: () => 2,
            getColumn: () => 7,
            setValue: () => {},
            setFormula: () => {},
            getA1Notation: () => 'A1',
            getValue: () => 'some old value' // Add mock to prevent crash in sync function
        }),
        // Add mocks to allow LastEditService to run without crashing
        insertColumnAfter: () => {},
        hideColumns: () => {},
        getParent: () => mockSpreadsheet,
    };

    const mockEvent = {
        range: {
            getSheet: () => mockSheet,
            getRow: () => 2,
            getColumn: () => 7, // User edits the 'Progress' column
            getNumRows: () => 1,
            getNumColumns: () => 1,
        },
        value: "In Progress",
        oldValue: "Some Progress",
        source: mockSpreadsheet,
    };

    // 2. Call the onEdit function (already loaded globally) to trigger the automation.
    onEdit(mockEvent);

    // 4. Assert the result. This test should FAIL before the fix is applied.
    if (!capturedSourceRowData) {
        throw new Error("Test failed: executeTransfer was not called.");
    }

    const expectedDeadline = "2025-12-31";
    const actualDeadline = capturedSourceRowData[CONFIG.FORECASTING_COLS.DEADLINE - 1];

    if (actualDeadline !== expectedDeadline) {
        throw new Error(`Test FAILED: Deadline data was lost. Expected '${expectedDeadline}', but got '${actualDeadline}'. This confirms the bug.`);
    }

    console.log("Test PASSED: Deadline data was correctly transferred.");
}