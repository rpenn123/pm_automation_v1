// This needs to be at the top level for the test file to use it.
const fs = require('fs');

function runTransferEngineTests() {
  console.log("Running TransferEngine Tests...");
  test_executeTransfer_appends_row_with_full_sheet_width();
  console.log("TransferEngine Tests Passed.");
}

/**
 * Test case for the bug where executeTransfer uses getLastColumn() instead of getMaxColumns(),
 * potentially causing data misalignment when appending rows to sheets with empty trailing columns.
 */
function test_executeTransfer_appends_row_with_full_sheet_width() {
    // Save original global objects to prevent test pollution
    const originalConfig = global.CONFIG;
    const originalLogAudit = global.logAudit;

    try {
        // Load the TransferEngine script directly into this scope.
        // This ensures the executeTransfer function is defined locally and will
        // use the mocks defined in this test, avoiding global scope issues.
        const transferEngineScript = fs.readFileSync('src/core/TransferEngine.gs', 'utf8');
        eval(transferEngineScript);

        // 1. Setup Mocks
        const mockAppendRow = {
            called: false,
            lastArg: null,
            call: function(arg) {
                this.called = true;
                this.lastArg = arg;
            }
        };

        const mockDestinationSheet = {
            getName: () => 'Destination Sheet', // Add getName for logging
            getLastColumn: () => 10,
            getMaxColumns: () => 20,
            getLastRow: () => 5,
            appendRow: (row) => mockAppendRow.call(row),
            getRange: (row, col, numRows, numCols) => {
                // This is the specific call we expect from findRowByValue
                if (row === 2 && col === 1 && numRows === 4 && numCols === 1) {
                    return { getValues: () => [] }; // No duplicate found
                }
                // Default mock for other calls
                return { getValues: () => [[]] };
            },
        };

        const mockSourceSheet = {
            getName: () => 'Source Sheet',
            getRange: () => ({
                getValues: () => [[ 'SFID-123', 'Project A', 'Value C' ]]
            }),
            getLastColumn: () => 3,
        };

        const mockSpreadsheet = {
            getSheetByName: (name) => (name === 'Destination Sheet' ? mockDestinationSheet : null)
        };

        global.logAudit = () => {}; // Simple mock to prevent errors

        // 2. Setup Test Data and Configuration
        const e = {
            range: { getSheet: () => mockSourceSheet, getRow: () => 2 },
            source: mockSpreadsheet,
        };

        const transferConfig = {
            transferName: "Test Transfer",
            destinationSheetName: "Destination Sheet",
            destinationColumnMapping: { 1: 1, 2: 2, 3: 5 },
            duplicateCheckConfig: {
                checkEnabled: true,
                sfidSourceCol: 1, sfidDestCol: 1,
                projectNameSourceCol: 2, projectNameDestCol: 2,
            },
            postTransferActions: { sort: false }
        };

        global.CONFIG = {
            PROPERTIES: { ERROR_EMAIL_PROP: 'error-email-key' },
            LAST_EDIT: { TRACKED_SHEETS: [] }
        };

        // 3. Execute the function under test
        executeTransfer(e, transferConfig, null);

        // 4. Assertions
        if (!mockAppendRow.called) {
            throw new Error("Assertion failed: destinationSheet.appendRow() was not called.");
        }

        const appendedRow = mockAppendRow.lastArg;
        const expectedWidth = mockDestinationSheet.getMaxColumns();
        const actualWidth = appendedRow.length;

        if (actualWidth !== expectedWidth) {
            throw new Error(`Assertion failed: Appended row width is incorrect. Expected: ${expectedWidth}, Actual: ${actualWidth}`);
        }

        console.log("test_executeTransfer_appends_row_with_full_sheet_width: PASSED");

    } finally {
        // Restore globals
        global.CONFIG = originalConfig;
        global.logAudit = originalLogAudit;
    }
}