/**
 * Test suite for TransferEngine.gs sorting logic.
 */

function runTransferEngineSortTests() {
  test_sortingOnFirstDataRowShouldNotThrowError();
}

/**
 * This test verifies a fix for a bug where the post-transfer sort operation
 * would throw an error when adding the very first data row to a sheet.
 * The bug triggered a "sorting failed" notification, which was confusing for the user.
 * The test is designed to FAIL before the fix and PASS after the fix.
 */
function test_sortingOnFirstDataRowShouldNotThrowError() {
  const testName = 'test_sortingOnFirstDataRowShouldNotThrowError';
  console.log('Running test: ' + testName);
  const original = JSON.parse(JSON.stringify(global.CONFIG));
  const notifyErrorSpy = jest.spyOn(global, 'notifyError').mockImplementation(() => {});

  try {
    // 1. ARRANGE
    // Load dependencies
    const utilsCode = fs.readFileSync('src/core/Utilities.gs', 'utf8');
    eval(utilsCode);

    // Mock global objects and functions
    global.logAudit = () => {};
    global.updateLastEditForRow = () => {};
    global.CONFIG = {
        ...original,
        LAST_EDIT: { TRACKED_SHEETS: ['DestinationSheet'] },
        PROPS: { ERROR_EMAIL_PROP: 'error_email_prop' },
        SETTINGS: { DEFAULT_NOTIFICATION_EMAIL: 'test@example.com' }
    };

    // Mock the behavior of range.sort() which throws on a 1-row range.
    const mockSortRange = {
      sort: jest.fn(() => {
        throw new Error("The number of rows in the range must be at least 2.");
      })
    };

    // Mock the destination sheet to simulate adding the first data row (row index 2)
    const mockDestSheet = {
      _lastRow: 1, // Start with only a header row
      getMaxColumns: () => 5,
      getLastRow: function() { return this._lastRow; },
      appendRow: function(rowArray) {
        this._lastRow++; // After append, lastRow is 2
      },
      // This mock is crucial. It returns the sortable range that will cause the error.
      getRange: jest.fn((row, col, numRows, numCols) => {
        // The buggy code calls getRange(2, 1, 1, ...) when appendedRow is 2.
        if (row === 2 && numRows === 1) {
          return mockSortRange;
        }
        // Fallback for other getRange calls (e.g., duplicate check)
        return { getValues: () => [[]] };
      }),
      getName: () => 'DestinationSheet'
    };

    // Mock the event object
    const mockEvent = {
      source: {
        getSheetByName: (name) => (name === 'DestinationSheet' ? mockDestSheet : null),
        flush: () => {}
      },
      range: {
        getSheet: () => ({ getName: () => 'SourceSheet', getMaxColumns: () => 2 }),
        getRow: () => 2
      }
    };

    // Mock SpreadsheetApp, which is called for `flush()`
    global.SpreadsheetApp = {
      flush: () => mockEvent.source.flush()
    };

    // Transfer config with sorting enabled
    const transferConfig = {
      transferName: "Sort Test Transfer",
      destinationSheetName: "DestinationSheet",
      destinationColumnMapping: { 1: 1, 2: 2 },
      duplicateCheckConfig: {
          checkEnabled: false,
          // This is the key fix for the test itself. It was missing,
          // causing the transfer to skip before hitting the sort logic.
          projectNameSourceCol: 1
      },
      postTransferActions: {
        sort: true,
        sortColumn: 1,
        sortAscending: true
      }
    };

    // Load TransferEngine for local execution
    let transferEngineCode = fs.readFileSync('src/core/TransferEngine.gs', 'utf8');
    transferEngineCode = transferEngineCode.replace('function executeTransfer(', 'global.executeTransfer = function executeTransfer(');
    transferEngineCode = transferEngineCode.replace('function isDuplicateInDestination(', 'global.isDuplicateInDestination = function isDuplicateInDestination(');
    eval(transferEngineCode);

    // 2. ACT
    executeTransfer(mockEvent, transferConfig, ['Project X', 'Some Data']);


    // 3. ASSERT
    const errorCalls = notifyErrorSpy.mock.calls;
    if (errorCalls.length > 0) {
      throw new Error(`${testName} FAILED: notifyError was called unexpectedly. Message: "${errorCalls[0][0]}"`);
    }

    console.log(testName + ' PASSED');

  } finally {
      // Clean up spy
      notifyErrorSpy.mockRestore();
      global.CONFIG = original;
  }
}
