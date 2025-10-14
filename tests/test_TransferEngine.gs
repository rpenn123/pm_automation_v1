/**
 * Test suite for TransferEngine.gs
 */

function runTransferEngineTests() {
  test_executeTransfer_handlesTrailingEmptyColumns();
}

/**
 * Tests the `executeTransfer` function to ensure it correctly sizes the new row
 * when the destination sheet has trailing empty columns.
 *
 * The bug occurs when the destination column mapping requires writing to a column
 * that is beyond the last column with data (`getLastColumn()`) but within the
 * sheet's full width (`getMaxColumns()`). The original code sized the new row
 * using `getLastColumn()`, which could truncate the row and cause data to be lost or
 * written to the wrong column.
 *
 * The fix is to use `getMaxColumns()` when calculating the new row's width, ensuring
 * it always matches the full sheet width.
 */
function test_executeTransfer_handlesTrailingEmptyColumns() {
  const testName = 'test_executeTransfer_handlesTrailingEmptyColumns';
  console.log('Running test: ' + testName);
  const original = JSON.parse(JSON.stringify(global.CONFIG));

  try {
    // 1. ARRANGE

    // Mock global functions needed by TransferEngine
    const utilsCode = fs.readFileSync('src/core/Utilities.gs', 'utf8');
    eval(utilsCode);
    global.logAudit = () => {}; // Mocked logger
    global.notifyError = (msg, err) => { throw err || new Error(msg); }; // Mocked notifier
    global.updateLastEditForRow = () => {}; // Mocked helper
    global.CONFIG = { SHEETS: {}, LAST_EDIT: { TRACKED_SHEETS: [] } }; // Minimal config


    // Mock destination sheet with trailing empty columns
    const mockDestSheet = {
      _appendedRow: null,
      getMaxColumns: () => 5,    // The sheet actually has 5 columns.
      getLastColumn: () => 2,   // But data only exists up to column 2. BUG SOURCE.
      getLastRow: () => 1,      // No data rows yet.
      appendRow: function(rowArray) {
        this._appendedRow = rowArray; // Capture the appended row for inspection.
      },
      getRange: () => ({ getValues: () => [[]] }), // For duplicate check
      getName: () => 'DestinationSheet'
    };

    // Spy on the appendRow method to check what's being passed to it.
    const appendRowSpy = jest.spyOn(mockDestSheet, 'appendRow');

    // Mock source data and event object
    const sourceRowData = ["Project Y", "Important Data"];
    const mockEvent = {
      source: {
        getSheetByName: (name) => (name === 'DestinationSheet' ? mockDestSheet : null)
      },
      range: {
        getSheet: () => ({ getName: () => 'SourceSheet', getLastColumn: () => 2 }), // Add mock for getLastColumn
        getRow: () => 2
      }
    };

    // Transfer configuration that maps to a trailing empty column (col 5)
    const transferConfig = {
      transferName: "Test Transfer",
      destinationSheetName: "DestinationSheet",
      destinationColumnMapping: {
        1: 1, // Project Name -> Project Name
        2: 5  // Data -> Mapped to column 5, which is in the "empty" zone
      },
      // FINAL FIX: Provide a complete, albeit disabled, configuration to ensure
      // the function does not exit prematurely due to missing properties.
      duplicateCheckConfig: {
        checkEnabled: false,
        sfidSourceCol: null,
        sfidDestCol: null,
        projectNameSourceCol: 1,
        projectNameDestCol: 1,
        compoundKeySourceCols: [],
        compoundKeyDestCols: []
      }
    };

    const expectedRowWidth = mockDestSheet.getMaxColumns(); // Should be 5

    // 2. ACT
    // We need to load TransferEngine locally to use the modified version.
    let transferEngineCode = fs.readFileSync('src/core/TransferEngine.gs', 'utf8');
    transferEngineCode = transferEngineCode.replace('function executeTransfer(', 'global.executeTransfer = function executeTransfer(');
    transferEngineCode = transferEngineCode.replace('function isDuplicateInDestination(', 'global.isDuplicateInDestination = function isDuplicateInDestination(');
    eval(transferEngineCode);

    executeTransfer(mockEvent, transferConfig, sourceRowData);


    // 3. ASSERT
    const calls = appendRowSpy.mock.calls;
    if (calls.length === 0) {
      throw new Error(`${testName} FAILED: appendRow was never called. The transfer function likely exited early due to a configuration issue.`);
    }

    const appendedRowArray = calls[0][0];
    const actualRowWidth = appendedRowArray.length;

    if (actualRowWidth !== expectedRowWidth) {
      throw new Error(`${testName} FAILED: Expected appended row to have a width of ${expectedRowWidth} (based on getMaxColumns), but it had a width of ${actualRowWidth}. This indicates the row was truncated.`);
    }

    if (appendedRowArray[4] !== "Important Data") {
        throw new Error(`${testName} FAILED: Data was not placed in the correct column. Expected 'Important Data' at index 4, but found '${appendedRowArray[4]}'.`);
    }

    console.log(testName + ' PASSED');
  } finally {
      global.CONFIG = original;
  }
}
