/**
 * Test suite for the findRowByValue function in Utilities.gs
 */

function runFindRowByValueTests() {
  test_findRowByValue_canFindFalsyValues();
  test_findRowByValue_returnsNegativeOneForNotFound();
}

/**
 * This test verifies that the `findRowByValue` function can correctly find rows
 * containing "falsy" values like `0` and `""` (empty string).
 *
 * The bug this test targets is in the initial guard clause of the function:
 * `if (!sheet || !value || !column) return -1;`
 * The `!value` check incorrectly returns -1 for `0` and `""`, preventing the
 * search from ever executing for these valid identifiers.
 *
 * This test is expected to FAIL before the fix and PASS after the fix.
 */
function test_findRowByValue_canFindFalsyValues() {
  const testName = 'test_findRowByValue_canFindFalsyValues';
  console.log(`Running test: ${testName}`);

  // 1. ARRANGE
  const mockSheet = {
    _data: [[5], [0], [""], ["hello"]], // Test data, 1-based indexing for rows starts at 1
    getLastRow: function() { return this._data.length + 1; }, // +1 for header
    getRange: function(row, col, numRows, numCols) {
      // The function requests data starting from row 2
      const requestedData = this._data.slice(row - 2, row - 2 + numRows);
      return {
        getValues: () => requestedData
      };
    }
  };

  // The function is globally available from Utilities.gs loaded by run_test.js
  const findRowByValueFunc = findRowByValue;

  // 2. ACT & 3. ASSERT
  let failed = false;
  let messages = [];

  // Test case 1: Find the number 0
  const resultForZero = findRowByValueFunc(mockSheet, 0, 1);
  const expectedForZero = 3; // 0 is in the 2nd data row, which is sheet row 3
  if (resultForZero !== expectedForZero) {
    failed = true;
    messages.push(`Expected to find '0' at row ${expectedForZero}, but got ${resultForZero}.`);
  }

  // Test case 2: Find an empty string ""
  const resultForEmptyString = findRowByValueFunc(mockSheet, "", 1);
  const expectedForEmptyString = 4; // "" is in the 3rd data row, which is sheet row 4
  if (resultForEmptyString !== expectedForEmptyString) {
    failed = true;
    messages.push(`Expected to find '""' at row ${expectedForEmptyString}, but got ${resultForEmptyString}.`);
  }

  // Test case 3: Find a normal value to ensure function still works
  const resultForHello = findRowByValueFunc(mockSheet, "hello", 1);
  const expectedForHello = 5;
  if (resultForHello !== expectedForHello) {
      failed = true;
      messages.push(`Expected to find '"hello"' at row ${expectedForHello}, but got ${resultForHello}.`);
  }

  if (failed) {
    console.error(`${testName} FAILED: ${messages.join(' ')}`);
    throw new Error(`${testName} FAILED: ${messages.join(' ')}`);
  } else {
    console.log(`${testName} PASSED`);
  }
}

/**
 * Verifies that findRowByValue returns -1 when a value is not found.
 */
function test_findRowByValue_returnsNegativeOneForNotFound() {
    const testName = 'test_findRowByValue_returnsNegativeOneForNotFound';
    console.log(`Running test: ${testName}`);

    const mockSheet = {
        _data: [["A"], ["B"]],
        getLastRow: function() { return this._data.length + 1; },
        getRange: function(row, col, numRows, numCols) {
            const requestedData = this._data.slice(row - 2, row - 2 + numRows);
            return { getValues: () => requestedData };
        }
    };

    const result = findRowByValue(mockSheet, "C", 1);

    if (result !== -1) {
        throw new Error(`${testName} FAILED: Expected -1 for a value not found, but got ${result}.`);
    }

    console.log(`${testName} PASSED`);
}