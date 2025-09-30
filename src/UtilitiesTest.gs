/**
 * @OnlyCurrentDoc
 * UtilitiesTest.gs
 * This file contains unit tests for the functions in Utilities.gs.
 */

/**
 * Tests the `findRowByValue` function to ensure it correctly finds a row
 * even when the value in the sheet has leading or trailing whitespace.
 *
 * **Bug Context:** The `findRowByValue` function did not trim whitespace from
 * sheet values before comparison. This test ensures that a search for "12345"
 * correctly matches a cell containing "  12345  ".
 */
function testFindRowByValue_WhitespaceBug() {
  const testName = "testFindRowByValue_WhitespaceBug";
  let assertions = 0;

  // 1. Mock the sheet object with whitespace in the target value
  const mockSheet = {
    _data: [
      ["Header"],
      ["Value1"],
      ["  12345  "], // The value with whitespace
      ["Value3"]
    ],
    getLastRow: function() { return this._data.length; },
    getRange: function(row, col, numRows, numCols) {
      const self = this;
      return {
        getValues: function() {
          const result = [];
          for (let i = 0; i < numRows; i++) {
            const rowData = self._data[row + i - 1];
            if (rowData) {
              result.push(rowData.slice(col - 1, col - 1 + numCols));
            }
          }
          return result;
        }
      };
    }
  };

  // 2. Define the search value (trimmed) and the column
  const searchValue = "12345";
  const searchColumn = 1;
  const expectedRow = 3; // 1-based index

  // 3. Execute the function under test
  const foundRow = findRowByValue(mockSheet, searchValue, searchColumn);

  // 4. Assert the result
  if (foundRow === expectedRow) {
    console.log(`✅ ${testName}: PASSED - Correctly found row ${foundRow}.`);
    assertions++;
  } else {
    console.error(`❌ ${testName}: FAILED - Expected to find row ${expectedRow}, but got ${foundRow}.`);
  }

  if (assertions !== 1) {
    throw new Error(`${testName}: Assertion failed.`);
  }
}