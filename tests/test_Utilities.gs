/**
 * Test suite for Utilities.gs
 */

function runUtilityTests() {
  test_getHeaderColumnIndex_withEmptyTrailingColumns();
}

/**
 * Tests the `getHeaderColumnIndex` function to ensure it can find a header
 * even when there are empty columns after the last populated cell in the header row.
 * This scenario causes `sheet.getLastColumn()` to return a value smaller than
 * the actual column count, which can cause the function to miss headers.
 * The fix is to use `sheet.getMaxColumns()` instead.
 */
function test_getHeaderColumnIndex_withEmptyTrailingColumns() {
  const testName = 'test_getHeaderColumnIndex_withEmptyTrailingColumns';

  // 1. ARRANGE
  // Mock Sheet object to simulate a sheet where the last column with data
  // is not the last column with a header.
  const mockSheet = {
    // Simulate a sheet with 5 total columns.
    getMaxColumns: function() {
      return 5;
    },
    // Simulate that the last column with any content in the entire sheet is column 2.
    // This is the source of the bug.
    getLastColumn: function() {
      return 2;
    },
    // The getRange function will be called by getHeaderColumnIndex. We can inspect
    // its arguments to see if the buggy or fixed version is being used.
    getRange: function(row, col, numRows, numCols) {
      // The full header row we want to be able to search.
      const allHeaders = ["Project Name", "Date", "Status", "", ""];
      // Slice the headers based on what the function is asking for.
      const requestedHeaders = allHeaders.slice(col - 1, col - 1 + numCols);

      return {
        getValues: function() {
          return [requestedHeaders];
        }
      };
    }
  };

  const headerToFind = "Status";
  const expectedColumn = 3; // "Status" is the 3rd column (1-indexed).

  // 2. ACT
  // Call the function under test.
  const result = getHeaderColumnIndex(mockSheet, headerToFind);

  // 3. ASSERT
  // With the bug, the function uses getLastColumn() (2), reads only the first 2 headers,
  // fails to find "Status", and returns -1.
  // The fix uses getMaxColumns() (5), reads all 5 headers, finds "Status", and returns 3.
  if (result !== expectedColumn) {
    throw new Error(testName + ' FAILED: Expected to find header "' + headerToFind + '" at column ' + expectedColumn + ', but got ' + result + '. This indicates the function is likely using getLastColumn() instead of getMaxColumns().');
  } else {
    Logger.log(testName + ' PASSED');
  }
}