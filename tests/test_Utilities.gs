/**
 * Test suite for Utilities.gs
 */

// Mock SpreadsheetApp objects for testing purposes
const MockSpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: (name) => {
      if (name === 'TestSheet') {
        return MockSheet;
      }
      return null;
    },
  }),
};

const MockSheet = {
  getLastRow: () => 3,
  getRange: (row, column, numRows, numColumns) => {
    // Mock data for the sheet
    const data = [
      ['ID'], // Header row
      ['id-1'],
      ['  id-2  '], // Data with whitespace
    ];
    return {
      getValues: () => {
        return data.slice(row - 1, row - 1 + numRows);
      },
    };
  },
};


/**
 * Test runner
 */
function runTests() {
  console.log('Running tests for Utilities.gs...');
  testFindRowByValue_WhitespaceBug();
  console.log('All tests passed!');
}


/**
 * Test case for the whitespace bug in findRowByValue.
 * This test fails before the fix and passes after it.
 */
function testFindRowByValue_WhitespaceBug() {
  console.log('Running test: testFindRowByValue_WhitespaceBug');

  // --- Test Case 1: Search value with leading/trailing whitespace ---
  let searchValue = '  id-2  ';
  let expectedRow = 3;
  let actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 1 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  // --- Test Case 2: Search value without whitespace ---
  searchValue = 'id-1';
  expectedRow = 2;
  actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 2 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  // --- Test Case 3: Value not found ---
  searchValue = 'id-3';
  expectedRow = -1;
  actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 3 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  console.log('Test passed: testFindRowByValue_WhitespaceBug');
}