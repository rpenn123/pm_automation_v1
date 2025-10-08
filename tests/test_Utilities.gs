/**
 * Test suite for Utilities.gs
 */

// =================================================================
// ========================= MOCKS =================================
// =================================================================

// Mock implementation of Utilities.formatDate for the test environment
const Utilities = {
  formatDate: (date, timeZone, format) => {
    if (format === "yyyy-MM-dd") {
      const d = new Date(date);
      const year = d.getUTCFullYear();
      const month = String(d.getUTCMonth() + 1).padStart(2, '0');
      const day = String(d.getUTCDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
    return date.toString();
  }
};

// Mock for Session object to provide a script timezone
const Session = {
  getScriptTimeZone: () => "America/New_York",
};

const MockSheet = {
  getLastRow: () => 3,
  getRange: (row, column, numRows, numColumns) => {
    const data = [
      ['ID'],
      ['id-1'],
      ['  id-2  '],
    ];
    return {
      getValues: () => data.slice(row - 1, row - 1 + numRows),
    };
  },
};

// =================================================================
// ======================= TEST RUNNER =============================
// =================================================================

/**
 * Test runner
 */
function runUtilityTests() {
  console.log('Running tests for Utilities.gs...');
  testFindRowByValue_WhitespaceBug();
  testFormatValueForKey_DateHandling();
  console.log('All Utilities tests passed!');
}

// =================================================================
// ======================= TEST CASES ==============================
// =================================================================

/**
 * Test case for the date format handling in formatValueForKey.
 * This test verifies that date objects and various date strings are all
 * normalized to the same "yyyy-MM-dd" format. This is the fix for the
 * duplicate-entry bug.
 */
function testFormatValueForKey_DateHandling() {
  console.log('Running test: testFormatValueForKey_DateHandling');
  const expected = "2024-10-25";

  const testCases = {
    "Date Object": new Date(2024, 9, 25), // native Date object
    "MM/DD/YYYY String": "10/25/2024",     // common string format
    "YYYY-MM-DD String": "2024-10-25",    // ISO-like string
    "Non-date String": "Project X",
    "Empty Value": "",
    "Null Value": null,
  };

  // Test date-like values
  let actual = formatValueForKey(testCases["Date Object"]);
  if (actual !== expected) {
    throw new Error(`Test Failed (Date Object): Expected '${expected}', but got '${actual}'`);
  }

  actual = formatValueForKey(testCases["MM/DD/YYYY String"]);
  if (actual !== expected) {
    throw new Error(`Test Failed (MM/DD/YYYY String): Expected '${expected}', but got '${actual}'`);
  }

  actual = formatValueForKey(testCases["YYYY-MM-DD String"]);
  if (actual !== expected) {
    throw new Error(`Test Failed (YYYY-MM-DD String): Expected '${expected}', but got '${actual}'`);
  }

  // Test non-date values to ensure they are handled correctly
  actual = formatValueForKey(testCases["Non-date String"]);
  if (actual !== "project x") {
    throw new Error(`Test Failed (Non-date String): Expected 'project x', but got '${actual}'`);
  }

  actual = formatValueForKey(testCases["Empty Value"]);
  if (actual !== "") {
    throw new Error(`Test Failed (Empty Value): Expected '', but got '${actual}'`);
  }

  actual = formatValueForKey(testCases["Null Value"]);
  if (actual !== "") {
    throw new Error(`Test Failed (Null Value): Expected '', but got '${actual}'`);
  }

  console.log('Test passed: testFormatValueForKey_DateHandling');
}


/**
 * Test case for the whitespace bug in findRowByValue.
 */
function testFindRowByValue_WhitespaceBug() {
  console.log('Running test: testFindRowByValue_WhitespaceBug');

  let searchValue = '  id-2  ';
  let expectedRow = 3;
  let actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 1 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  searchValue = 'id-1';
  expectedRow = 2;
  actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 2 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  searchValue = 'id-3';
  expectedRow = -1;
  actualRow = findRowByValue(MockSheet, searchValue, 1);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Case 3 Failed: Expected row ${expectedRow}, but got ${actualRow}`);
  }

  console.log('Test passed: testFindRowByValue_WhitespaceBug');
}