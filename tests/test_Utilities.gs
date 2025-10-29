/**
 * Test suite for Utilities.gs
 */

function runUtilityTests() {
  test_getHeaderColumnIndex_withEmptyTrailingColumns();
  test_formatValueForKey_handlesDateLikeStrings();
}

/**
 * Tests the `formatValueForKey` function for robust key normalization.
 * This test verifies that the function correctly:
 * 1.  Normalizes actual `Date` objects to a "yyyy-MM-dd" format.
 * 2.  Normalizes valid date strings (e.g., "5/10/2024") to the same "yyyy-MM-dd" format.
 * 3.  Treats non-date strings, including numeric strings and strings containing dates, as simple,
 *     lowercased strings, preventing them from being misinterpreted as dates.
 */
function test_formatValueForKey_handlesDateLikeStrings() {
    const testName = 'test_formatValueForKey_handlesDateLikeStrings';
    let passed = true;
    const failures = [];

    // Test cases for values that should be treated as DATES
    const dateCases = {
        '5/10/2024': '2024-05-10',
        '2024-03-15': '2024-03-15',
        '01-05-2023': '2023-01-05',
    };

    // Test cases for values that should be treated as STRINGS
    const stringCases = {
        'Project 5/10/2024': 'project 5/10/2024',
        '12345': '12345', // Numeric string should not be a date
        'Main Street Bridge': 'main street bridge',
        '': '',
    };

    // --- 1. ARRANGE & 2. ACT ---

    // Test date-like strings that SHOULD be converted
    for (const input in dateCases) {
        const expected = dateCases[input];
        const actual = formatValueForKey(input);
        if (actual !== expected) {
            passed = false;
            failures.push(`FAIL: Input "${input}" (date string) | Expected: "${expected}" | Got: "${actual}"`);
        }
    }

    // Test strings that should NOT be converted
    for (const input in stringCases) {
        const expected = stringCases[input];
        const actual = formatValueForKey(input);
        if (actual !== expected) {
            passed = false;
            failures.push(`FAIL: Input "${input}" (string) | Expected: "${expected}" | Got: "${actual}"`);
        }
    }

    // Test a real Date object
    const realDate = new Date('2024-10-05T12:00:00Z'); // Use noon to avoid timezone issues
    const expectedDateString = '2024-10-05';
    const actualDateResult = formatValueForKey(realDate);
    if (actualDateResult !== expectedDateString) {
        passed = false;
        failures.push(`FAIL: Input Date Object | Expected: "${expectedDateString}" | Got: "${actualDateResult}"`);
    }

    // --- 3. ASSERT ---
    if (!passed) {
        throw new Error(`${testName} FAILED:\n- ${failures.join('\n- ')}`);
    } else {
        Logger.log(`${testName} PASSED`);
    }
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