/**
 * @OnlyCurrentDoc
 * BugFixTest.gs
 * This file contains the test case for a bug fix in `TransferEngine.gs`.
 * It specifically tests the `isDuplicateInDestination` function to ensure it correctly
 * identifies duplicates when a project name in the destination sheet is a Date object.
 */

/**
 * A specific unit test to verify a bug fix in the `isDuplicateInDestination` function.
 *
 * **Bug Context:** The `isDuplicateInDestination` function could fail to detect a duplicate
 * if the destination sheet contained a Date object in its project name column, while the
 * source data being checked had a string representation of that same date. This was because
 * the values were not being consistently normalized before comparison.
 *
 * **Test Scenario:**
 * 1.  **Mocks** a destination sheet where a "Project Name" is a `Date` object.
 * 2.  **Defines** source data where the `projectName` is a string that matches the date in the mock sheet.
 * 3.  **Executes** `isDuplicateInDestination` with this data.
 * 4.  **Asserts** that the function correctly returns `true`, confirming that the normalization
 *     logic inside the function (via `formatValueForKey`) handles this type mismatch correctly.
 *
 * This function is intended to be run manually from the Apps Script editor to validate the fix.
 * @returns {void} Throws an error if the assertion fails.
 */
function testIsDuplicateInDestination_DateInProjectName() {
  // 1. Setup the test environment
  const testName = "testIsDuplicateInDestination_DateInProjectName";
  let assertions = 0;

  // Mock destination sheet object
  const mockSheet = {
    _data: [
      ["Header1", "Header2", "Header3"],
      [new Date("2024-09-29T00:00:00.000Z"), "Some Value", new Date("2025-01-01T00:00:00.000Z")] // Project Name is a Date object
    ],
    getLastRow: function() { return this._data.length; },
    getLastColumn: function() { return this._data[0].length; },
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

  // 2. Define the test case data
  const sfid = null; // Test the fallback compound key logic
  const projectName = "2024-09-29"; // Source project name is a string
  const sourceRowData = ["", projectName, "", "", new Date("2025-01-01T00:00:00.000Z")]; // Mock source row data
  const sourceReadWidth = sourceRowData.length;

  // Configuration for the duplicate check
  const dupConfig = {
    projectNameDestCol: 1, // Project Name is in column A
    compoundKeySourceCols: [5], // e.g., a deadline date
    compoundKeyDestCols: [3]    // in column C
  };

  // 3. Execute the function under test
  const isDuplicate = isDuplicateInDestination(
    mockSheet,
    sfid,
    projectName,
    sourceRowData,
    sourceReadWidth,
    dupConfig
  );

  // 4. Assert the result
  if (isDuplicate === true) {
    console.log(`✅ ${testName}: PASSED - Correctly identified the duplicate.`);
    assertions++;
  } else {
    console.error(`❌ ${testName}: FAILED - Did not identify the duplicate. Expected true, got ${isDuplicate}.`);
  }

  // This is a standalone test, so we can't use a test runner.
  // We'll just log the result. A real environment would have a proper assertion library.
  if (assertions !== 1) {
    throw new Error(`${testName}: One or more assertions failed.`);
  }
}


/**
 * A specific unit test to verify the timezone bug fix in `isDuplicateInDestination`.
 *
 * **Bug Context:** The `formatValueForKey` function previously used the script's local
 * timezone. This meant a UTC date like `2024-10-31T02:00:00Z` could be formatted as
 * `"2024-10-30"` if the script ran in a timezone like `America/New_York`. This test
 * simulates that exact scenario.
 *
 * **Test Scenario:**
 * 1.  **Mocks** a destination sheet where the project name is a `Date` object representing
 *     an early morning UTC time.
 * 2.  **Defines** source data where the `projectName` is a string matching the *correct* UTC date.
 * 3.  **Executes** `isDuplicateInDestination`. With the bug, this would fail because the
 *     destination date would be formatted as the previous day.
 * 4.  **Asserts** that the function returns `true`, proving that the UTC-based formatting
 *     now works correctly, preventing timezone-related mis-matches.
 *
 * @returns {void} Throws an error if the assertion fails.
 */
function testIsDuplicateInDestination_TimezoneBug() {
  const testName = "testIsDuplicateInDestination_TimezoneBug";
  let assertions = 0;

  // Mock destination sheet with a date that would be the previous day in a non-UTC timezone
  const mockSheet = {
    _data: [
      ["Project Name", "Deadline"],
      // This is 2 AM UTC on Oct 31. In America/New_York, it's still Oct 30.
      [new Date("2024-10-31T02:00:00.000Z"), "Some Value"]
    ],
    getLastRow: function() { return this._data.length; },
    getLastColumn: function() { return this._data[0].length; },
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

  // Source project name is a string representing the correct UTC date.
  const projectName = "2024-10-31";
  const sourceRowData = [projectName, "Some Value"];
  const sourceReadWidth = sourceRowData.length;

  // Config for duplicate check (no compound key needed for this test)
  const dupConfig = {
    projectNameDestCol: 1,
    compoundKeySourceCols: [],
    compoundKeyDestCols: []
  };

  // Execute the function
  const isDuplicate = isDuplicateInDestination(
    mockSheet,
    null, // sfid
    projectName,
    sourceRowData,
    sourceReadWidth,
    dupConfig
  );

  // Assert
  if (isDuplicate === true) {
    console.log(`✅ ${testName}: PASSED - Correctly identified duplicate despite timezone difference.`);
    assertions++;
  } else {
    console.error(`❌ ${testName}: FAILED - Did not identify duplicate. Expected true, got ${isDuplicate}.`);
  }

  if (assertions !== 1) {
    throw new Error(`${testName}: Assertion failed.`);
  }
}


/**
 * A specific unit test to verify a bug fix in the `findRowByProjectNameRobust` function.
 *
 * **Bug Context:** The `findRowByProjectNameRobust` function could fail if the project
 * name being searched for was a date. The string representation of a Date object in a cell
 * (e.g., "Tue Oct 31 2024...") would not match the simple search string (e.g., "2024-10-31").
 *
 * **Test Scenario:**
 * 1.  **Mocks** a sheet where a "Project Name" is a `Date` object.
 * 2.  **Defines** a search term that is a string representation of that date.
 * 3.  **Executes** `findRowByProjectNameRobust`.
 * 4.  **Asserts** that the function correctly returns the row number `2`, proving that the
 *     lookup can handle the type mismatch between the search string and the cell's Date object.
 *
 * @returns {void} Throws an error if the assertion fails.
 */
function testFindRowByProjectName_DateInProjectName() {
  const testName = "testFindRowByProjectName_DateInProjectName";
  let assertions = 0;

  // Mock sheet object where a project name is a Date
  const mockSheet = {
    _data: [
      ["Project Name", "Status"],
      [new Date("2024-10-31T00:00:00.000Z"), "In Progress"]
    ],
    getLastRow: function() { return this._data.length; },
    getLastColumn: function() { return this._data[0].length; },
    getRange: function(row, col, numRows, numCols) {
      const self = this;
      const range = {
        getValues: function() {
          const result = [];
          for (let i = 0; i < numRows; i++) {
            const rowData = self._data[row + i - 1];
            if (rowData) {
              result.push(rowData.slice(col - 1, col - 1 + numCols));
            }
          }
          return result;
        },
        // Mock text finder for robustness check
        createTextFinder: function(text) {
            return {
                matchCase: function() { return this; },
                matchEntireCell: function() { return this; },
                findNext: function() { return null; } // Simulate TextFinder failing to force fallback
            };
        }
      };
      return range;
    },
    // The robust function needs a parent to call getParent() on error
    getParent: function() { return null; }
  };

  // The project name to search for (as a string)
  const projectNameToFind = "2024-10-31";
  const projectNameCol = 1;

  // Execute the function under test
  // NOTE: This test will FAIL until the fix is applied in Utilities.gs
  const foundRow = findRowByProjectNameRobust(mockSheet, projectNameToFind, projectNameCol);

  // Assert the result
  const expectedRow = 2;
  if (foundRow === expectedRow) {
    console.log(`✅ ${testName}: PASSED - Correctly found row ${expectedRow}.`);
    assertions++;
  } else {
    console.error(`❌ ${testName}: FAILED - Expected row ${expectedRow}, but got ${foundRow}.`);
  }

  if (assertions !== 1) {
    throw new Error(`${testName}: Assertion failed.`);
  }
}