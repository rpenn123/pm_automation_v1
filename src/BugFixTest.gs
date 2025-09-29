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