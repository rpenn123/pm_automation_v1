/**
 * Test suite for the overdue bug fix.
 * To run this test:
 * 1. Open the Google Apps Script editor.
 * 2. Select the function `runOverdueBugFixTests` from the function dropdown list.
 * 3. Click the "Run" button.
 * 4. View the logs by clicking "Execution log" in the left-hand menu.
 *
 * This test is designed to FAIL before the bug is fixed and PASS after the fix is applied.
 */
function runOverdueBugFixTests() {
  const testName = "testShouldCountNonActiveButNotCompleteProjectAsOverdue";
  Logger.log(`Executing test suite for overdue bug fix...`);

  try {
    testShouldCountNonActiveButNotCompleteProjectAsOverdue();
    Logger.log(`---`);
    Logger.log(`SUCCESS: Test case "${testName}" passed. The fix appears to be working correctly.`);
    Logger.log(`---`);
  } catch (e) {
    Logger.log(`---`);
    Logger.log(`FAILURE: Test case "${testName}" failed. This is the EXPECTED outcome before the bug is fixed.`);
    Logger.log(`Error: ${e.message}`);
    Logger.log(`---`);
    // Re-throwing the error to ensure the execution log in Apps Script clearly shows a failure.
    throw e;
  }
}

/**
 * This test verifies that a project with a status that is neither "active" (In Progress, Scheduled)
 * nor "complete" (Completed, Cancelled) is correctly counted as overdue if its deadline has passed.
 * The original bug failed to count such projects (e.g., status "On Hold").
 */
function testShouldCountNonActiveButNotCompleteProjectAsOverdue() {
  const testName = "testShouldCountNonActiveButNotCompleteProjectAsOverdue";
  Logger.log(`Running test: ${testName}`);

  // 1. ARRANGE: Set up the test environment.
  // This mock config simulates the project's configuration for the test.
  const mockConfig = {
    FORECASTING_COLS: {
      DEADLINE: 10,
      PROGRESS: 7,
      PERMITS: 8,
    },
    // Define all relevant statuses. "On Hold" is the key to this test case.
    STATUS_STRINGS: {
      IN_PROGRESS: "In Progress",
      SCHEDULED: "Scheduled",
      COMPLETED: "Completed",
      CANCELLED: "Cancelled",
      ON_HOLD: "On Hold",
      PERMIT_APPROVED: "approved",
    }
  };

  // Create a date for yesterday to represent a past deadline.
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  // This mock data represents the exact case the bug misses: a project that is
  // "On Hold" and has a past deadline. It should be considered overdue.
  const mockForecastingValues = [
    createMockRowForOverdueTest(mockConfig.STATUS_STRINGS.ON_HOLD, yesterday, mockConfig.FORECASTING_COLS)
  ];

  // 2. ACT: Execute the function under test.
  // This function is defined in `src/ui/Dashboard.gs` and is globally available.
  const result = processDashboardData(mockForecastingValues, mockConfig);

  // 3. ASSERT: Check if the outcome is as expected.
  const expectedOverdueCount = 1;
  const actualOverdueCount = result.allOverdueItems.length;

  if (actualOverdueCount !== expectedOverdueCount) {
    // This is the failure condition. We throw an error to make the test failure explicit.
    const errorMessage = `FAIL: ${testName}. Expected ${expectedOverdueCount} overdue item, but found ${actualOverdueCount}.`;
    throw new Error(errorMessage);
  } else {
    // This is the success condition.
    Logger.log(`PASS: ${testName}. Correctly identified ${actualOverdueCount} overdue item.`);
  }
}

/**
 * Helper function to create a mock data row for the overdue status test.
 * @param {string} status The project status (e.g., "On Hold").
 * @param {Date} deadline The project deadline.
 * @param {object} cols The column configuration object from CONFIG.
 * @returns {Array<any>} A mock spreadsheet row array.
 */
function createMockRowForOverdueTest(status, deadline, cols) {
  // Ensure the array is long enough to hold all necessary columns.
  const rowLength = Math.max(cols.DEADLINE, cols.PROGRESS, (cols.PERMITS || 0));
  const row = new Array(rowLength).fill(null);

  // Populate the columns relevant to the test.
  row[cols.PROGRESS - 1] = status;
  row[cols.DEADLINE - 1] = deadline;

  return row;
}

/*
NOTE: For this test to run in the Apps Script environment, the following functions
must be available in the global scope, which is true as they are in other .gs files
within the same Clasp project:
- processDashboardData (from src/ui/Dashboard.gs)
- parseAndNormalizeDate (from src/core/Utilities.gs)
- normalizeString (from src/core/Utilities.gs)
*/