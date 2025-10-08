/**
 * Test suite for Dashboard.gs functionality.
 */

/**
 * Test case to ensure that only ACTIVE projects with a past deadline are counted as "Overdue".
 * A project that is non-complete (e.g., "On Hold") but not active should not be counted.
 */
function test_nonActiveProject_isNotCountedAsOverdue() {
  // 1. Setup
  const testName = "test_nonActiveProject_isNotCountedAsOverdue";
  let assertions = 0;
  let failures = 0;
  console.log(`RUNNING: ${testName}...`);

  // Mock CONFIG object
  const mockConfig = {
    FORECASTING_COLS: {
      DEADLINE: 1,
      PROGRESS: 2,
      PERMITS: 3,
    },
    STATUS_STRINGS: {
      PERMIT_APPROVED: "Approved",
      IN_PROGRESS: "In Progress",
      SCHEDULED: "Scheduled",
      COMPLETED: "Completed",
      CANCELLED: "Cancelled",
    },
  };

  // Mock Data
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayString = Utilities.formatDate(yesterday, "GMT", "MM/dd/yyyy");

  // This project is "On Hold", which is not an "active" or "complete" status.
  // With the bug, it will be incorrectly counted as Overdue.
  const mockForecastingValues = [
    [yesterdayString, "On Hold", "Pending"],
  ];

  // 2. Execution
  // Manually load dependencies since GAS doesn't have a module system.
  // In a real test environment, these would be loaded globally.
  const dependencies = `
    ${normalizeString.toString()}
    ${parseAndNormalizeDate.toString()}
    ${processDashboardData.toString()}
  `;

  // Note: This is a simulation. In a real GAS environment, we would just call the function.
  // For this simulation, we assume the functions are available.
  const result = processDashboardData(mockForecastingValues, mockConfig);
  const { monthlySummaries } = result;

  const key = yesterday.getFullYear() + '-' + yesterday.getMonth();
  const summary = monthlySummaries.get(key);
  const overdueCount = summary ? summary[2] : 0;

  // 3. Assertion
  try {
    // The correct behavior is that the "On Hold" project should NOT be counted as overdue.
    assertEquals(0, overdueCount, "Overdue count should be 0 for a non-active (On Hold) project.");
    assertions++;
  } catch (e) {
    console.error(`${testName} FAILED: ${e.message}`);
    failures++;
  }

  // 4. Teardown & Reporting
  if (failures === 0) {
    console.log(`${testName} PASSED (${assertions} assertions).`);
  } else {
    console.error(`${testName} FINISHED WITH ${failures} FAILURES.`);
  }

  return failures === 0;
}

/**
 * A simple assertion helper for testing. In a real scenario, this would be in a shared test library.
 */
function assertEquals(expected, actual, message) {
  if (expected !== actual) {
    throw new Error(message + ` Expected: ${expected}, but got: ${actual}`);
  }
}

// In a real GAS test runner, you'd have a main function to run all tests.
// function runAllDashboardTests() {
//   test_nonActiveProject_isNotCountedAsOverdue();
// }

// Since we cannot run this in a real GAS environment, I will now manually trace the execution
// to prove the test fails with the current code.

/*
MANUAL TRACE (FAILURE PROOF):
1. `processDashboardData` is called with a project that has a status "On Hold" and a deadline of yesterday.
2. `today` is calculated. `deadlineDate` is parsed to yesterday's date.
3. `isComplete` is `false` because "On Hold" is not "Completed" or "Cancelled".
4. `isActive` is `false` because "On Hold" is not "In Progress" or "Scheduled".
5. The code enters the `if (!isComplete)` block.
6. The condition `if (deadlineDate < today)` is checked. `yesterday < today` is TRUE.
7. The code enters the `if` block and increments `monthData[2]` (the overdue count).
8. The function returns. `overdueCount` is 1.
9. The test's assertion `assertEquals(0, 1, ...)` fails.

This confirms the test will fail as expected.
*/