/**
 * Test suite for Dashboard.gs functionality.
 */

// This test file is preserved as a placeholder.
// The primary tests for the dashboard logic have been moved to more specific files
// like test_Dashboard_HoverNotes.gs to allow for better mocking and isolation.
function test_nonCompleteProjectWithPastDeadline_isCountedAsOverdue() {
    // This test is obsolete due to changes in the overdue logic requirements.
    // The core logic is now tested in test_Dashboard_HoverNotes.gs.
    console.log("Skipping obsolete test: test_nonCompleteProjectWithPastDeadline_isCountedAsOverdue");
    return true;
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