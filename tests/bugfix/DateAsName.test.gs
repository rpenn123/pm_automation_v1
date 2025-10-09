/**
 * @fileoverview Tests for the bug fix related to duplicate detection logic
 * when a project name resembles a date.
 * Verifies that `isDuplicateInDestination` and `findRowByProjectNameRobust`
 * correctly use string comparison instead of date parsing.
 */

/**
 * A simple assertion helper for these tests.
 * @param {any} actual The actual value.
 * @param {any} expected The expected value.
 * @param {string} message The message to display on failure.
 */
function assertEquals(actual, expected, message) {
    if (actual !== expected) {
        throw new Error(`Assertion Failed: ${message}\nExpected: ${expected}\nActual:   ${actual}`);
    }
}

/**
 * Main function to run the tests, following the project's convention.
 */
function runDateAsNameBugfixTests() {
    console.log("  > Running test: isDuplicateInDestination should not find a duplicate for different date-like names...");
    test_isDuplicateInDestination_handlesDateLikeNames();
    console.log("    ... PASSED");

    console.log("  > Running test: findRowByProjectNameRobust should not find a match for different date-like names...");
    test_findRowByProjectNameRobust_handlesDateLikeNames();
    console.log("    ... PASSED");
}

/**
 * Mocks a sheet with the given 2D array of data.
 * @param {any[][]} data The data for the sheet.
 * @returns {object} A mock sheet object.
 */
function mockSheetForBugfix(data) {
    return {
        getRange: function(row, col, numRows, numCols) {
            const requestedData = [];
            const r = row || 1;
            const c = col || 1;
            const nr = numRows || 1;
            const nc = numCols || 1;
            for (let i = 0; i < nr; i++) {
                const rowData = data[r + i - 1] || [];
                const sliced = rowData.slice(c - 1, c - 1 + nc);
                requestedData.push(sliced);
            }
            return {
                getValues: () => requestedData
            };
        },
        getLastRow: () => data.length,
        getMaxColumns: () => (data[0] ? data[0].length : 0),
    };
}

/**
 * Tests that `isDuplicateInDestination` correctly distinguishes between
 * two different project names that look like different date formats.
 */
function test_isDuplicateInDestination_handlesDateLikeNames() {
    // Setup
    const existingData = [
        ["Project Name"],
        ["May 10, 2024"],
    ];
    const destinationSheet = mockSheetForBugfix(existingData);
    const sourceRowData = ["5/10/2024"];
    const projectName = "5/10/2024";
    const dupConfig = {
        projectNameDestCol: 1
    };

    // Execute
    const isDuplicate = isDuplicateInDestination(
        destinationSheet,
        null, // sfid
        projectName,
        sourceRowData,
        sourceRowData.length,
        dupConfig,
        "test-correlation-id"
    );

    // Assert
    assertEquals(isDuplicate, false, "Should not have found a duplicate for '5/10/2024' vs 'May 10, 2024'.");
}

/**
 * Tests that `findRowByProjectNameRobust` correctly distinguishes between
 * two different project names that look like different date formats.
 */
function test_findRowByProjectNameRobust_handlesDateLikeNames() {
    // Setup
    const existingData = [
        ["Header"],
        ["May 10, 2024"], // Row 2
    ];
    const sheet = mockSheetForBugfix(existingData);
    const projectNameToFind = "5/10/2024";

    // Execute
    const foundRow = findRowByProjectNameRobust(sheet, projectNameToFind, 1);

    // Assert
    assertEquals(foundRow, -1, "Should not have found row for '5/10/2024' when 'May 10, 2024' exists.");
}