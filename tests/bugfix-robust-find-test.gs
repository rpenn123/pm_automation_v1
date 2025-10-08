/**
 * Test suite for the findRowByProjectNameRobust bug fix.
 */

// Global mocks for Utilities, Session, Logger, etc. are now provided by run_test.js

// =================================================================
// ======================= TEST RUNNER =============================
// =================================================================

/**
 * Test runner for this specific bug fix test.
 */
function runRobustFindTest() {
  console.log('Running tests for findRowByProjectNameRobust bug...');
  testFindRowByProjectName_DateNormalizationBug();
  console.log('Robust find test passed!');
}

// =================================================================
// ======================= TEST CASES ==============================
// =================================================================

/**
 * Test case for the date normalization inconsistency in findRowByProjectNameRobust.
 *
 * This test verifies that the function correctly uses `formatValueForKey` normalization.
 *
 * SCENARIO:
 * - Sheet contains a Date object for "Oct 25, 2024" in row 2.
 * - Sheet contains a string "10/25/2024" in row 3.
 * - We search for the string "10/25/2024".
 *
 * EXPECTED (after fix):
 * - The search term is normalized to "2024-10-25".
 * - The function should scan, normalize row values, and find the *first* match,
 *   which is the Date object in row 2.
 */
function testFindRowByProjectName_DateNormalizationBug() {
  console.log('Running test: testFindRowByProjectName_DateNormalizationBug');

  const mockSheetData = [
    ['Project Name'],
    [new Date(2024, 9, 25)], // Row 2: Actual Date object
    ['10/25/2024'],          // Row 3: String representation of the date
    ['Project Titan']        // Row 4: A regular project name
  ];

  const MockSheet = {
    getLastRow: () => mockSheetData.length,
    getLastColumn: () => 1,
    getRange: (row, column, numRows, numColumns) => {
      const dataSlice = mockSheetData.slice(row - 1, row - 1 + numRows).map(r => r.slice(column - 1, column - 1 + numColumns));
      // The fixed code doesn't use TextFinder, so we just need a simple range mock.
      return {
        getValues: () => dataSlice,
      };
    },
    getName: () => 'TestSheet',
    // This is needed for the error reporting path in the function under test.
    getParent: () => ({})
  };

  const searchTerm = '10/25/2024';
  const projectNameCol = 1;
  const expectedRow = 2; // Should find the Date object first.

  const actualRow = findRowByProjectNameRobust(MockSheet, searchTerm, projectNameCol);

  if (actualRow !== expectedRow) {
    throw new Error(`Test Failed: Expected to find row ${expectedRow}, but got ${actualRow}.`);
  }

  console.log('Test passed: testFindRowByProjectName_DateNormalizationBug');
}