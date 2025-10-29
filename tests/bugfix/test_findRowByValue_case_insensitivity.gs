/**
 * Test suite for bugfix in Utilities.gs
 */

function runBugfixTests() {
  test_findRowByValue_caseInsensitive();
}

/**
 * Tests the `findRowByValue` function to ensure it performs a case-insensitive search.
 * This is a bugfix validation test.
 */
function test_findRowByValue_caseInsensitive() {
  const testName = 'test_findRowByValue_caseInsensitive';

  // 1. ARRANGE
  const mockSheet = {
    getLastRow: function() {
      return 3;
    },
    getRange: function(row, col, numRows, numCols) {
      const data = [
        ['SFID-001'],
        ['sfid-002']
      ];
      return {
        getValues: function() {
          return data;
        }
      };
    }
  };

  const valueToFind = 'SFID-002';
  const expectedRow = 3;

  // 2. ACT
  const result = findRowByValue(mockSheet, valueToFind, 1);

  // 3. ASSERT
  if (result !== expectedRow) {
    throw new Error(`${testName} FAILED: Expected to find value "${valueToFind}" at row ${expectedRow}, but got ${result}. The search is likely case-sensitive.`);
  } else {
    Logger.log(`${testName} PASSED`);
  }
}
