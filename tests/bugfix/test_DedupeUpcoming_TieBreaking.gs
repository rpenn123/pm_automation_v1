/**
 * Test suite for DedupeUpcoming.gs bug fix.
 */

function runDedupeUpcomingTieBreakerTest() {
  test_dedupeUpcoming_keepsNewestRowOnTimestampTie();
}

/**
 * Verifies that the deduplication logic correctly keeps the newest (bottom-most)
 * row when multiple rows have the same timestamp or no timestamp.
 */
function test_dedupeUpcoming_keepsNewestRowOnTimestampTie() {
  const testName = 'test_dedupeUpcoming_keepsNewestRowOnTimestampTie';

  // 1. ARRANGE

  // Mock data: Three identical projects, all with the same blank timestamp.
  // The logic should keep the one at row index 5 (the last one).
  const mockData = [
    // Headers (row 1)
    ["Project Name", "SFID", "Location", "Last Edit At (hidden)"],
    // Data rows
    ["Test Project", "SF123", "Site A", ""], // Should be deleted (row 2)
    ["Another Project", "SF456", "Site B", ""], // Should be kept (unique)
    ["Test Project", "SF123", "Site A", ""], // Should be deleted (row 4)
    ["Test Project", "SF123", "Site A", ""], // Should be KEPT (row 5, newest)
  ];

  // Store which rows are deleted
  const deletedRows = [];

  // Mock Sheet
  const mockSheet = {
    getName: () => 'Upcoming',
    getLastRow: () => mockData.length,
    getMaxColumns: () => mockData[0].length,
    getRange: function(row, col, numRows, numCols) {
      const requestedData = [];
      // Adjust for 0-based array index
      const startRow = row - 1;
      for (let i = 0; i < numRows; i++) {
        const currentRow = mockData[startRow + i];
        if (currentRow) {
          const slicedRow = currentRow.slice(col - 1, col - 1 + numCols);
          requestedData.push(slicedRow);
        }
      }
      return {
        getValues: () => requestedData,
      };
    },
    deleteRow: function(rowIndex) {
      deletedRows.push(rowIndex);
    },
  };

  // Mock SpreadsheetApp
  global.SpreadsheetApp = {
    getActive: () => ({
      getSheetByName: () => mockSheet,
    }),
  };

  // Mock CONFIG
  global.CONFIG = {
    SHEETS: {
      UPCOMING: 'Upcoming',
    },
    UPCOMING_COLS: {
      PROJECT_NAME: 1,
      SFID: 2,
      LOCATION: 3,
    },
    LAST_EDIT: {
      AT_HEADER: 'Last Edit At (hidden)',
    },
  };

  // Mock Utilities needed by the function under test
  global.getHeaderColumnIndex = function(sheet, header) {
    const headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
    return headers.indexOf(header) + 1;
  };
  global.normalizeString = function(str) {
    return (str || "").toLowerCase().trim();
  }

  // 2. ACT
  // Load the function into the test scope and run it
  const fs = require('fs');
  eval(fs.readFileSync('src/maintenance/DedupeUpcoming.gs', 'utf8'));
  dedupeUpcomingBySfidOrNameLoc();

  // 3. ASSERT
  const expectedDeletedRows = [2, 4];
  // Sort for consistent comparison
  deletedRows.sort((a,b) => a - b);

  let passed = true;
  if (deletedRows.length !== expectedDeletedRows.length) {
    passed = false;
  } else {
    for (let i = 0; i < deletedRows.length; i++) {
      if (deletedRows[i] !== expectedDeletedRows[i]) {
        passed = false;
        break;
      }
    }
  }

  if (!passed) {
    throw new Error(`${testName} FAILED: Expected rows [${expectedDeletedRows.join(', ')}] to be deleted, but got [${deletedRows.join(', ')}].`);
  } else {
    Logger.log(`${testName} PASSED`);
  }

}
