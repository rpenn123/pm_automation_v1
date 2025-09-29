/**
 * @OnlyCurrentDoc
 * BugFixTest.gs
 * This file contains a test case to demonstrate and verify the fix for a bug
 * in the TransferEngine where the source data read width was calculated incorrectly.
 */

/**
 * Creates temporary sheets for the test.
 * This is a helper to be called by the main test function.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The active spreadsheet.
 * @returns {{sourceSheet: GoogleAppsScript.Spreadsheet.Sheet, destSheet: GoogleAppsScript.Spreadsheet.Sheet}}
 */
function setupTestSheets(ss) {
  // Clean up old sheets first
  cleanupTestSheets(ss);

  const sourceSheet = ss.insertSheet("Test_Source");
  const destSheet = ss.insertSheet("Test_Destination");

  // Set up headers
  sourceSheet.getRange("A1:J1").setValues([["Project", "Col2", "Col3", "Col4", "Col5", "Col6", "Col7", "Col8", "Col9", "Critical_Data"]]);
  destSheet.getRange("A1:B1").setValues([["Project_Name", "Mapped_Data"]]);

  return { sourceSheet, destSheet };
}

/**
 * Removes the temporary sheets created by the test.
 * Can be run manually for cleanup.
 */
function cleanupTestSheets(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Test_Source");
  const destSheet = ss.getSheetByName("Test_Destination");

  if (sourceSheet) ss.deleteSheet(sourceSheet);
  if (destSheet) ss.deleteSheet(destSheet);
}


/**
 * Test case to demonstrate the bug in `executeTransfer`.
 * This test sets up a scenario where a column in `destinationColumnMapping` has a higher
 * index than any column in `sourceColumnsNeeded`.
 *
 * TO RUN THIS TEST:
 * 1. Open the Google Apps Script editor.
 * 2. Select this function (`runTransferWidthBugTest`) from the function dropdown.
 * 3. Click "Run".
 * 4. Inspect the "Test_Destination" sheet.
 *
 * EXPECTED RESULT (BEFORE FIX):
 * The "Mapped_Data" column for "Project Bug" will be EMPTY.
 *
 * EXPECTED RESULT (AFTER FIX):
 * The "Mapped_Data" column for "Project Bug" will contain the value "IMPORTANT".
 */
function runTransferWidthBugTest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { sourceSheet, destSheet } = setupTestSheets(ss);

  // 1. Add sample data to the source sheet.
  // The critical data is in column 10 (J).
  sourceSheet.getRange("A2:J2").setValues([["Project Bug", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "IMPORTANT"]]);

  // 2. Create a mock onEdit event object.
  const e = {
    range: sourceSheet.getRange("A2"),
    source: ss
  };

  // 3. Define the transfer configuration that exposes the bug.
  // `sourceColumnsNeeded` only goes up to column 2.
  // `destinationColumnMapping` needs data from column 10.
  // The bug prevents column 10 from being read.
  const bugConfig = {
    transferName: "Bug Test Transfer",
    destinationSheetName: "Test_Destination",
    sourceColumnsNeeded: [1, 2], // Max is 2
    destinationColumnMapping: {
      1: 1,  // Project Name -> Project_Name
      10: 2  // Critical_Data -> Mapped_Data (This is the key part)
    },
    duplicateCheckConfig: {
      checkEnabled: true,
      projectNameSourceCol: 1,
      projectNameDestCol: 1
    }
  };

  // 4. Execute the transfer.
  try {
    executeTransfer(e, bugConfig);
    SpreadsheetApp.flush(); // Ensure all changes are written
    Logger.log("Test finished. Check the 'Test_Destination' sheet.");
    // In a real test framework, you'd add assertions here.
    // For manual testing, we observe the sheet.
  } catch (error) {
    Logger.log("Test failed with an error: " + error.toString());
  }
}