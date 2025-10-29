/**
 * @file test_UpdateRow_NonDestructive.gs
 * @description Test suite to guarantee that `updateRowInDestination` only modifies
 * mapped columns, preserving any data in unmapped columns.
 */

function run_test_update_row_non_destructive() {
  const mockSpreadsheet = new MockSpreadsheet();
  global.SpreadsheetApp = new MockSpreadsheetApp(mockSpreadsheet);

  const testConfig = JSON.parse(JSON.stringify(CONFIG));

  // Arrange: Create a destination sheet with a row containing extra, unmapped data
  const destSheet = mockSpreadsheet.getSheetByName("Upcoming"); // Using "Upcoming" as a stand-in
  destSheet.appendRow(["SFID", "Project Name", "Deadline", "Progress", "", "", "", "Construction", "Notes"]); // Header
  const initialData = ["SFID-XYZ", "Project Gamma", "12/01/2025", "In Progress", "", "", "", "Construction data", "Important Note"];
  destSheet.appendRow(initialData);

  // Arrange: Define a transfer config with a limited mapping
  const transferConfig = {
    transferName: "Non-Destructive Test",
    destinationSheetName: "Upcoming",
    destinationColumnMapping: {
      2: 2, // Project Name
      4: 4, // Progress
    }
  };

  // Arrange: New data that will be "merged" in
  const newRowData = ["", "Project Gamma", "", "Inspections", "", "", "", "", ""];

  // Act: Call the function under test directly
  updateRowInDestination(destSheet, 2, newRowData, transferConfig, "test-nd-corr-id"); // Update row 2 (the data row)

  // Assert: Check the sheet's state
  const finalData = destSheet.getDataRange().getValues()[1]; // Check the data row

  // 1. Mapped column "Progress" should be updated
  console.assert(finalData[3] === "Inspections", `Assertion Failed: Progress should be 'Inspections'. Got: ${finalData[3]}`);

  // 2. Unmapped columns should be identical to their initial state
  console.assert(finalData[0] === "SFID-XYZ", `Assertion Failed: SFID column should be unchanged. Got: ${finalData[0]}`);
  console.assert(finalData[7] === "Construction data", `Assertion Failed: Construction column should be unchanged. Got: ${finalData[7]}`);
  console.assert(finalData[8] === "Important Note", `Assertion Failed: Notes column should be unchanged. Got: ${finalData[8]}`);

  console.log("run_test_update_row_non_destructive finished.");
}
