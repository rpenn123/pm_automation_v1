/**
 * @file test_Upcoming_SyncOnDuplicate.gs
 * @description Test suite for verifying that the "Upcoming" transfer correctly updates
 * an existing row when a duplicate is found and `syncOnDuplicate` is enabled.
 */

function run_test_upcoming_sync_on_duplicate() {
  const mockSpreadsheet = new MockSpreadsheet();
  global.SpreadsheetApp = new MockSpreadsheetApp(mockSpreadsheet);

  const testConfig = JSON.parse(JSON.stringify(CONFIG)); // Deep copy for isolation

  // Arrange: Create an existing row in the "Upcoming" sheet
  const upcomingSheet = mockSpreadsheet.getSheetByName(testConfig.SHEETS.UPCOMING);
  upcomingSheet.appendRow(["SFID", "Project Name", "Deadline", "Progress", "Equipment", "Permits", "", "", "", "Location", "Notes"]); // Header
  const initialUpcomingRow = ["SFID-123", "Project Alpha", "10/31/2025", "Pre-con", "Excavator", "Pending", "", "", "", "Ocean City", "Notes here"];
  upcomingSheet.appendRow(initialUpcomingRow);

  // Arrange: Create a source row in the "Forecasting" sheet with the same SFID
  const forecastingSheet = mockSpreadsheet.getSheetByName(testConfig.SHEETS.FORECASTING);
  forecastingSheet.appendRow(new Array(16).fill("Header")); // Header
  const sourceRowData = new Array(16).fill("");
  sourceRowData[testConfig.FORECASTING_COLS.SFID - 1] = "SFID-123";
  sourceRowData[testConfig.FORECASTING_COLS.PROJECT_NAME - 1] = "Project Alpha";
  sourceRowData[testConfig.FORECASTING_COLS.DEADLINE - 1] = "11/15/2025";
  sourceRowData[testConfig.FORECASTING_COLS.PROGRESS - 1] = "Ready";
  sourceRowData[testConfig.FORECASTING_COLS.EQUIPMENT - 1] = "Crane";
  sourceRowData[testConfig.FORECASTING_COLS.PERMITS - 1] = "Approved";
  sourceRowData[testConfig.FORECASTING_COLS.LOCATION - 1] = "Ocean City";
  forecastingSheet.appendRow(sourceRowData);

  const mockEvent = {
    range: forecastingSheet.getRange(2, 4), // Edit "Progress" column on data row
    source: mockSpreadsheet,
    value: "Ready",
  };

  // Act: Trigger the transfer
  triggerUpcomingTransfer(mockEvent, sourceRowData, testConfig, "test-corr-id");

  // Assert: Check the state of the "Upcoming" sheet
  const upcomingData = upcomingSheet.getDataRange().getValues();

  // 1. There should be a header and one data row
  console.assert(upcomingData.length === 2, "Assertion Failed: Row count should be 2. Got: " + upcomingData.length);

  const updatedRow = upcomingData[1]; // Check the data row

  // 2. Mapped fields should be updated
  console.assert(updatedRow[testConfig.UPCOMING_COLS.PROGRESS - 1] === "Ready", "Assertion Failed: Progress should be 'Ready'.");
  console.assert(updatedRow[testConfig.UPCOMING_COLS.DEADLINE - 1] === "11/15/2025", "Assertion Failed: Deadline should be updated.");

  // 3. Unmapped fields should remain unchanged
  console.assert(updatedRow[10] === "Notes here", "Assertion Failed: Unmapped 'Notes' column should be unchanged.");

  console.log("run_test_upcoming_sync_on_duplicate finished.");
}
