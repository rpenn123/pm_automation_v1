/**
 * @file test_Upcoming_FallbackCompoundKey.gs
 * @description Test suite for verifying the fallback to a compound key (Project Name + Location)
 * when SFID is missing during an "Upcoming" transfer.
 */

function run_test_upcoming_fallback_compound_key() {
  const mockSpreadsheet = new MockSpreadsheet();
  global.SpreadsheetApp = new MockSpreadsheetApp(mockSpreadsheet);

  const testConfig = JSON.parse(JSON.stringify(CONFIG)); // Deep copy

  // Arrange: Create an existing row in "Upcoming" with no SFID
  const upcomingSheet = mockSpreadsheet.getSheetByName(testConfig.SHEETS.UPCOMING);
  upcomingSheet.appendRow(["SFID", "Project Name", "Deadline", "Progress", "Equipment", "Permits", "", "", "", "Location", "Notes"]); // Header
  const initialUpcomingRow = ["", "123 Oak", "10/31/2025", "Scheduled", "Forklift", "Pending", "", "", "", "Ocean City", "Original Notes"];
  upcomingSheet.appendRow(initialUpcomingRow);

  // Arrange: Create a source row in "Forecasting" with no SFID but matching Name and Location
  const forecastingSheet = mockSpreadsheet.getSheetByName(testConfig.SHEETS.FORECASTING);
  forecastingSheet.appendRow(new Array(16).fill("Header")); // Header
  const sourceRowData = new Array(16).fill("");
  sourceRowData[testConfig.FORECASTING_COLS.PROJECT_NAME - 1] = "123 Oak";
  sourceRowData[testConfig.FORECASTING_COLS.DEADLINE - 1] = "10/31/2025";
  sourceRowData[testConfig.FORECASTING_COLS.PROGRESS - 1] = "Scheduled";
  sourceRowData[testConfig.FORECASTING_COLS.EQUIPMENT - 1] = "Bulldozer";
  sourceRowData[testConfig.FORECASTING_COLS.PERMITS - 1] = "Approved";
  sourceRowData[testConfig.FORECASTING_COLS.LOCATION - 1] = "Ocean City";
  forecastingSheet.appendRow(sourceRowData);

  const mockEvent = {
    range: forecastingSheet.getRange(2, 6), // Edit "Permits" column on data row
    source: mockSpreadsheet,
    value: "Approved",
  };

  // Act: Trigger the transfer
  triggerUpcomingTransfer(mockEvent, sourceRowData, testConfig, "test-fallback-corr-id");

  // Assert: Check the state of the "Upcoming" sheet
  const upcomingData = upcomingSheet.getDataRange().getValues();

  // 1. There should be a header and one data row
  console.assert(upcomingData.length === 2, `Assertion Failed: Row count should be 2. Got: ${upcomingData.length}`);

  const updatedRow = upcomingData[1]; // Check the data row

  // 2. Mapped fields should be updated (e.g., Equipment)
  console.assert(updatedRow[testConfig.UPCOMING_COLS.EQUIPMENT - 1] === "Bulldozer", `Assertion Failed: Equipment should be 'Bulldozer'. Got: ${updatedRow[testConfig.UPCOMING_COLS.EQUIPMENT - 1]}`);

  // 3. Unmapped fields should be unchanged
  console.assert(updatedRow[10] === "Original Notes", `Assertion Failed: Unmapped 'Notes' column should be 'Original Notes'. Got: ${updatedRow[10]}`);

  console.log("run_test_upcoming_fallback_compound_key finished.");
}
