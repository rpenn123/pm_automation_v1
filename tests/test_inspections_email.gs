/**
 * @OnlyCurrentDoc
 *
 * Test Suite: Inspection Email Trigger
 *
 * This test file validates the functionality of the `triggerInspectionEmail` function
 * located in `src/core/Automations.gs`. It ensures that the email is correctly composed
 * and sent when a project's status is updated to 'Inspections'.
 */

/**
 * Test case for the `triggerInspectionEmail` function.
 * It verifies that the function correctly:
 * 1.  Constructs the email with the right subject and body.
 * 2.  Calls the `MailApp.sendEmail` service with the correct parameters.
 * 3.  Handles the lookup for the 'Construction' status from the 'Upcoming' sheet.
 */
function test_sendInspectionEmail_sendsCorrectEmail() {
  // --------- Setup Mocks ---------
  const mockMailApp = {
    sendEmail: function(options) {
      this.lastCall = options;
    },
    lastCall: null
  };
  global.MailApp = mockMailApp;

  const mockSpreadsheet = {
    getSheetByName: function(name) {
      if (name === "Upcoming") {
        return mockUpcomingSheet;
      }
      return null;
    }
  };

  const mockUpcomingSheet = {
    getRange: function(row, col, numRows, numCols) {
      if (numRows > 1) { // Likely a column read for findRowByValue
        return {
          getValues: function() {
            // Return a dummy column of data for the search
            return [["SFID-001"], ["SFID-002"]];
          }
        };
      }
      // single cell read for construction status
      return {
        getValue: function() {
          return "In Progress";
        }
      };
    },
    getLastRow: () => 3
  };

  const e = {
    source: mockSpreadsheet,
    range: {
      getSheet: () => ({ getName: () => "Forecasting" }),
      getRow: () => 2
    },
    value: "Inspections"
  };

  const sourceRowData = [];
  sourceRowData[CONFIG.FORECASTING_COLS.PROJECT_NAME - 1] = "Test Project Alpha";
  sourceRowData[CONFIG.FORECASTING_COLS.EQUIPMENT - 1] = "Elevator Model X";
  sourceRowData[CONFIG.FORECASTING_COLS.LOCATION - 1] = "123 Main St, Anytown";
  sourceRowData[CONFIG.FORECASTING_COLS.SFID - 1] = "SFID-001";

  const correlationId = "test-correlation-id";

  // --------- Execute Function ---------
  triggerInspectionEmail(e, sourceRowData, CONFIG, correlationId);

  // --------- Assertions ---------
  if (!mockMailApp.lastCall) {
    throw new Error("MailApp.sendEmail was not called!");
  }

  const expectedTo = "pm@mobility123.com";
  const expectedSubject = "Re: Inspection Update | Test Project Alpha";
  const expectedBody = `
Project: Test Project Alpha
Status (Progress): Ready for Inspections
Equipment: Elevator Model X
Construction: In Progress
Address: 123 Main St, Anytown
  `.trim();

  if (mockMailApp.lastCall.to !== expectedTo) {
    throw new Error(`Incorrect recipient. Expected: ${expectedTo}, Got: ${mockMailApp.lastCall.to}`);
  }

  if (mockMailApp.lastCall.subject !== expectedSubject) {
    throw new Error(`Incorrect subject. Expected: "${expectedSubject}", Got: "${mockMailApp.lastCall.subject}"`);
  }

  if (mockMailApp.lastCall.body.trim() !== expectedBody) {
    throw new Error(`Incorrect body. Expected: "${expectedBody}", Got: "${mockMailApp.lastCall.body.trim()}"`);
  }

  Logger.log("test_sendInspectionEmail_sendsCorrectEmail passed.");
}
