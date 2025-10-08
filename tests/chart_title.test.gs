/**
 * @file Test suite for the dashboard chart title generation logic.
 */

// Mock necessary global objects for the test environment
global.Charts = {
  ChartHiddenDimensionStrategy: {
    IGNORE_ROWS: 'IGNORE_ROWS'
  }
};

/**
 * Main entry point for running the chart title tests.
 */
function runChartTitleTests() {
  console.log('--- Running Dashboard Chart Title Tests ---');
  test_upcomingMonthsChartTitle_isNowCorrect();
  console.log('--- Dashboard Chart Title Tests Passed ---');
}

/**
 * Test case to verify the fix for the "Upcoming Months" chart title generation.
 *
 * This test confirms that the chart title now correctly uses the configured
 * time window (`UPCOMING_MONTHS_COUNT`) regardless of the number of months
 * that contain data.
 */
function test_upcomingMonthsChartTitle_isNowCorrect() {
  const testName = 'test_upcomingMonthsChartTitle_isNowCorrect';
  console.log(`Running test: ${testName}`);

  let capturedTitle = '';
  let chartInserted = false;
  let placeholderMessage = '';

  // 1. MOCK the environment
  const mockSheet = {
    getCharts: () => [],
    removeChart: () => {},
    insertChart: () => { chartInserted = true; },
    getRange: () => ({
      setValues: () => mockSheet.getRange(),
      setNumberFormat: () => mockSheet.getRange(),
      clearContent: () => mockSheet.getRange(),
      clearDataValidations: () => mockSheet.getRange(),
      clearNote: () => mockSheet.getRange(),
      getValue: () => 0,
      setValue: () => {}
    }),
    hideColumns: () => {},
    insertColumnsAfter: () => {},
    getMaxColumns: () => 20,
    getMaxRows: () => 100,
    insertRowsAfter: () => {},
    newChart: () => ({
      asColumnChart: () => ({
        addRange: () => mockSheet.newChart().asColumnChart(),
        setNumHeaders: () => mockSheet.newChart().asColumnChart(),
        setHiddenDimensionStrategy: () => mockSheet.newChart().asColumnChart(),
        setOption: (option, value) => {
          if (option === 'title') {
            capturedTitle = value;
          }
          return mockSheet.newChart().asColumnChart();
        },
        setPosition: () => mockSheet.newChart().asColumnChart(),
        build: () => ({})
      })
    })
  };

  global.displayChartPlaceholder = (sheet, anchorRow, anchorCol, message) => {
    placeholderMessage = message;
  };

  const mockConfig = JSON.parse(JSON.stringify(CONFIG));
  mockConfig.DASHBOARD_CHARTING.PAST_MONTHS_COUNT = 0;
  mockConfig.DASHBOARD_CHARTING.UPCOMING_MONTHS_COUNT = 6;

  // 2. SET UP DATA
  const months = generateMonthList(new Date('2025-01-01'), new Date('2025-12-01'));
  const dashboardData = months.map(() => [0, 0, 0, 0]);
  dashboardData[10] = [1, 1, 0, 0]; // 1 upcoming project in November

  // 3. MOCK TIME
  const OriginalDate = global.Date;
  const FAKE_TODAY = '2025-09-15T12:00:00Z';
  global.Date = class extends OriginalDate {
    constructor(...args) {
      if (args.length === 0) { super(FAKE_TODAY); } else { super(...args); }
    }
    static now() { return new OriginalDate(FAKE_TODAY).getTime(); }
  };

  // 4. EXECUTE
  createOrUpdateDashboardCharts(mockSheet, months, dashboardData, mockConfig);

  // 5. RESTORE
  global.Date = OriginalDate;

  // 6. ASSERT
  const expectedCorrectTitle = 'Next 6 Months: Overdue, Upcoming, Total';

  if (!chartInserted) {
    throw new Error(`[${testName}] FAILED: No chart was inserted. Placeholder: "${placeholderMessage}"`);
  }

  // This assertion now checks for the CORRECT title.
  if (capturedTitle !== expectedCorrectTitle) {
    throw new Error(
      `[${testName}] FAILED: The fix is not working as expected. ` +
      `Expected title: "${expectedCorrectTitle}". ` +
      `Actual title: "${capturedTitle}".`
    );
  }

  console.log(`[${testName}] PASSED: The fix is verified. Chart title is now correctly "${capturedTitle}".`);
}