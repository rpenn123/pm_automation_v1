/**
 * @OnlyCurrentDoc
 * Dashboard.gs
 * This file contains all logic for generating the main project dashboard, including data
 * processing, report generation, chart creation, and the "Overdue Details" drill-down sheet.
 * It is designed to be highly configurable and efficient.
 */

/**
 * Main orchestrator function to generate or update the entire Dashboard.
 * This function is the primary entry point for creating the dashboard, typically called from a custom menu.
 *
 * **Execution Flow:**
 * 1.  **Initialization:** Gets UI, Spreadsheet, and sheet objects.
 * 2.  **Data Reading:** Calls `readForecastingData` to get the raw data from the 'Forecasting' sheet.
 * 3.  **Data Processing:** Calls `processForecastingData` for a highly efficient single pass over the
 *     raw data to produce all necessary aggregates (monthly summaries, grand totals, overdue items).
 * 4.  **Overdue Details:** Populates the 'Overdue Details' sheet with the list of overdue projects.
 * 5.  **Dashboard Rendering:**
 *     - Clears and prepares the main 'Dashboard' sheet.
 *     - Writes headers and informational notes.
 *     - Populates the monthly summary data, including hyperlink formulas to the overdue sheet.
 *     - Writes grand totals and a count of items with missing deadlines.
 * 6.  **Formatting & Charting:** Applies all visual formatting and calls `createOrUpdateDashboardCharts`
 *     to generate the data visualizations.
 *
 * All steps are wrapped in a try-catch block for robust error handling and user notification.
 * @returns {void}
 */
function updateDashboard() {
  const ui = SpreadsheetApp.getUi();
  const scriptStartTime = new Date();
  Logger.log(`Dashboard update initiated at ${scriptStartTime.toLocaleString()}`);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const { FORECASTING, DASHBOARD, OVERDUE_DETAILS } = CONFIG.SHEETS;

    const forecastSheet = ss.getSheetByName(FORECASTING);
    if (!forecastSheet) throw new Error(`Sheet "${FORECASTING}" not found.`);

    // Initialize destination sheets
    const dashboardSheet = getOrCreateSheet(ss, DASHBOARD);
    const overdueDetailsSheet = getOrCreateSheet(ss, OVERDUE_DETAILS);
    const overdueSheetGid = overdueDetailsSheet.getSheetId();

    // 1. Read data
    const { forecastingValues, forecastingHeaders } = readForecastingData(forecastSheet);
    if (!forecastingValues) throw new Error(`Failed to read data from ${FORECASTING}.`);

    // 2. Process data (Optimized single-pass)
    const { monthlySummaries, grandTotals, allOverdueItems, missingDeadlinesCount } = processForecastingData(forecastingValues);
    Logger.log(`Processing complete. Found ${allOverdueItems.length} overdue items and ${missingDeadlinesCount} rows with missing deadlines.`);

    // 3. Populate Overdue Details
    populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders);

    // 4. Prepare and Populate Dashboard
    clearAndResizeSheet(dashboardSheet, CONFIG.DASHBOARD_LAYOUT.FIXED_ROW_COUNT, CONFIG.DASHBOARD_LAYOUT.HIDE_COL_END);
    setDashboardHeaders(dashboardSheet); // Updated to handle Year/Month split
    setDashboardHeaderNotes(dashboardSheet);

    const months = generateMonthList(CONFIG.DASHBOARD_DATES.START, CONFIG.DASHBOARD_DATES.END);
    const dataStartRow = 2;

    // Map processed data to the months list
    const dashboardData = months.map(month => {
        const monthKey = `${month.getFullYear()}-${month.getMonth()}`;
        return monthlySummaries.get(monthKey) || [0, 0, 0, 0]; // [total, upcoming, overdue, approved]
    });

    if (dashboardData.length > 0) {
      const numDataRows = dashboardData.length;
      const DL = CONFIG.DASHBOARD_LAYOUT;

      // Prepare data for batch writing
      const yearMonthData = months.map(date => [date.getFullYear(), date]);
      const overdueFormulas = dashboardData.map(row => [`=HYPERLINK("#gid=${overdueSheetGid}", ${row[2] || 0})`]);
      const otherData = dashboardData.map(row => [row[0], row[1], row[3]]); // [total, upcoming, approved]

      // Write data in batches - UPDATED FOR YEAR/MONTH SPLIT
      dashboardSheet.getRange(dataStartRow, DL.YEAR_COL, numDataRows, 2).setValues(yearMonthData);
      dashboardSheet.getRange(dataStartRow, DL.TOTAL_COL, numDataRows, 2).setValues(otherData.map(row => [row[0], row[1]]));
      dashboardSheet.getRange(dataStartRow, DL.OVERDUE_COL, numDataRows, 1).setFormulas(overdueFormulas);
      dashboardSheet.getRange(dataStartRow, DL.APPROVED_COL, numDataRows, 1).setValues(otherData.map(row => [row[2]]));

      // Write Grand Totals
      const [gtUpcoming, gtOverdue, gtTotal, gtApproved] = grandTotals;
      dashboardSheet.getRange(dataStartRow, DL.GT_UPCOMING_COL).setValue(gtUpcoming);
      dashboardSheet.getRange(dataStartRow, DL.GT_OVERDUE_COL).setFormula(`=HYPERLINK("#gid=${overdueSheetGid}", ${gtOverdue})`);
      dashboardSheet.getRange(dataStartRow, DL.GT_TOTAL_COL).setValue(gtTotal);
      dashboardSheet.getRange(dataStartRow, DL.GT_APPROVED_COL).setValue(gtApproved);

      // Write Missing Deadlines report
      const missingCell = dashboardSheet.getRange(DL.MISSING_DEADLINE_CELL);
      missingCell.setValue("Missing/Invalid Deadlines:");
      missingCell.offset(0, 1).setValue(missingDeadlinesCount).setNumberFormat("0").setFontWeight("bold");
      missingCell.setFontWeight("bold");

      // 5. Apply Formatting
      applyDashboardFormatting(dashboardSheet, numDataRows); // Updated for Year/Month split

      // 6. Generate Charts
      if (CONFIG.DASHBOARD_CHARTING.ENABLED) {
        createOrUpdateDashboardCharts(dashboardSheet, months, dashboardData);
        hideDataColumns(dashboardSheet);
      }
    }

    SpreadsheetApp.flush();
    const duration = (new Date().getTime() - scriptStartTime.getTime()) / 1000;
    Logger.log(`Dashboard update complete (Duration: ${duration.toFixed(2)} seconds).`);

  } catch (error) {
    Logger.log(`ERROR in updateDashboard: ${error.message}\nStack: ${error.stack}`);
    notifyError("Dashboard Update Failed", error, ss);
    ui.alert(`An error occurred updating the dashboard. Please check logs and the notification email.\nError: ${error.message}`);
  }
}

// =================================================================
// ==================== DATA PROCESSING ============================
// =================================================================

/**
 * Reads the necessary data from the 'Forecasting' sheet efficiently.
 * It determines the maximum column number required by any dashboard logic from `CONFIG`
 * and reads only up to that column in a single `getValues()` call, improving performance
 * on sheets with many unnecessary columns.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} forecastSheet The sheet object for the 'Forecasting' data source.
 * @returns {{forecastingValues: any[][], forecastingHeaders: string[]}|null} An object containing the 2D array
 *   of data values and a 1D array of header values, or null on failure.
 */
function readForecastingData(forecastSheet) {
  try {
    const dataRange = forecastSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const forecastingHeaders = numRows > 0 ? forecastSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0] : [];
    if (numRows <= 1) return { forecastingValues: [], forecastingHeaders };

    const colIndices = Object.values(CONFIG.FORECASTING_COLS);
    const lastColNumNeeded = Math.max(...colIndices);
    const numColsToRead = Math.min(lastColNumNeeded, dataRange.getNumColumns());

    const forecastingValues = forecastSheet.getRange(2, 1, numRows - 1, numColsToRead).getValues();
    return { forecastingValues, forecastingHeaders };
  } catch (e) {
    Logger.log(`ERROR reading data from ${forecastSheet.getName()}: ${e.message}`);
    return null;
  }
}

/**
 * Processes the raw forecasting data in a single, efficient pass to generate all dashboard metrics.
 * This is a key performance optimization. Instead of iterating over the data multiple times for
 * different calculations, it iterates once and calculates everything simultaneously.
 *
 * @param {any[][]} forecastingValues A 2D array of the raw data from the 'Forecasting' sheet.
 * @returns {{
 *   monthlySummaries: Map<string, number[]>,
 *   grandTotals: number[],
 *   allOverdueItems: any[][],
 *   missingDeadlinesCount: number
 * }} An object containing all the aggregated data required by the dashboard.
 */
function processForecastingData(forecastingValues) {
  const monthlySummaries = new Map();
  const allOverdueItems = [];
  let grandTotals = [0, 0, 0, 0]; // [upcoming, overdue, total, approved]
  let missingDeadlinesCount = 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const FC = CONFIG.FORECASTING_COLS;
  const deadlineIdx = FC.DEADLINE - 1;
  const progressIdx = FC.PROGRESS - 1;
  const permitsIdx = FC.PERMITS - 1;

  const { IN_PROGRESS, SCHEDULED, PERMIT_APPROVED } = CONFIG.STATUS_STRINGS;
  const inProgressLower = IN_PROGRESS.toLowerCase();
  const scheduledLower = SCHEDULED.toLowerCase();
  const approvedLower = PERMIT_APPROVED.toLowerCase();

  for (const row of forecastingValues) {
    const deadlineDate = parseAndNormalizeDate(row[deadlineIdx]);

    if (!deadlineDate) {
      missingDeadlinesCount++;
      continue;
    }

    const monthKey = `${deadlineDate.getFullYear()}-${deadlineDate.getMonth()}`;
    if (!monthlySummaries.has(monthKey)) {
      monthlySummaries.set(monthKey, [0, 0, 0, 0]); // [total, upcoming, overdue, approved]
    }
    const monthData = monthlySummaries.get(monthKey);

    monthData[0]++; // Total for month
    grandTotals[2]++; // GT Total

    const currentStatus = normalizeString(row[progressIdx]);
    const isActuallyInProgress = currentStatus === inProgressLower;
    const isActuallyScheduled = currentStatus === scheduledLower;

    if (isActuallyInProgress || isActuallyScheduled) {
      if (deadlineDate > today) {
        monthData[1]++; // Upcoming
        grandTotals[0]++;
      } else if (isActuallyInProgress && !isActuallyScheduled) {
        monthData[2]++; // Overdue
        grandTotals[1]++;
        allOverdueItems.push(row);
      }
    }

    if (normalizeString(row[permitsIdx]) === approvedLower) {
      monthData[3]++; // Approved
      grandTotals[3]++;
    }
  }

  return { monthlySummaries, grandTotals, allOverdueItems, missingDeadlinesCount };
}

// =================================================================
// ==================== PRESENTATION LOGIC =========================
// =================================================================

/**
 * Displays a placeholder message on the dashboard in the area where a chart would normally appear.
 * This provides clear user feedback when a chart cannot be generated due to a lack of data,
 * improving the user experience by explaining why a chart is missing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {number} anchorRow The 1-based starting row for the placeholder message.
 * @param {number} anchorCol The 1-based starting column for the placeholder message.
 * @param {string} message The text message to display in the placeholder.
 * @returns {void}
 */
function displayChartPlaceholder(sheet, anchorRow, anchorCol, message) {
  try {
    const placeholderRange = sheet.getRange(anchorRow + 5, anchorCol, 1, 4); // A small area in the middle of the chart space
    placeholderRange.merge();
    placeholderRange.setValue(message)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontStyle('italic')
      .setFontColor('#999999');
  } catch (e) {
    Logger.log(`Could not create chart placeholder: ${e.message}`);
  }
}

/**
 * Populates the 'Overdue Details' sheet with the full data for all overdue projects.
 * This function clears the sheet, resizes it to fit the data, writes the original headers,
 * and then writes the overdue project rows. It includes guards against errors that could
 * occur if the source 'Forecasting' sheet is empty or malformed.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} overdueDetailsSheet The sheet object for 'Overdue Details'.
 * @param {any[][]} allOverdueItems A 2D array of row data for projects determined to be overdue.
 * @param {string[]} forecastingHeaders An array of header strings from the source sheet.
 * @returns {void}
 */
function populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders) {
  try {
    // Guard against an empty/invalid source sheet.
    if (!forecastingHeaders || forecastingHeaders.length === 0) {
      overdueDetailsSheet.clear();
      overdueDetailsSheet.getRange(1, 1).setValue("Source 'Forecasting' sheet is empty or has no header row.");
      Logger.log("Skipped populating Overdue Details: 'Forecasting' sheet appears to be empty.");
      return;
    }

    const numRows = allOverdueItems.length;
    const numCols = allOverdueItems.length > 0 ? allOverdueItems[0].length : forecastingHeaders.length;

    overdueDetailsSheet.clear();
    // Smart resizing to keep the sheet clean.
    if (overdueDetailsSheet.getMaxRows() > 1) {
      overdueDetailsSheet.deleteRows(2, overdueDetailsSheet.getMaxRows() - 1);
    }
    if (overdueDetailsSheet.getMaxColumns() > numCols) {
        overdueDetailsSheet.deleteColumns(numCols + 1, overdueDetailsSheet.getMaxColumns() - numCols);
    }

    const headersToWrite = forecastingHeaders.slice(0, numCols);
    overdueDetailsSheet.getRange(1, 1, 1, headersToWrite.length).setValues([headersToWrite]).setFontWeight("bold");

    if (numRows > 0) {
      if (overdueDetailsSheet.getMaxRows() < numRows + 1) {
        overdueDetailsSheet.insertRowsAfter(1, numRows);
      }
      overdueDetailsSheet.getRange(2, 1, numRows, numCols).setValues(allOverdueItems);
    }
    Logger.log(`Populated Overdue Details sheet with ${numRows} items.`);
  } catch (e) {
    Logger.log(`ERROR in populateOverdueDetailsSheet: ${e.message}`);
  }
}

/**
 * Sets the static main headers for the dashboard's summary table. This function writes the
 * header titles for both the monthly data and the grand total columns and applies standard
 * header formatting (background color, font color, etc.).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @returns {void}
 */
function setDashboardHeaders(sheet) {
  const DL = CONFIG.DASHBOARD_LAYOUT;
  const DF = CONFIG.DASHBOARD_FORMATTING;

  const headers = [
    "Year", "Month", "Total Projects", "Upcoming", "Overdue", "Approved",
    "GT Upcoming", "GT Overdue", "GT Total", "GT Approved"
  ];
  const headerRanges = [
    sheet.getRange(1, DL.YEAR_COL, 1, 6), // Year to Approved
    sheet.getRange(1, DL.GT_UPCOMING_COL, 1, 4) // Grand Totals
  ];

  headerRanges[0].setValues([headers.slice(0, 6)]);
  headerRanges[1].setValues([headers.slice(6, 10)]);

  headerRanges.forEach(range => {
    range.setBackground(DF.HEADER_BACKGROUND)
         .setFontColor(DF.HEADER_FONT_COLOR)
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
  });
}

/**
 * Sets explanatory notes on the dashboard header cells. These notes appear when a user
 * hovers over a header, providing helpful context about what each metric represents.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @returns {void}
 */
function setDashboardHeaderNotes(sheet) {
    const DL = CONFIG.DASHBOARD_LAYOUT;
    sheet.getRange(1, DL.TOTAL_COL).setNote("Total projects with a deadline in this month.");
    sheet.getRange(1, DL.UPCOMING_COL).setNote("Projects 'In Progress' or 'Scheduled' with a deadline in the future.");
    sheet.getRange(1, DL.OVERDUE_COL).setNote("Projects 'In Progress' with a deadline in the past. Click number to see details.");
    sheet.getRange(1, DL.APPROVED_COL).setNote("Projects with 'Permits' status set to 'approved'.");
    sheet.getRange(1, DL.GT_TOTAL_COL).setNote("Grand total of all projects with a valid deadline.");
}

/**
 * Applies all visual formatting to the dashboard data rows. This includes row banding
 * (alternating colors), number formatting for dates and counts, and cell borders to create
 * a clean, readable report.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {number} numDataRows The number of data rows to which formatting should be applied.
 * @returns {void}
 */
function applyDashboardFormatting(sheet, numDataRows) {
  const DL = CONFIG.DASHBOARD_LAYOUT;
  const DF = CONFIG.DASHBOARD_FORMATTING;
  const dataRange = sheet.getRange(2, DL.YEAR_COL, numDataRows, 6);

  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
           .setHeaderRowColor(null)
           .setFirstRowColor(DF.BANDING_COLOR_ODD)
           .setSecondRowColor(DF.BANDING_COLOR_EVEN);

  sheet.getRange(2, 1, numDataRows, DL.GT_APPROVED_COL).setHorizontalAlignment("center");

  sheet.getRange(2, DL.YEAR_COL, numDataRows, 1).setNumberFormat("0000");
  sheet.getRange(2, DL.MONTH_COL, numDataRows, 1).setNumberFormat(DF.MONTH_FORMAT);

  sheet.getRange(2, DL.TOTAL_COL, numDataRows, 4).setNumberFormat(DF.COUNT_FORMAT);
  sheet.getRange(2, DL.GT_UPCOMING_COL, 1, 4).setNumberFormat(DF.COUNT_FORMAT);

  sheet.getRange(1, 1, numDataRows + 1, DL.GT_APPROVED_COL).setBorder(true, true, true, true, true, true, DF.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
}

/**
 * Hides the temporary data columns used for chart generation. This keeps the user-facing
 * sheet clean, as these columns are an implementation detail and not meant for direct viewing.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @returns {void}
 */
function hideDataColumns(sheet) {
    const DL = CONFIG.DASHBOARD_LAYOUT;
    if (sheet.getMaxColumns() < DL.HIDE_COL_START) {
        Logger.log(`Skipping hideDataColumns: Sheet only has ${sheet.getMaxColumns()} columns, which is less than the required ${DL.HIDE_COL_START}.`);
        return;
    }
    sheet.hideColumns(DL.HIDE_COL_START, DL.HIDE_COL_END - DL.HIDE_COL_START + 1);
}

/**
 * Creates or updates the dashboard charts. This function first removes any existing charts
 * to ensure a clean slate. It then creates a temporary, hidden sheet to stage the data for
 * each chart, which is a robust method for building charts programmatically. It generates
 * two charts: one for past months and one for upcoming months, and includes on-sheet
 * placeholder feedback if no data is available for a chart.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet where charts will be inserted.
 * @param {Date[]} months An array of Date objects representing all months in the dashboard's range.
 * @param {any[][]} dashboardData A 2D array of the processed monthly summary data.
 * @returns {void}
 */
function createOrUpdateDashboardCharts(sheet, months, dashboardData) {
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheetName = "TempChartData_Dashboard_v4"; // Versioned to avoid conflicts
    let tempSheet = ss.getSheetByName(tempSheetName);
    if (tempSheet) ss.deleteSheet(tempSheet);
    tempSheet = ss.insertSheet(tempSheetName).hideSheet();

    try {
        const DC = CONFIG.DASHBOARD_CHARTING;
        const DL = CONFIG.DASHBOARD_LAYOUT;
        const DF = CONFIG.DASHBOARD_FORMATTING.CHART_COLORS;
        const timeZone = ss.getSpreadsheetTimeZone();

        const createChart = (title, data, headers, colors, anchorRow) => {
            tempSheet.clearContents();
            tempSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
            tempSheet.getRange(2, 1, data.length, headers.length).setValues(data);
            const dataRange = tempSheet.getRange(1, 1, data.length + 1, headers.length);
            const chart = sheet.newChart().asColumnChart()
                .addRange(dataRange)
                .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
                .setNumHeaders(1)
                .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
                .setOption('title', title)
                .setOption('width', DC.CHART_WIDTH)
                .setOption('height', DC.CHART_HEIGHT)
                .setOption('colors', colors)
                .setOption('legend', { position: 'top' })
                .setPosition(anchorRow, DL.CHART_ANCHOR_COL, 0, 0)
                .build();
            sheet.insertChart(chart);
        };

        const today = new Date();
        today.setDate(1);
        today.setHours(0, 0, 0, 0);

        const pastStartDate = new Date(today.getTime());
        pastStartDate.setMonth(pastStartDate.getMonth() - DC.PAST_MONTHS_COUNT);

        const upcomingEndDate = new Date(today.getTime());
        upcomingEndDate.setMonth(upcomingEndDate.getMonth() + DC.UPCOMING_MONTHS_COUNT);

        const combinedData = months.map((month, i) => ({
            month: month,
            monthLabel: Utilities.formatDate(month, timeZone, "MMM yyyy"),
            overdue: dashboardData[i][2],
            upcoming: dashboardData[i][1],
            total: dashboardData[i][0]
        }));

        const pastData = combinedData.filter(d => d.month >= pastStartDate && d.month < today)
                                     .map(d => [d.monthLabel, d.overdue, d.total]);

        const upcomingData = combinedData.filter(d => d.month >= today && d.month < upcomingEndDate)
                                         .map(d => [d.monthLabel, d.upcoming, d.total]);

        if (pastData.length > 0) {
            createChart(`Past ${pastData.length} Months: Overdue vs. Total`, pastData, ['Month', 'Overdue', 'Total'], [DF.overdue, DF.total], DL.CHART_START_ROW);
        } else {
            const message = `No project data found for the past ${DC.PAST_MONTHS_COUNT} months.`;
            displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, message);
            Logger.log(`Skipping 'Past Months' chart: No data in the specified date range.`);
        }

        if (upcomingData.length > 0) {
            createChart(`Next ${upcomingData.length} Months: Upcoming vs. Total`, upcomingData, ['Month', 'Upcoming', 'Total'], [DF.upcoming, DF.total], DL.CHART_START_ROW + DC.ROW_SPACING);
        } else {
            const message = `No project data found for the next ${DC.UPCOMING_MONTHS_COUNT} months.`;
            displayChartPlaceholder(sheet, DL.CHART_START_ROW + DC.ROW_SPACING, DL.CHART_ANCHOR_COL, message);
            Logger.log(`Skipping 'Upcoming Months' chart: No data in the specified date range.`);
        }

    } catch (e) {
        Logger.log(`A critical error occurred in createOrUpdateDashboardCharts: ${e.message}\n${e.stack}`);
    } finally {
        // Cleanup the temporary sheet
        if (ss.getSheetByName(tempSheetName)) {
            ss.deleteSheet(tempSheet);
        }
    }
}

/**
 * Generates a continuous list of Date objects, one for the first day of each month
 * between a specified start and end date.
 *
 * @param {Date} startDate The first month to include in the list.
 * @param {Date} endDate The last month to include in the list.
 * @returns {Date[]} An array of Date objects.
 */
function generateMonthList(startDate, endDate) {
    const months = [];
    let currentDate = new Date(startDate.getTime());
    while (currentDate <= endDate) {
        months.push(new Date(currentDate));
        currentDate.setMonth(currentDate.getMonth() + 1);
    }
    return months;
}