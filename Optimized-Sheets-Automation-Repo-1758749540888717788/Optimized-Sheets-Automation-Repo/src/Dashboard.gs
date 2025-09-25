/**
 * @OnlyCurrentDoc
 * Dashboard.gs
 * Logic for generating the dashboard report, charts, and overdue details.
 * Utilizes efficient single-pass data processing.
 */

/**
 * Main orchestrator function to update the Dashboard sheet.
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
    clearAndResizeSheet(dashboardSheet, CONFIG.DASHBOARD_LAYOUT.FIXED_ROW_COUNT);
    setDashboardHeaders(dashboardSheet);
    setDashboardHeaderNotes(dashboardSheet);

    const months = generateMonthList(CONFIG.DASHBOARD_DATES.START, CONFIG.DASHBOARD_DATES.END);
    const dataStartRow = 2;

    // Map processed data to the months list
    const dashboardData = months.map(month => {
        // Use standard JS month indexing (0-11) for the map key
        const monthKey = `${month.getFullYear()}-${month.getMonth()}`;
        // [total, upcoming, overdue, approved]
        return monthlySummaries.get(monthKey) || [0, 0, 0, 0]; 
    });

    if (dashboardData.length > 0) {
      const numDataRows = dashboardData.length;
      // Ensure enough rows exist
      if (dashboardSheet.getMaxRows() < dataStartRow + numDataRows - 1) {
           dashboardSheet.insertRowsAfter(dashboardSheet.getMaxRows(), (dataStartRow + numDataRows - 1) - dashboardSheet.getMaxRows());
      }

      const DL = CONFIG.DASHBOARD_LAYOUT;

      // Prepare data for batch writing
      const overdueFormulas = dashboardData.map(row => [`=HYPERLINK("#gid=${overdueSheetGid}", ${row[2] || 0})`]);
      // Extract [total, upcoming, approved]
      const otherData = dashboardData.map(row => [row[0], row[1], row[3]]);

      // Write data in batches
      dashboardSheet.getRange(dataStartRow, DL.MONTH_COL, numDataRows, 1).setValues(months.map(date => [date]));
      // Write Total and Upcoming
      dashboardSheet.getRange(dataStartRow, DL.TOTAL_COL, numDataRows, 2).setValues(otherData.map(row => [row[0], row[1]]));
      // Write Overdue (with formulas/links)
      dashboardSheet.getRange(dataStartRow, DL.OVERDUE_COL, numDataRows, 1).setFormulas(overdueFormulas);
      // Write Approved
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
      applyDashboardFormatting(dashboardSheet, numDataRows, dashboardData);

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
    // Use the centralized error notification system
    notifyError("Dashboard Update Failed", error, ss);
    ui.alert(`An error occurred updating the dashboard. Please check logs and the notification email.\nError: ${error.message}`);
  }
}

// =================================================================
// ==================== DATA PROCESSING ============================
// =================================================================

/** Reads necessary data from the Forecasting sheet. */
function readForecastingData(forecastSheet) {
  try {
    const dataRange = forecastSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const forecastingHeaders = numRows > 0 ? forecastSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0] : [];
    if (numRows <= 1) return { forecastingValues: [], forecastingHeaders };
    
    // Determine columns needed based on CONFIG (1-based indices)
    const colIndices = Object.values(CONFIG.FORECASTING_COLS);
    const lastColNumNeeded = Math.max(...colIndices);
    const numColsToRead = Math.min(lastColNumNeeded, dataRange.getNumColumns());
    
    // Read data in a single batch starting from row 2
    const forecastingValues = forecastSheet.getRange(2, 1, numRows - 1, numColsToRead).getValues();
    return { forecastingValues, forecastingHeaders };
  } catch (e) {
    Logger.log(`ERROR reading data from ${forecastSheet.getName()}: ${e.message}`);
    return null;
  }
}

/**
 * OPTIMIZED: Processes forecasting data in a single pass.
 * @param {Array[]} forecastingValues 2D array of data.
 * @returns {object} Summarized data.
 */
function processForecastingData(forecastingValues) {
  const monthlySummaries = new Map();
  const allOverdueItems = [];
  // [upcoming, overdue, total, approved]
  let grandTotals = [0, 0, 0, 0]; 
  let missingDeadlinesCount = 0;

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  // Get column indices and convert to 0-based for array access
  const FC = CONFIG.FORECASTING_COLS;
  const deadlineIdx = FC.DEADLINE - 1;
  const progressIdx = FC.PROGRESS - 1;
  const permitsIdx = FC.PERMITS - 1;

  // Prepare status strings for comparison
  const { IN_PROGRESS, SCHEDULED, PERMIT_APPROVED } = CONFIG.STATUS_STRINGS;
  const inProgressLower = IN_PROGRESS.toLowerCase();
  const scheduledLower = SCHEDULED.toLowerCase();
  const approvedLower = PERMIT_APPROVED.toLowerCase();

  // Single pass iteration
  for (const row of forecastingValues) {
    const deadlineDate = parseAndNormalizeDate(row[deadlineIdx]);

    if (!deadlineDate) {
      missingDeadlinesCount++;
      continue; // Skip row if deadline is invalid
    }

    // Use standard JS month indexing (0-11) for the key
    const monthKey = `${deadlineDate.getFullYear()}-${deadlineDate.getMonth()}`;
    if (!monthlySummaries.has(monthKey)) {
      // [total, upcoming, overdue, approved]
      monthlySummaries.set(monthKey, [0, 0, 0, 0]); 
    }
    const monthData = monthlySummaries.get(monthKey);

    // --- Calculations ---
    monthData[0]++; // Total for month
    grandTotals[2]++; // GT Total

    const currentStatus = normalizeString(row[progressIdx]);
    const isActuallyInProgress = currentStatus === inProgressLower;
    const isActuallyScheduled = currentStatus === scheduledLower;

    if (isActuallyInProgress || isActuallyScheduled) {
      if (deadlineDate > today) {
        // Upcoming
        monthData[1]++;
        grandTotals[0]++;
      } else if (isActuallyInProgress && !isActuallyScheduled) {
        // Overdue criteria: In Progress AND Deadline <= Today AND NOT Scheduled
        monthData[2]++;
        grandTotals[1]++;
        allOverdueItems.push(row); // Add full row to detailed list
      }
    }

    if (normalizeString(row[permitsIdx]) === approvedLower) {
      // Approved
      monthData[3]++;
      grandTotals[3]++;
    }
  }

  return { monthlySummaries, grandTotals, allOverdueItems, missingDeadlinesCount };
}

// =================================================================
// ==================== PRESENTATION LOGIC =========================
// =================================================================

/**
 * Populates the Overdue Details sheet with the provided data.
 * @param {Sheet} overdueDetailsSheet The destination sheet.
 * @param {Array[]} allOverdueItems The data rows to write.
 * @param {Array<string>} forecastingHeaders The header row.
 */
function populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders) {
  try {
    const numRows = allOverdueItems.length;
    const numCols = forecastingHeaders.length;

    // Clear previous data and formatting
    overdueDetailsSheet.clear();
    if (overdueDetailsSheet.getMaxRows() > 1) {
      overdueDetailsSheet.deleteRows(2, overdueDetailsSheet.getMaxRows() - 1);
    }
    if (overdueDetailsSheet.getMaxColumns() > numCols) {
        overdueDetailsSheet.deleteColumns(numCols + 1, overdueDetailsSheet.getMaxColumns() - numCols);
    }

    // Write headers
    overdueDetailsSheet.getRange(1, 1, 1, numCols).setValues([forecastingHeaders]).setFontWeight("bold");

    if (numRows > 0) {
      // Ensure enough rows exist for the data
      if (overdueDetailsSheet.getMaxRows() < numRows + 1) {
        overdueDetailsSheet.insertRowsAfter(1, numRows);
      }
      // Write data
      overdueDetailsSheet.getRange(2, 1, numRows, numCols).setValues(allOverdueItems);
    }

    Logger.log(`Populated Overdue Details sheet with ${numRows} items.`);
  } catch (e) {
    Logger.log(`ERROR in populateOverdueDetailsSheet: ${e.message}`);
    // No need to throw here, as dashboard can still partially function
  }
}

/** Sets the main headers for the dashboard. */
function setDashboardHeaders(sheet) {
  const DL = CONFIG.DASHBOARD_LAYOUT;
  const DF = CONFIG.DASHBOARD_FORMATTING;

  const headers = [
    "Month", "Total Projects", "Upcoming", "Overdue", "Approved",
    "GT Upcoming", "GT Overdue", "GT Total", "GT Approved"
  ];
  const headerRanges = [
    sheet.getRange(1, DL.MONTH_COL, 1, 5),
    sheet.getRange(1, DL.GT_UPCOMING_COL, 1, 4)
  ];

  headerRanges[0].setValues([headers.slice(0, 5)]);
  headerRanges[1].setValues([headers.slice(5, 9)]);

  // Apply formatting to all headers
  headerRanges.forEach(range => {
    range.setBackground(DF.HEADER_BACKGROUND)
         .setFontColor(DF.HEADER_FONT_COLOR)
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
  });
}

/** Sets explanatory notes for dashboard headers. */
function setDashboardHeaderNotes(sheet) {
    const DL = CONFIG.DASHBOARD_LAYOUT;
    sheet.getRange(1, DL.TOTAL_COL).setNote("Total projects with a deadline in this month.");
    sheet.getRange(1, DL.UPCOMING_COL).setNote("Projects 'In Progress' or 'Scheduled' with a deadline in the future.");
    sheet.getRange(1, DL.OVERDUE_COL).setNote("Projects 'In Progress' with a deadline in the past. Click number to see details.");
    sheet.getRange(1, DL.APPROVED_COL).setNote("Projects with 'Permits' status set to 'approved'.");
    sheet.getRange(1, DL.GT_TOTAL_COL).setNote("Grand total of all projects with a valid deadline.");
}

/** Applies conditional formatting and banding to the dashboard. */
function applyDashboardFormatting(sheet, numDataRows) {
  const DL = CONFIG.DASHBOARD_LAYOUT;
  const DF = CONFIG.DASHBOARD_FORMATTING;
  const dataRange = sheet.getRange(2, 1, numDataRows, 5);

  // Apply and configure banding
  const banding = dataRange.applyRowBanding(); // Apply default banding theme
  banding.setHeaderRow(null); // No header color from banding
  banding.setFirstRowColor(DF.BANDING_COLOR_ODD);
  banding.setSecondRowColor(DF.BANDING_COLOR_EVEN);

  // Center align all data
  sheet.getRange(2, 1, numDataRows, DL.GT_APPROVED_COL).setHorizontalAlignment("center");

  // Format month column
  sheet.getRange(2, DL.MONTH_COL, numDataRows, 1).setNumberFormat(DF.MONTH_FORMAT);

  // Format count columns
  sheet.getRange(2, DL.TOTAL_COL, numDataRows, 4).setNumberFormat(DF.COUNT_FORMAT);
  sheet.getRange(2, DL.GT_UPCOMING_COL, 1, 4).setNumberFormat(DF.COUNT_FORMAT);

  // Add borders for clarity
  sheet.getRange(1, 1, numDataRows + 1, DL.GT_APPROVED_COL).setBorder(true, true, true, true, true, true, DF.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
}

/** Hides temporary data columns used for charting. */
function hideDataColumns(sheet) {
    const DL = CONFIG.DASHBOARD_LAYOUT;
    sheet.hideColumns(DL.HIDE_COL_START, DL.HIDE_COL_END - DL.HIDE_COL_START + 1);
}

/** Creates or updates dashboard charts, now with robust temp sheet handling. */
function createOrUpdateDashboardCharts(sheet, months, dashboardData) {
    // Remove existing charts to prevent duplicates
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));

    const DC = CONFIG.DASHBOARD_CHARTING;
    const DL = CONFIG.DASHBOARD_LAYOUT;
    const DF = CONFIG.DASHBOARD_FORMATTING.CHART_COLORS;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheetName = "TempChartData_Dashboard"; // Use a more specific name

    // Prepare data for charts
    const chartData = months.map((month, i) => [
        month,
        dashboardData[i][2], // Overdue
        dashboardData[i][1], // Upcoming
        dashboardData[i][0]  // Total
    ]);

    // --- Robust Temp Sheet Handling ---
    let tempSheet = ss.getSheetByName(tempSheetName);
    if (tempSheet) {
      ss.deleteSheet(tempSheet); // Delete if it exists from a previous failed run
    }
    tempSheet = ss.insertSheet(tempSheetName);
    // --- End Robust Handling ---

    try {
        tempSheet.getRange(1, 1, chartData.length, 4).setValues(chartData);

        const createChart = (title, dataRange, seriesColors) => {
            const builder = sheet.newChart()
                .setChartType(Charts.ChartType.COLUMN)
                .addRange(dataRange)
                .setOption('title', title)
                .setOption('width', DC.CHART_WIDTH)
                .setOption('height', DC.CHART_HEIGHT)
                .setOption('colors', seriesColors)
                .setOption('legend', { position: 'top' })
                .asColumnChart();
            return builder.build();
        };

        // Past 3 Months Trend
        const pastDataRange = tempSheet.getRange(1, 1, DC.PAST_MONTHS_COUNT, 4);
        const pastChart = createChart('Past 3 Months: Overdue vs. Total', pastDataRange, [DF.overdue, DF.total]);

        // Upcoming 6 Months Trend
        const upcomingDataRange = tempSheet.getRange(DC.PAST_MONTHS_COUNT + 1, 1, DC.UPCOMING_MONTHS_COUNT, 4);
        const upcomingChart = createChart('Next 6 Months: Upcoming vs. Total', upcomingDataRange, [DF.upcoming, DF.total]);

        sheet.insertChart(pastChart);
        sheet.insertChart(upcomingChart);

    } finally {
        // --- Cleanup ---
        // Ensure the temporary sheet is always deleted, even if chart creation fails.
        if (tempSheet) {
            ss.deleteSheet(tempSheet);
        }
        // --- End Cleanup ---
    }
}

/** Generates a list of months between a start and end date. */
function generateMonthList(startDate, endDate) {
    const months = [];
    let currentDate = new Date(startDate.getTime());
    while (currentDate <= endDate) {
        months.push(new Date(currentDate));
        currentDate.setMonth(currentDate.getMonth() + 1);
    }
    return months;
}