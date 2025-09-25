/**
 * @OnlyCurrentDoc
 * Dashboard.gs
 * Logic for generating the dashboard report, charts, and overdue details.
 * Utilizes efficient single-pass data processing.
 */

/**
 * Main orchestrator function to generate or update the entire Dashboard.
 * This function, typically triggered from the custom menu, follows a sequence:
 * 1. Reads raw data from the 'Forecasting' sheet.
 * 2. Processes the data in a single pass to create summaries and identify overdue items.
 * 3. Populates the 'Overdue Details' sheet with a drill-down list of overdue projects.
 * 4. Clears, resizes, and populates the main 'Dashboard' sheet with monthly summaries and grand totals.
 * 5. Applies all formatting, including colors, number formats, and borders.
 * 6. Generates and embeds summary charts directly into the dashboard.
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
    populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems);
    
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
      applyDashboardFormatting(dashboardSheet, numDataRows);

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

/**
 * Reads the necessary data from the 'Forecasting' sheet efficiently.
 * It determines the required columns from `CONFIG` to avoid reading the entire sheet,
 * and returns both the data values and the header row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} forecastSheet The 'Forecasting' sheet object.
 * @returns {{forecastingValues: Array<Array<*>>, forecastingHeaders: Array<string>}|null} An object containing the 2D data array and the 1D header array, or null on failure.
 */
function readForecastingData(forecastSheet) {
  try {
    const dataRange = forecastSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const forecastingHeaders = numRows > 0 ? forecastSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0] : [];
    if (numRows <= 1) {
      Logger.log("'Forecasting' sheet is empty or contains only headers. No data to process.");
      return { forecastingValues: [], forecastingHeaders };
    }
    
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
 * Processes the raw forecasting data in a single, efficient pass to generate all necessary summaries.
 * This function iterates through each row, categorizes it based on its deadline and status,
 * and aggregates the results into monthly summaries, grand totals, a list of overdue items,
 * and a count of rows with missing or invalid deadlines.
 *
 * @param {Array<Array<*>>} forecastingValues A 2D array of the data rows from the 'Forecasting' sheet.
 * @returns {{monthlySummaries: Map<string, number[]>, grandTotals: number[], allOverdueItems: Array<Array<*>>, missingDeadlinesCount: number}} An object containing the processed data.
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
 * Clears and populates the 'Overdue Details' sheet with a focused subset of data for all overdue projects.
 * This provides a cleaner "drill-down" view by only showing columns specified in `CONFIG.OVERDUE_DETAILS_DISPLAY_KEYS`.
 * This approach prevents column mismatch errors by building the data array with an explicit structure.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} overdueDetailsSheet The destination sheet object.
 * @param {Array<Array<*>>} allOverdueItems A 2D array of the full data rows for overdue projects.
 */
function populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems) {
  try {
    // 1. Get the desired column keys and map them to 0-based indices
    const displayKeys = CONFIG.OVERDUE_DETAILS_DISPLAY_KEYS;
    const colIndices = displayKeys.map(key => CONFIG.FORECASTING_COLS[key] - 1);

    // 2. Create the new, focused headers
    const newHeaders = displayKeys.map(key => key.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase()));

    // 3. Build the new data array with only the required columns
    const overdueDataSubset = allOverdueItems.map(fullRow =>
      colIndices.map(colIdx => fullRow[colIdx] !== undefined ? fullRow[colIdx] : "")
    );

    const numRows = overdueDataSubset.length;
    const numCols = newHeaders.length;

    // 4. Efficiently clear and resize the sheet
    overdueDetailsSheet.clear();
    if (overdueDetailsSheet.getMaxRows() > 1) {
      overdueDetailsSheet.deleteRows(2, overdueDetailsSheet.getMaxRows() - 1);
    }
     if (overdueDetailsSheet.getMaxColumns() > numCols) {
        overdueDetailsSheet.deleteColumns(numCols + 1, overdueDetailsSheet.getMaxColumns() - numCols);
    }

    // 5. Write headers and data in two batches
    overdueDetailsSheet.getRange(1, 1, 1, numCols).setValues([newHeaders]).setFontWeight("bold");

    if (numRows > 0) {
      if (overdueDetailsSheet.getMaxRows() < numRows + 1) {
        overdueDetailsSheet.insertRowsAfter(1, numRows);
      }
      overdueDetailsSheet.getRange(2, 1, numRows, numCols).setValues(overdueDataSubset);
    }

    Logger.log(`Successfully populated 'Overdue Details' sheet with ${numRows} items and ${numCols} columns.`);

  } catch (e) {
    const errorMessage = `Failed to populate 'Overdue Details' sheet. This can happen if 'OVERDUE_DETAILS_DISPLAY_KEYS' in CONFIG contains an invalid key. Error: ${e.message}`;
    Logger.log(`ERROR in populateOverdueDetailsSheet: ${errorMessage}\nStack: ${e.stack}`);
    // Do not throw, as the main dashboard can still be generated.
  }
}

/**
 * Sets the static main headers for the dashboard summary table.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The 'Dashboard' sheet object.
 */
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

/**
 * Sets explanatory notes on the dashboard header cells to provide context for each metric.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The 'Dashboard' sheet object.
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
 * Applies all visual formatting to the dashboard data range, including row banding,
 * text alignment, number formatting, and borders.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The 'Dashboard' sheet object.
 * @param {number} numDataRows The number of data rows (months) being displayed.
 */
function applyDashboardFormatting(sheet, numDataRows) {
  const DL = CONFIG.DASHBOARD_LAYOUT;
  const DF = CONFIG.DASHBOARD_FORMATTING;
  const dataRange = sheet.getRange(2, 1, numDataRows, 5);

  // Apply banding
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
           .setHeaderRowColor(null) // No header color from banding
           .setFirstRowColor(DF.BANDING_COLOR_ODD)
           .setSecondRowColor(DF.BANDING_COLOR_EVEN);

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

/**
 * Hides the temporary data columns that are used as a source for the dashboard charts.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The 'Dashboard' sheet object.
 */
function hideDataColumns(sheet) {
    const DL = CONFIG.DASHBOARD_LAYOUT;
    sheet.hideColumns(DL.HIDE_COL_START, DL.HIDE_COL_END - DL.HIDE_COL_START + 1);
}

/**
 * Creates or updates the charts on the dashboard.
 * This robust function first removes any existing charts to ensure a clean slate. It then creates
 * a temporary, hidden sheet to stage the data for each chart. It filters the main dashboard data
 * into "past" and "upcoming" periods. If data is missing for a chart, it provides clear feedback
 * on the sheet itself instead of failing silently.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The 'Dashboard' sheet object.
 * @param {Date[]} months The full list of month Date objects for the dashboard's time range.
 * @param {Array<Array<number>>} dashboardData The 2D array of summary data corresponding to the months list.
 */
function createOrUpdateDashboardCharts(sheet, months, dashboardData) {
    // Clean up previous state
    sheet.getCharts().forEach(chart => sheet.removeChart(chart));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tempSheetName = "TempChartData_Dashboard_v3";
    let tempSheet = ss.getSheetByName(tempSheetName);
    if (tempSheet) ss.deleteSheet(tempSheet);
    tempSheet = ss.insertSheet(tempSheetName).hideSheet();

    try {
        const DC = CONFIG.DASHBOARD_CHARTING;
        const DL = CONFIG.DASHBOARD_LAYOUT;
        const DF = CONFIG.DASHBOARD_FORMATTING.CHART_COLORS;
        const timeZone = ss.getSpreadsheetTimeZone();

        // --- Helper to write feedback message on the dashboard ---
        const setChartPlaceholder = (anchorRow, title, message) => {
            const range = sheet.getRange(anchorRow, DL.CHART_ANCHOR_COL, 2, 4);
            range.clearContent();
            range.merge();
            range.setValue(`${title}\n\n${message}`)
                 .setVerticalAlignment("middle")
                 .setHorizontalAlignment("center")
                 .setFontColor("#9E9E9E") // Grey color for placeholder
                 .setFontStyle("italic")
                 .setBackground("#F5F5F5")
                 .setBorder(true, true, true, true, null, null, "#E0E0E0", SpreadsheetApp.BorderStyle.DASHED);
        };

        // --- Generic Chart Creation Function ---
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

        // --- Date Calculations ---
        const today = new Date();
        today.setDate(1); // Normalize to the first of the month
        today.setHours(0, 0, 0, 0);

        const pastStartDate = new Date(today.getTime());
        pastStartDate.setMonth(pastStartDate.getMonth() - DC.PAST_MONTHS_COUNT);

        const upcomingEndDate = new Date(today.getTime());
        upcomingEndDate.setMonth(upcomingEndDate.getMonth() + DC.UPCOMING_MONTHS_COUNT);

        const formatDateForLog = (date) => Utilities.formatDate(date, timeZone, "yyyy-MM-dd");

        // --- Data Filtering using Date Objects ---
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

        // --- Chart Generation ---
        const pastChartTitle = `Past ${DC.PAST_MONTHS_COUNT} Months: Overdue vs. Total`;
        if (pastData.length > 0) {
            createChart(
                pastChartTitle,
                pastData,
                ['Month', 'Overdue', 'Total'],
                [DF.overdue, DF.total],
                DL.CHART_START_ROW
            );
        } else {
            const logMsg = `Skipping 'Past Months' chart: No projects found with deadlines between ${formatDateForLog(pastStartDate)} and ${formatDateForLog(today)}.`;
            Logger.log(logMsg);
            setChartPlaceholder(DL.CHART_START_ROW, pastChartTitle, "Not enough recent data to generate this chart.");
        }

        const upcomingChartTitle = `Next ${DC.UPCOMING_MONTHS_COUNT} Months: Upcoming vs. Total`;
        if (upcomingData.length > 0) {
            createChart(
                upcomingChartTitle,
                upcomingData,
                ['Month', 'Upcoming', 'Total'],
                [DF.upcoming, DF.total],
                DL.CHART_START_ROW + DC.ROW_SPACING
            );
        } else {
            const logMsg = `Skipping 'Upcoming Months' chart: No projects found with deadlines between ${formatDateForLog(today)} and ${formatDateForLog(upcomingEndDate)}.`;
            Logger.log(logMsg);
            setChartPlaceholder(DL.CHART_START_ROW + DC.ROW_SPACING, upcomingChartTitle, "No upcoming project data to generate this chart.");
        }

    } catch (e) {
        const errorMessage = `A critical error occurred while creating dashboard charts. This can be due to issues with chart data ranges or invalid chart options in CONFIG. Error: ${e.message}`;
        Logger.log(`${errorMessage}\nStack: ${e.stack}`);
    } finally {
        if (ss.getSheetByName(tempSheetName)) {
            ss.deleteSheet(tempSheet);
        }
    }
}

/**
 * Generates an array of Date objects, representing the first day of each month
 * between a specified start and end date (inclusive).
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