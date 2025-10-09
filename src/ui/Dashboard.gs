/**
 * @OnlyCurrentDoc
 *
 * Dashboard.gs
 *
 * This script generates the main dashboard, which includes data processing, table rendering, charts,
 * and an overdue items drill-down. The design focuses on correctness, idempotence, performance, and a clean user experience.
 *
 * @version 1.6.0
 * @release 2025-10-08
 */

/**
 * The main orchestrator function to generate or update the Dashboard sheet.
 * It follows a comprehensive sequence:
 * 1. Reads data from the 'Forecasting' sheet.
 * 2. Processes the data to calculate monthly summaries and identify overdue items.
 * 3. Renders the main data table, including grand totals and hover notes for overdue items.
 * 4. Creates or updates summary charts if enabled in the configuration.
 * 5. Hides temporary data columns to maintain a clean UI.
 * It includes robust error handling and logging throughout the process.
 *
 * @returns {void} This function does not return a value.
 */
function updateDashboard() {
  const ui = SpreadsheetApp.getUi();
  const scriptStartTime = new Date();
  Logger.log('Dashboard update initiated at ' + scriptStartTime.toLocaleString());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG;

  try {
    const { FORECASTING, DASHBOARD } = config.SHEETS;
    const forecastSheet = ss.getSheetByName(FORECASTING);
    if (!forecastSheet) throw new Error('Sheet "' + FORECASTING + '" not found.');

    const dashboardSheet = getOrCreateSheet(ss, DASHBOARD);

    // Delete the "Overdue Details" sheet as it's no longer needed.
    const overdueSheetToDelete = ss.getSheetByName("Overdue Details");
    if (overdueSheetToDelete) {
      ss.deleteSheet(overdueSheetToDelete);
      Logger.log('Successfully deleted "Overdue Details" sheet.');
    }

    const read = readForecastingData(forecastSheet, config);
    if (!read || !read.forecastingValues) throw new Error('Failed to read data from ' + FORECASTING + '.');
    const { forecastingValues } = read;

    const processed = processDashboardData(forecastingValues, config);
    const { monthlySummaries, allOverdueItems, missingDeadlinesCount } = processed;
    Logger.log('Processing complete. Found ' + allOverdueItems.length + ' overdue items and ' + missingDeadlinesCount + ' rows with missing deadlines.');

    const months = generateMonthList(config.DASHBOARD_DATES.START, config.DASHBOARD_DATES.END);
    const dashboardData = months.map(function(month) {
      const key = month.getFullYear() + '-' + month.getMonth();
      // The new summary format is [total, upcoming, overdue, approved, overdueItems[]]
      return monthlySummaries.get(key) || [0, 0, 0, 0, []];
    });

    const grandTotals = dashboardData.reduce(function(totals, summary) {
      totals[0] += summary[0];
      totals[1] += summary[1];
      totals[2] += summary[2];
      totals[3] += summary[3];
      return totals;
    }, [0, 0, 0, 0]);

    renderDashboardTable(dashboardSheet, { allOverdueItems, missingDeadlinesCount }, months, dashboardData, grandTotals, config);

    if (config.DASHBOARD_CHARTING.ENABLED) {
      createOrUpdateDashboardCharts(dashboardSheet, months, dashboardData, config);
      hideDataColumns(dashboardSheet, config);
    }

    SpreadsheetApp.flush();
    const duration = (new Date().getTime() - scriptStartTime.getTime()) / 1000;
    Logger.log('Dashboard update complete (Duration: ' + duration.toFixed(2) + ' seconds).');

  } catch (error) {
    Logger.log('ERROR in updateDashboard: ' + error.message + '\nStack: ' + error.stack);
    notifyError('Dashboard Update Failed', error, ss, config);
    ui.alert('An error occurred updating the dashboard. Please check logs and the notification email.\nError: ' + error.message);
  }
}

/**
 * Efficiently reads the necessary data from the 'Forecasting' sheet.
 * It reads only the columns required for dashboard processing, as defined in the configuration,
 * to optimize performance on large sheets.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} forecastSheet The sheet object for the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {{forecastingValues: any[][], forecastingHeaders: string[]}|null} An object containing the data values and headers, or `null` on failure.
 */
function readForecastingData(forecastSheet, config) {
  try {
    const dataRange = forecastSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const forecastingHeaders = numRows > 0 ? forecastSheet.getRange(1, 1, 1, dataRange.getNumColumns()).getValues()[0] : [];
    if (numRows <= 1) return { forecastingValues: [], forecastingHeaders: forecastingHeaders };

    const colIndices = Object.values(config.FORECASTING_COLS);
    const lastColNumNeeded = Math.max.apply(null, colIndices);
    const numColsToRead = Math.min(lastColNumNeeded, dataRange.getNumColumns());

    const forecastingValues = forecastSheet.getRange(2, 1, numRows - 1, numColsToRead).getValues();
    return { forecastingValues: forecastingValues, forecastingHeaders: forecastingHeaders };
  } catch (e) {
    Logger.log('ERROR reading data from ' + forecastSheet.getName() + ': ' + e.message);
    return null;
  }
}

/**
 * Processes the raw forecasting data to generate aggregated dashboard metrics.
 * It calculates monthly summaries for total, upcoming, overdue, and approved projects. It also compiles a
 * list of all overdue items and counts rows with missing or invalid deadlines.
 *
 * @param {any[][]} forecastingValues A 2D array of data from the 'Forecasting' sheet.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {{monthlySummaries: Map<string, any[]>, allOverdueItems: any[][], missingDeadlinesCount: number}} An object containing the processed data.
 */
function processDashboardData(forecastingValues, config) {
  const monthlySummaries = new Map();
  const allOverdueItems = [];
  let missingDeadlinesCount = 0;

  const tz = 'America/New_York';
  const today = new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'00:00:00'Z'"));

  const FC = config.FORECASTING_COLS;
  const deadlineIdx = FC.DEADLINE - 1;
  const progressIdx = FC.PROGRESS - 1;
  const permitsIdx = FC.PERMITS - 1;
  const projectNameIdx = FC.PROJECT_NAME - 1;

  const S = config.STATUS_STRINGS;
  const approvedLower = normalizeString(S.PERMIT_APPROVED);
  const inProgressLower = normalizeString(S.IN_PROGRESS);
  const scheduledLower = normalizeString(S.SCHEDULED);

  for (let i = 0; i < forecastingValues.length; i++) {
    const row = forecastingValues[i];
    const rawDeadline = row[deadlineIdx];

    const deadlineDate = parseAndNormalizeDate(rawDeadline);

    if (!deadlineDate) {
      if(rawDeadline) missingDeadlinesCount++;
      continue;
    }

    const key = deadlineDate.getFullYear() + '-' + deadlineDate.getMonth();
    if (!monthlySummaries.has(key)) {
      // [total, upcoming, overdue, approved, overdueItems[]]
      monthlySummaries.set(key, [0, 0, 0, 0, []]);
    }
    const monthData = monthlySummaries.get(key);
    monthData[0]++;

    const currentStatus = normalizeString(row[progressIdx]);
    const normalizedDeadline = new Date(Utilities.formatDate(deadlineDate, tz, "yyyy-MM-dd'T'00:00:00'Z'"));

    // Overdue if deadline is today or earlier AND status is 'In Progress'.
    if (normalizedDeadline <= today && currentStatus === inProgressLower) {
        monthData[2]++; // Overdue
        const overdueItem = {
            name: row[projectNameIdx] || 'Unnamed Project',
            deadline: Utilities.formatDate(deadlineDate, Session.getScriptTimeZone(), 'MMM d, yyyy')
        };
        monthData[4].push(overdueItem);
        allOverdueItems.push(row);
    }
    // Upcoming if deadline is today or later AND status is active ('In Progress' or 'Scheduled').
    else if (normalizedDeadline >= today && (currentStatus === inProgressLower || currentStatus === scheduledLower)) {
        monthData[1]++; // Upcoming
    }

    // Check for Permit Approval status, which is an independent category.
    if (normalizeString(row[permitsIdx]) === approvedLower) {
      monthData[3]++; // Approved
    }
  }
  return { monthlySummaries, allOverdueItems, missingDeadlinesCount };
}

/**
 * Renders the main data table on the Dashboard sheet.
 * This function clears and resizes the sheet, sets headers and notes, writes the monthly summary data,
 * adds hover notes for overdue items, displays grand totals, and applies all necessary formatting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dashboardSheet The sheet object for the 'Dashboard'.
 * @param {object} processedData An object containing processed data, including the count of items with missing deadlines.
 * @param {Date[]} months An array of Date objects representing the months in the report.
 * @param {any[][]} dashboardData A 2D array of the monthly summary data to be rendered.
 * @param {number[]} grandTotals An array containing the grand totals for the summary columns.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function renderDashboardTable(dashboardSheet, processedData, months, dashboardData, grandTotals, config) {
  const { missingDeadlinesCount } = processedData;

  clearAndResizeSheet(dashboardSheet, config.DASHBOARD_LAYOUT.FIXED_ROW_COUNT, config.DASHBOARD_LAYOUT.HIDE_COL_END);
  setDashboardHeaders(dashboardSheet, config);
  setDashboardHeaderNotes(dashboardSheet, config);

  const dataStartRow = 2;

  if (dashboardData.length > 0) {
    const DL = config.DASHBOARD_LAYOUT;
    const numDataRows = dashboardData.length;

    /*
     * Data Source and Column Mapping:
     * Col A (Year): Year of the summary period.
     * Col B (Month): Month of the summary period.
     * Col C (Total Projects): Count of all projects with a valid deadline in that month.
     * Col D (Upcoming): Count of active projects ('In Progress' or 'Scheduled') with a deadline on or after today.
     * Col E (Overdue): Count of non-terminal projects with a deadline before today. Hover for details.
     * Col F (Approved): Count of projects with 'Permits' status set to 'approved' in that month.
     * Grand Totals (G-J) are sums of their respective columns.
     */

    // Prepare table data including the overdue count directly.
    const tableData = months.map(function(month, i) {
      const summary = dashboardData[i];
      // [Year, Month, Total, Upcoming, Overdue, Approved]
      return [month.getFullYear(), month, summary[0], summary[1], summary[2], summary[3]];
    });

    // Write all data to the sheet in one call.
    dashboardSheet.getRange(dataStartRow, DL.YEAR_COL, numDataRows, 6).setValues(tableData);

    // Set hover notes for the 'Overdue' column.
    const overdueNotes = dashboardData.map(function(summary) {
      const overdueItems = summary[4] || [];
      if (overdueItems.length === 0) return [null];

      let note = overdueItems.slice(0, 20).map(item => `${item.name} — ${item.deadline}`).join('\n');
      if (overdueItems.length > 20) {
        note += `\n(+${overdueItems.length - 20} more)`;
      }
      return [note];
    });
    dashboardSheet.getRange(dataStartRow, DL.OVERDUE_COL, numDataRows, 1).setNotes(overdueNotes);

    // Set grand totals.
    const [gtTotal, gtUpcoming, gtOverdue, gtApproved] = grandTotals;
    dashboardSheet.getRange(dataStartRow, DL.GT_UPCOMING_COL).setValue(gtUpcoming);
    dashboardSheet.getRange(dataStartRow, DL.GT_OVERDUE_COL).setValue(gtOverdue); // No more hyperlink
    dashboardSheet.getRange(dataStartRow, DL.GT_TOTAL_COL).setValue(gtTotal);
    dashboardSheet.getRange(dataStartRow, DL.GT_APPROVED_COL).setValue(gtApproved);

    // Display count of items with missing deadlines.
    const missingCell = dashboardSheet.getRange(DL.MISSING_DEADLINE_CELL);
    missingCell.setValue('Missing/Invalid Deadlines:');
    missingCell.offset(0, 1).setValue(missingDeadlinesCount).setNumberFormat('0').setFontWeight('bold');
    missingCell.setFontWeight('bold');

    applyDashboardFormatting(dashboardSheet, numDataRows, config);
  }
}

/**
 * Sets explanatory notes on the dashboard header cells.
 * These notes provide users with clear definitions for each column, improving the dashboard's usability.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function setDashboardHeaderNotes(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  sheet.getRange(1, DL.TOTAL_COL).setNote('Total projects with a valid deadline in this month.');
  sheet.getRange(1, DL.UPCOMING_COL).setNote('Active projects ("In Progress" or "Scheduled") with a deadline on or after today.');
  sheet.getRange(1, DL.OVERDUE_COL).setNote('Projects with status "In Progress" and a deadline in the past. Click to see details.');
  sheet.getRange(1, DL.APPROVED_COL).setNote("Projects with 'Permits' status set to 'approved'.");
  sheet.getRange(1, DL.GT_TOTAL_COL).setNote('Grand total for all projects shown in the table.');
}

/**
 * Applies all visual formatting to the dashboard table.
 * This function handles row banding, text alignment, number formats, and borders to create a polished look.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {number} numDataRows The number of data rows in the table to format.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function applyDashboardFormatting(sheet, numDataRows, config) {
  const DL = config.DASHBOARD_LAYOUT;
  const DF = config.DASHBOARD_FORMATTING;
  const dataRange = sheet.getRange(2, DL.YEAR_COL, numDataRows, 6);

  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
           .setHeaderRowColor(null)
           .setFirstRowColor(DF.BANDING_COLOR_ODD)
           .setSecondRowColor(DF.BANDING_COLOR_EVEN);

  sheet.getRange(2, 1, numDataRows, DL.GT_APPROVED_COL).setHorizontalAlignment('center');
  sheet.getRange(2, DL.YEAR_COL,  numDataRows, 1).setNumberFormat('0000');
  sheet.getRange(2, DL.MONTH_COL, numDataRows, 1).setNumberFormat(DF.MONTH_FORMAT);
  sheet.getRange(2, DL.TOTAL_COL, numDataRows, 4).setNumberFormat(DF.COUNT_FORMAT);
  sheet.getRange(2, DL.GT_UPCOMING_COL, 1, 4).setNumberFormat(DF.COUNT_FORMAT);
  sheet.getRange(1, 1, numDataRows + 1, DL.GT_APPROVED_COL)
       .setBorder(true, true, true, true, true, true, DF.BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
}

/**
 * Displays a placeholder message in the area designated for a chart.
 * This is used as a fallback when chart data is unavailable or if a chart fails to render,
 * providing clear feedback to the user.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {number} anchorRow The row where the chart would normally be anchored.
 * @param {number} anchorCol The column where the chart would normally be anchored.
 * @param {string} message The message to display in the placeholder.
 * @returns {void} This function does not return a value.
 */
function displayChartPlaceholder(sheet, anchorRow, anchorCol, message) {
  try {
    var placeholderRange = sheet.getRange(anchorRow + 5, anchorCol, 1, 4);
    placeholderRange.merge();
    placeholderRange.setValue(message).setHorizontalAlignment('center').setVerticalAlignment('middle').setFontStyle('italic').setFontColor('#999999');
  } catch (e) {
    Logger.log('Could not create chart placeholder: ' + e.message);
  }
}

/**
 * Writes and formats the main headers for the dashboard tables and chart area.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function setDashboardHeaders(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  const DF = config.DASHBOARD_FORMATTING;

  // Set headers for the main data table
  const mainHeaders = ['Year', 'Month', 'Total Projects', 'Upcoming', 'Overdue', 'Approved'];
  const mainHeaderRange = sheet.getRange(1, DL.YEAR_COL, 1, mainHeaders.length);
  mainHeaderRange.setValues([mainHeaders]);

  // Set headers for the grand totals
  const gtHeaders = ['GT Upcoming', 'GT Overdue', 'GT Total', 'GT Approved'];
  const gtHeaderRange = sheet.getRange(1, DL.GT_UPCOMING_COL, 1, gtHeaders.length);
  gtHeaderRange.setValues([gtHeaders]);

  // Set header for the Charts column and adjust its width
  const chartsHeaderRange = sheet.getRange(1, DL.CHART_ANCHOR_COL);
  chartsHeaderRange.setValue("Charts");
  sheet.setColumnWidth(DL.CHART_ANCHOR_COL, 485);

  // Apply consistent formatting to all header cells
  const allHeaderRanges = [mainHeaderRange, gtHeaderRange, chartsHeaderRange];
  for (const range of allHeaderRanges) {
    range.setBackground(DF.HEADER_BACKGROUND)
         .setFontColor(DF.HEADER_FONT_COLOR)
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }
}

/**
 * Hides the columns used for chart data staging to keep the user interface clean and uncluttered.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function hideDataColumns(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  if (sheet.getMaxColumns() < DL.HIDE_COL_START) return;
  const numColsToHide = DL.HIDE_COL_END - DL.HIDE_COL_START + 1;
  sheet.hideColumns(DL.HIDE_COL_START, numColsToHide);
}

/**
 * Ensures the sheet has enough columns to store hidden chart data, adding them if necessary, and then hides them.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to modify.
 * @param {number} startCol The starting column for the hidden data range.
 * @param {number} columnsNeeded The total number of columns required for the hidden data.
 * @returns {void} This function does not return a value.
 */
function ensureHiddenColumnCapacity(sheet, startCol, columnsNeeded) {
  const requiredEndCol = startCol + columnsNeeded - 1;
  const currentMaxCol = sheet.getMaxColumns();
  if (currentMaxCol < requiredEndCol) {
    sheet.insertColumnsAfter(currentMaxCol, requiredEndCol - currentMaxCol);
  }
  sheet.hideColumns(startCol, columnsNeeded);
}

/**
 * Ensures the sheet has at least a minimum number of rows, adding them if necessary.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to modify.
 * @param {number} minRows The minimum number of rows required.
 * @returns {void} This function does not return a value.
 */
function ensureRowCapacity(sheet, minRows) {
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < minRows) {
    sheet.insertRowsAfter(currentMaxRows, minRows - currentMaxRows);
  }
}

/**
 * Safely clears the content, notes, and data validations from a block of cells, typically used for hidden chart data.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object containing the block.
 * @param {number} startRow The starting row of the block to clear.
 * @param {number} startCol The starting column of the block to clear.
 * @param {number} numRows The number of rows in the block.
 * @param {number} numCols The number of columns in the block.
 * @returns {void} This function does not return a value.
 */
function clearHiddenBlock(sheet, startRow, startCol, numRows, numCols) {
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (startRow > maxRows || startCol > maxCols) return;
    const actualRows = Math.min(numRows, Math.max(0, maxRows - startRow + 1));
    const actualCols = Math.min(numCols, Math.max(0, maxCols - startCol + 1));
    if (actualRows <= 0 || actualCols <= 0) return;
    sheet.getRange(startRow, startCol, actualRows, actualCols).clearContent().clearDataValidations().clearNote();
  } catch (e) {
    Logger.log('WARNING: Could not clear hidden block at R' + startRow + 'C' + startCol + ': ' + e.message);
  }
}

/**
 * Retrieves a numeric count stored in a cell (typically a hidden one).
 * This is used to track the size of chart data ranges between updates to ensure proper cleanup.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object where the count is stored.
 * @param {number} col The 1-based column index containing the count (in row 1).
 * @returns {number} The stored count, or `0` if the value is not found or invalid.
 */
function getStoredCount(sheet, col) {
  try {
    var v = sheet.getRange(1, col).getValue();
    var n = parseInt(v, 10);
    return isNaN(n) ? 0 : n;
  } catch (e) { return 0; }
}

/**
 * Stores a numeric count in a cell for later retrieval.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to store the count in.
 * @param {number} col The 1-based column index to use for storage (in row 1).
 * @param {number} count The count to store.
 * @returns {void} This function does not return a value.
 */
function setStoredCount(sheet, col, count) {
  try {
    sheet.getRange(1, col).setValue(count);
  } catch (e) { /* non fatal */ }
}

/**
 * Creates or updates the summary charts on the dashboard.
 * It removes all existing charts, stages the necessary data in hidden columns, and then builds new charts
 * for past and upcoming months based on the configuration.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The dashboard sheet object.
 * @param {Date[]} months The array of Date objects representing the months for the x-axis.
 * @param {any[][]} dashboardData The aggregated summary data for all months.
 * @param {object} config The global configuration object (`CONFIG`).
 * @returns {void} This function does not return a value.
 */
function createOrUpdateDashboardCharts(sheet, months, dashboardData, config) {
  sheet.getCharts().forEach(function(chart) { sheet.removeChart(chart); });
  const DC = config.DASHBOARD_CHARTING;
  const DL = config.DASHBOARD_LAYOUT;
  const DF = config.DASHBOARD_FORMATTING;
  const COLORS = DF.CHART_COLORS;
  const STACKED = typeof DC.STACKED === 'boolean' ? DC.STACKED : false;
  const MONTH_FMT = DF.MONTH_FORMAT || 'mmm yyyy';
  try {
    var n = Math.min(months.length, dashboardData.length);
    if (n === 0) {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'No project data available to chart.');
      return;
    }
    var HIDDEN_START_COL = DL.HIDE_COL_START;
    var PAST_COL = HIDDEN_START_COL;
    var UPC_COL = HIDDEN_START_COL + 4;
    var HIDDEN_COLS_NEEDED = 8;
    ensureHiddenColumnCapacity(sheet, HIDDEN_START_COL, HIDDEN_COLS_NEEDED);
    var today = getMonthStart_(new Date());
    var pastStart = getMonthStart_(new Date(today));
    pastStart.setMonth(pastStart.getMonth() - DC.PAST_MONTHS_COUNT);
    var upcomingEnd = getMonthStart_(new Date(today));
    upcomingEnd.setMonth(upcomingEnd.getMonth() + DC.UPCOMING_MONTHS_COUNT);
    var pastData = [];
    var upcomingData = [];
    for (var i = 0; i < n; i++) {
      var m = months[i];
      var d = dashboardData[i];
      var row = [m, d[2], d[1], d[0]];
      if (m >= pastStart && m < today) pastData.push(row);
      else if (m >= today && m < upcomingEnd) upcomingData.push(row);
    }
    var DATA_START_ROW = 2;
    var HEADER = [['Month', 'Overdue', 'Upcoming', 'Total']];
    var neededRows = Math.max(DATA_START_ROW + 1 + pastData.length, DATA_START_ROW + 1 + upcomingData.length, 20);
    ensureRowCapacity(sheet, neededRows);
    var prevPast = getStoredCount(sheet, PAST_COL);
    var prevUpc = getStoredCount(sheet, UPC_COL);
    var rowsToClear = Math.max(pastData.length, upcomingData.length, prevPast, prevUpc) + 2;
    clearHiddenBlock(sheet, DATA_START_ROW, PAST_COL, rowsToClear, 4);
    clearHiddenBlock(sheet, DATA_START_ROW, UPC_COL, rowsToClear, 4);
    sheet.getRange(DATA_START_ROW, PAST_COL, 1, 4).setValues(HEADER);
    sheet.getRange(DATA_START_ROW, UPC_COL, 1, 4).setValues(HEADER);
    if (pastData.length > 0) {
      sheet.getRange(DATA_START_ROW + 1, PAST_COL, pastData.length, 4).setValues(pastData);
      sheet.getRange(DATA_START_ROW + 1, PAST_COL, pastData.length, 1).setNumberFormat(MONTH_FMT);
    }
    if (upcomingData.length > 0) {
      sheet.getRange(DATA_START_ROW + 1, UPC_COL, upcomingData.length, 4).setValues(upcomingData);
      sheet.getRange(DATA_START_ROW + 1, UPC_COL, upcomingData.length, 1).setNumberFormat(MONTH_FMT);
    }
    setStoredCount(sheet, PAST_COL, pastData.length);
    setStoredCount(sheet, UPC_COL, upcomingData.length);
    var buildChart = function(title, leftCol, rowsCount, anchorRow) {
      if (rowsCount <= 0) return null;
      var range = sheet.getRange(DATA_START_ROW, leftCol, rowsCount + 1, 4);
      return sheet.newChart().asColumnChart().addRange(range).setNumHeaders(1)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
        .setOption('title', title).setOption('width', DC.CHART_WIDTH).setOption('height', DC.CHART_HEIGHT)
        .setOption('colors', [COLORS.overdue, COLORS.upcoming, COLORS.total])
        .setOption('legend', { position: 'top' }).setOption('isStacked', STACKED)
        .setPosition(anchorRow, DL.CHART_ANCHOR_COL, 0, 0).build();
    };
    if (pastData.length > 0) {
      var c1 = buildChart('Past ' + DC.PAST_MONTHS_COUNT + ' Months: Overdue, Upcoming, Total', PAST_COL, pastData.length, DL.CHART_START_ROW);
      if (c1) sheet.insertChart(c1);
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'No project data found for the past ' + DC.PAST_MONTHS_COUNT + ' months.');
    }
    if (upcomingData.length > 0) {
      var c2 = buildChart('Next ' + DC.UPCOMING_MONTHS_COUNT + ' Months: Overdue, Upcoming, Total', UPC_COL, upcomingData.length, DL.CHART_START_ROW + DC.ROW_SPACING);
      if (c2) sheet.insertChart(c2);
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW + DC.ROW_SPACING, DL.CHART_ANCHOR_COL, 'No project data found for the next ' + DC.UPCOMING_MONTHS_COUNT + ' months.');
    }
  } catch (error) {
    Logger.log('ERROR in createOrUpdateDashboardCharts: ' + error.message + '\n' + error.stack);
    displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'Chart creation failed. Check logs for details.');
  }
}

/**
 * Normalizes a given date to the beginning of its month (midnight on the 1st day).
 *
 * @private
 * @param {Date} d The date to normalize.
 * @returns {Date} A new Date object set to the start of the month.
 */
function getMonthStart_(d) {
  var x = new Date(d);
  x.setDate(1);
  x.setHours(0, 0, 0, 0);
  return x;
}

/**
 * Generates an array of Date objects, with one entry for each month between a start and end date, inclusive.
 *
 * @param {Date} startDate The first month to include in the list.
 * @param {Date} endDate The last month to include in the list.
 * @returns {Date[]} An array of Date objects, where each object represents the first day of a month.
 */
function generateMonthList(startDate, endDate) {
  const months = [];
  var current = new Date(startDate.getTime());
  current = getMonthStart_(current);
  var end = getMonthStart_(new Date(endDate.getTime()));
  while (current <= end) {
    months.push(new Date(current));
    current.setMonth(current.getMonth() + 1);
  }
  return months;
}

/**
 * A verification function to audit the Dashboard's overdue calculations against the source data.
 * This can be run manually from the Apps Script editor to ensure data integrity. It checks:
 * 1. If the grand total overdue count on the dashboard matches a direct recalculation from the 'Forecasting' sheet.
 * 2. If the monthly overdue counts match the number of items listed in their corresponding hover notes.
 *
 * @returns {void} This function does not return a value; it logs its findings.
 */
function runDashboardVerification() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG;
  const { FORECASTING, DASHBOARD } = config.SHEETS;
  const forecastSheet = ss.getSheetByName(FORECASTING);
  const dashboardSheet = ss.getSheetByName(DASHBOARD);

  if (!forecastSheet || !dashboardSheet) {
    Logger.log('Verification failed: Could not find required sheets.');
    return;
  }

  Logger.log('Starting Dashboard verification...');

  // 1. Recalculate overdue count directly from Forecasting sheet.
  const { forecastingValues } = readForecastingData(forecastSheet, config);
  const { allOverdueItems } = processDashboardData(forecastingValues, config);
  const calculatedGrandTotal = allOverdueItems.length;

  // 2. Get the Grand Total Overdue value from the Dashboard.
  const DL = config.DASHBOARD_LAYOUT;
  const dashboardGrandTotal = dashboardSheet.getRange(2, DL.GT_OVERDUE_COL).getValue();

  // 3. Compare and log the results.
  Logger.log(`Calculated Grand Total Overdue: ${calculatedGrandTotal}`);
  Logger.log(`Dashboard Grand Total Overdue: ${dashboardGrandTotal}`);
  if (calculatedGrandTotal === dashboardGrandTotal) {
    Logger.log('✅ SUCCESS: Grand totals match.');
  } else {
    Logger.log(`❌ FAILURE: Grand totals DO NOT match. Discrepancy: ${calculatedGrandTotal - dashboardGrandTotal}`);
  }

  // 4. Spot-check a monthly total against its hover note.
  const firstDataRowWithOverdue = dashboardSheet.getRange(2, DL.OVERDUE_COL, dashboardSheet.getLastRow() -1, 1)
                                      .getValues()
                                      .findIndex(row => row[0] > 0);

  if (firstDataRowWithOverdue !== -1) {
    const checkRow = firstDataRowWithOverdue + 2; // +2 for 0-index and header
    const monthlyCount = dashboardSheet.getRange(checkRow, DL.OVERDUE_COL).getValue();
    const note = dashboardSheet.getRange(checkRow, DL.OVERDUE_COL).getNote();
    const noteLines = note ? note.split('\n').filter(line => line.trim() !== '' && !line.startsWith('(+')) : [];
    const noteCount = noteLines.length;

    Logger.log(`\nSpot-checking row ${checkRow}...`);
    Logger.log(`- Monthly overdue count from cell: ${monthlyCount}`);
    Logger.log(`- Number of items in hover note: ${noteCount}`);

    // In a truncated note, the count won't match, but this check is still useful.
    if (monthlyCount === noteCount) {
      Logger.log('✅ SUCCESS: Monthly cell count matches the number of items in its hover note.');
    } else if (note.includes('(+')) {
       Logger.log('INFO: Hover note is truncated. Cannot perform exact count match for this cell.');
    } else {
      Logger.log(`❌ FAILURE: Monthly cell count (${monthlyCount}) does not match note count (${noteCount}).`);
    }
  } else {
    Logger.log('\nNo overdue items found in any month to spot-check.');
  }

  Logger.log('\nVerification complete.');
}