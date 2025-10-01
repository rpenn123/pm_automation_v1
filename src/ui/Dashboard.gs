/**
 * @OnlyCurrentDoc
 * Dashboard.gs
 * Generates the main dashboard: data processing, table rendering, charts, and overdue drill-down.
 * Design goals: correctness, idempotence, performance, and clean UX.
 */

/**
 * Main orchestrator to generate or update the Dashboard.
 * Entry point for custom menu.
 *
 * Flow:
 * 1) Initialize handles
 * 2) Read source data
 * 3) Process data (single pass)
 * 4) Populate Overdue Details
 * 5) Render Dashboard table and notes
 * 6) Format and build charts
 */
function updateDashboard() {
  const ui = SpreadsheetApp.getUi();
  const scriptStartTime = new Date();
  Logger.log('Dashboard update initiated at ' + scriptStartTime.toLocaleString());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG; // The only function that should access the global CONFIG.

  try {
    const { FORECASTING, DASHBOARD, OVERDUE_DETAILS } = config.SHEETS;

    const forecastSheet = ss.getSheetByName(FORECASTING);
    if (!forecastSheet) throw new Error('Sheet "' + FORECASTING + '" not found.');

    // Destination sheets
    const dashboardSheet = getOrCreateSheet(ss, DASHBOARD);
    const overdueDetailsSheet = getOrCreateSheet(ss, OVERDUE_DETAILS);
    const overdueSheetGid = overdueDetailsSheet.getSheetId();

    // 1) Read
    const read = readForecastingData(forecastSheet, config);
    if (!read || !read.forecastingValues) throw new Error('Failed to read data from ' + FORECASTING + '.');
    const { forecastingValues, forecastingHeaders } = read;

    // 2) Process
    const processed = processForecastingData(forecastingValues, config);
    const { monthlySummaries, grandTotals, allOverdueItems, missingDeadlinesCount } = processed;
    Logger.log('Processing complete. Found ' + allOverdueItems.length + ' overdue items and ' + missingDeadlinesCount + ' rows with missing deadlines.');

    // 3) Overdue sheet
    populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders);

    // 4) Dashboard table
    clearAndResizeSheet(dashboardSheet, config.DASHBOARD_LAYOUT.FIXED_ROW_COUNT, config.DASHBOARD_LAYOUT.HIDE_COL_END);
    setDashboardHeaders(dashboardSheet, config);
    setDashboardHeaderNotes(dashboardSheet, config);

    const months = generateMonthList(config.DASHBOARD_DATES.START, config.DASHBOARD_DATES.END);
    const dataStartRow = 2;

    // Align processed map to ordered month list
    const dashboardData = months.map(function(month) {
      const key = month.getFullYear() + '-' + month.getMonth();
      // [total, upcoming, overdue, approved]
      return monthlySummaries.get(key) || [0, 0, 0, 0];
    });

    if (dashboardData.length > 0) {
      const DL = config.DASHBOARD_LAYOUT;
      const numDataRows = dashboardData.length;

      // Batch write dashboard table
      const tableData = months.map(function(month, i) {
        const summary = dashboardData[i]; // [total, upcoming, overdue, approved]
        return [
          month.getFullYear(), // Year
          month,               // Month
          summary[0],          // Total Projects
          summary[1],          // Upcoming
          null,                // Placeholder for Overdue formula, which is set next
          summary[3]           // Approved
        ];
      });

      const overdueFormulas = dashboardData.map(function(row) {
        return ['=HYPERLINK("#gid=' + overdueSheetGid + '", ' + (row[2] || 0) + ')'];
      });

      // Write main data block, then overwrite the formula column
      dashboardSheet.getRange(dataStartRow, DL.YEAR_COL, numDataRows, 6).setValues(tableData);
      dashboardSheet.getRange(dataStartRow, DL.OVERDUE_COL, numDataRows, 1).setFormulas(overdueFormulas);

      // Grand totals, now aligned with monthlySummaries: [total, upcoming, overdue, approved]
      const gtTotal    = grandTotals[0];
      const gtUpcoming = grandTotals[1];
      const gtOverdue  = grandTotals[2];
      const gtApproved = grandTotals[3];

      // Note the order of assignment now matches the GT column order on the sheet
      dashboardSheet.getRange(dataStartRow, DL.GT_UPCOMING_COL).setValue(gtUpcoming);
      dashboardSheet.getRange(dataStartRow, DL.GT_OVERDUE_COL).setFormula('=HYPERLINK("#gid=' + overdueSheetGid + '", ' + gtOverdue + ')');
      dashboardSheet.getRange(dataStartRow, DL.GT_TOTAL_COL).setValue(gtTotal);
      dashboardSheet.getRange(dataStartRow, DL.GT_APPROVED_COL).setValue(gtApproved);

      // Missing deadlines note
      const missingCell = dashboardSheet.getRange(DL.MISSING_DEADLINE_CELL);
      missingCell.setValue('Missing/Invalid Deadlines:');
      missingCell.offset(0, 1).setValue(missingDeadlinesCount).setNumberFormat('0').setFontWeight('bold');
      missingCell.setFontWeight('bold');

      // 5) Format
      applyDashboardFormatting(dashboardSheet, numDataRows, config);

      // 6) Charts
      if (config.DASHBOARD_CHARTING.ENABLED) {
        createOrUpdateDashboardCharts(dashboardSheet, months, dashboardData, config);
        hideDataColumns(dashboardSheet, config);
      }
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

// =================================================================
// ==================== DATA PROCESSING ============================
// =================================================================

/**
 * Efficient read of the Forecasting sheet.
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
 * Single-pass processing that builds monthly summaries and totals.
 * Returns:
 * - monthlySummaries: Map key "YYYY-M" -> [total, upcoming, overdue, approved]
 * - grandTotals: [total, upcoming, overdue, approved]
 */
function processForecastingData(forecastingValues, config) {
  const monthlySummaries = new Map();
  const allOverdueItems = [];
  // Standardized to [total, upcoming, overdue, approved] to match monthlySummaries
  var grandTotals = [0, 0, 0, 0];
  var missingDeadlinesCount = 0;

  var today = new Date();
  today.setHours(0, 0, 0, 0);

  const FC = config.FORECASTING_COLS;
  const deadlineIdx = FC.DEADLINE - 1;
  const progressIdx = FC.PROGRESS - 1;
  const permitsIdx  = FC.PERMITS - 1;

  const S = config.STATUS_STRINGS;
  const inProgressLower  = S.IN_PROGRESS.toLowerCase();
  const scheduledLower   = S.SCHEDULED.toLowerCase();
  const approvedLower    = S.PERMIT_APPROVED.toLowerCase();

  for (var i = 0; i < forecastingValues.length; i++) {
    var row = forecastingValues[i];
    var deadlineDate = parseAndNormalizeDate(row[deadlineIdx]);

    if (!deadlineDate) {
      missingDeadlinesCount++;
      continue;
    }

    var key = deadlineDate.getFullYear() + '-' + deadlineDate.getMonth();
    if (!monthlySummaries.has(key)) {
      monthlySummaries.set(key, [0, 0, 0, 0]); // [total, upcoming, overdue, approved]
    }
    var monthData = monthlySummaries.get(key);

    monthData[0]++;       // total in month
    grandTotals[0]++;     // GT total

    var currentStatus = normalizeString(row[progressIdx]);
    var isInProgress  = currentStatus === inProgressLower;
    var isScheduled   = currentStatus === scheduledLower;

    if (isInProgress || isScheduled) {
      if (deadlineDate > today) {
        monthData[1]++;   // upcoming
        grandTotals[1]++; // GT upcoming
      } else if (isInProgress) {
        monthData[2]++;   // overdue
        grandTotals[2]++; // GT overdue
        allOverdueItems.push(row);
      }
    }

    if (normalizeString(row[permitsIdx]) === approvedLower) {
      monthData[3]++;     // approved
      grandTotals[3]++;
    }
  }

  return { monthlySummaries: monthlySummaries, grandTotals: grandTotals, allOverdueItems: allOverdueItems, missingDeadlinesCount: missingDeadlinesCount };
}

// =================================================================
// ==================== PRESENTATION LOGIC =========================
// =================================================================

function displayChartPlaceholder(sheet, anchorRow, anchorCol, message) {
  try {
    var placeholderRange = sheet.getRange(anchorRow + 5, anchorCol, 1, 4);
    placeholderRange.merge();
    placeholderRange.setValue(message)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setFontStyle('italic')
      .setFontColor('#999999');
  } catch (e) {
    Logger.log('Could not create chart placeholder: ' + e.message);
  }
}

function populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders) {
  try {
    if (!forecastingHeaders || forecastingHeaders.length === 0) {
      overdueDetailsSheet.clear();
      overdueDetailsSheet.getRange(1, 1).setValue("Source 'Forecasting' sheet is empty or has no header row.");
      Logger.log("Skipped populating Overdue Details: 'Forecasting' sheet appears to be empty.");
      return;
    }

    var numRows = allOverdueItems.length;
    var numCols = allOverdueItems.length > 0 ? allOverdueItems[0].length : forecastingHeaders.length;

    overdueDetailsSheet.clear();

    if (overdueDetailsSheet.getMaxRows() > 1) {
      overdueDetailsSheet.deleteRows(2, overdueDetailsSheet.getMaxRows() - 1);
    }
    if (overdueDetailsSheet.getMaxColumns() > numCols) {
      overdueDetailsSheet.deleteColumns(numCols + 1, overdueDetailsSheet.getMaxColumns() - numCols);
    }

    var headersToWrite = forecastingHeaders.slice(0, numCols);
    overdueDetailsSheet.getRange(1, 1, 1, headersToWrite.length).setValues([headersToWrite]).setFontWeight('bold');

    if (numRows > 0) {
      if (overdueDetailsSheet.getMaxRows() < numRows + 1) {
        overdueDetailsSheet.insertRowsAfter(1, numRows);
      }
      overdueDetailsSheet.getRange(2, 1, numRows, numCols).setValues(allOverdueItems);
    }
    Logger.log('Populated Overdue Details sheet with ' + numRows + ' items.');
  } catch (e) {
    Logger.log('ERROR in populateOverdueDetailsSheet: ' + e.message);
  }
}

function setDashboardHeaders(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  const DF = config.DASHBOARD_FORMATTING;

  const headers = [
    'Year', 'Month', 'Total Projects', 'Upcoming', 'Overdue', 'Approved',
    'GT Upcoming', 'GT Overdue', 'GT Total', 'GT Approved'
  ];
  const headerRanges = [
    sheet.getRange(1, DL.YEAR_COL, 1, 6),
    sheet.getRange(1, DL.GT_UPCOMING_COL, 1, 4)
  ];

  headerRanges[0].setValues([headers.slice(0, 6)]);
  headerRanges[1].setValues([headers.slice(6, 10)]);

  for (var i = 0; i < headerRanges.length; i++) {
    headerRanges[i]
      .setBackground(DF.HEADER_BACKGROUND)
      .setFontColor(DF.HEADER_FONT_COLOR)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
}

function setDashboardHeaderNotes(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  sheet.getRange(1, DL.TOTAL_COL).setNote('Total projects with a deadline in this month.');
  sheet.getRange(1, DL.UPCOMING_COL).setNote('Projects "In Progress" or "Scheduled" with a deadline in the future.');
  sheet.getRange(1, DL.OVERDUE_COL).setNote('Projects "In Progress" with a deadline in the past. Click number to see details.');
  sheet.getRange(1, DL.APPROVED_COL).setNote("Projects with 'Permits' status set to 'approved'.");
  sheet.getRange(1, DL.GT_TOTAL_COL).setNote('Grand total of all projects with a valid deadline.');
}

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

function hideDataColumns(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  if (sheet.getMaxColumns() < DL.HIDE_COL_START) {
    Logger.log('Skipping hideDataColumns: Sheet only has ' + sheet.getMaxColumns() + ' columns, less than required ' + DL.HIDE_COL_START + '.');
    return;
  }
  const numColsToHide = DL.HIDE_COL_END - DL.HIDE_COL_START + 1;
  sheet.hideColumns(DL.HIDE_COL_START, numColsToHide);
}

// =================================================================
// ==================== CHARTS: HIDDEN DATA ========================
// =================================================================

/**
 * Ensure hidden column capacity and always hide the span.
 */
function ensureHiddenColumnCapacity(sheet, startCol, columnsNeeded) {
  const requiredEndCol = startCol + columnsNeeded - 1;
  const currentMaxCol = sheet.getMaxColumns();

  if (currentMaxCol < requiredEndCol) {
    const toAdd = requiredEndCol - currentMaxCol;
    sheet.insertColumnsAfter(currentMaxCol, toAdd);
    Logger.log('Added ' + toAdd + ' columns to accommodate chart data.');
  }

  // Always hide the full span to avoid partial visibility
  sheet.hideColumns(startCol, columnsNeeded);
}

/**
 * Ensure row capacity to at least the requested row count.
 */
function ensureRowCapacity(sheet, minRows) {
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < minRows) {
    sheet.insertRowsAfter(currentMaxRows, minRows - currentMaxRows);
    Logger.log('Added ' + (minRows - currentMaxRows) + ' rows to accommodate chart data.');
  }
}

/**
 * Safe clear of a hidden block with bounds checks.
 */
function clearHiddenBlock(sheet, startRow, startCol, numRows, numCols) {
  try {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();

    if (startRow > maxRows || startCol > maxCols) {
      Logger.log('WARNING: Clear range out of bounds. Sheet ' + maxRows + 'x' + maxCols + ', Requested R' + startRow + 'C' + startCol);
      return;
    }
    const actualRows = Math.min(numRows, Math.max(0, maxRows - startRow + 1));
    const actualCols = Math.min(numCols, Math.max(0, maxCols - startCol + 1));

    if (actualRows <= 0 || actualCols <= 0) return;

    sheet.getRange(startRow, startCol, actualRows, actualCols)
         .clearContent()
         .clearDataValidations()
         .clearNote();

  } catch (e) {
    Logger.log('WARNING: Could not clear hidden block at R' + startRow + 'C' + startCol + ': ' + e.message);
  }
}

/**
 * Tiny assert helper that logs if condition fails.
 */
function assertCondition(condition, message) {
  if (!condition) Logger.log('ASSERT: ' + message);
}

/**
 * Persistent counters for previous chart row counts using hidden cells.
 * Stored at row 1 of each hidden table's first column.
 */
function getStoredCount(sheet, col) {
  try {
    var v = sheet.getRange(1, col).getValue();
    var n = parseInt(v, 10);
    return isNaN(n) ? 0 : n;
  } catch (e) {
    return 0;
  }
}

function setStoredCount(sheet, col, count) {
  try {
    sheet.getRange(1, col).setValue(count);
  } catch (e) {
    // non fatal
  }
}

/**
 * Build charts from hidden data blocks that live on the dashboard sheet.
 * Two 4-col tables: Past and Upcoming.
 * Column order per table: [Month, Overdue, Upcoming, Total]
 */
function createOrUpdateDashboardCharts(sheet, months, dashboardData, config) {
  // Remove existing charts first for idempotence
  sheet.getCharts().forEach(function(chart) { sheet.removeChart(chart); });

  const DC = config.DASHBOARD_CHARTING;
  const DL = config.DASHBOARD_LAYOUT;
  const DF = config.DASHBOARD_FORMATTING;
  const COLORS = DF.CHART_COLORS;
  const STACKED = typeof DC.STACKED === 'boolean' ? DC.STACKED : false;
  const MONTH_FMT = DF.MONTH_FORMAT || 'mmm yyyy';

  try {
    // Align lengths to avoid range errors
    var n = Math.min(months.length, dashboardData.length);
    if (n === 0) {
      Logger.log('No data available for charts.');
      displayChartPlaceholder(
        sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL,
        'No project data available to chart.'
      );
      return;
    }

    // Hidden layout: two adjacent 4-col tables
    var HIDDEN_START_COL = DL.HIDE_COL_START;
    var PAST_COL = HIDDEN_START_COL;        // 4 cols
    var UPC_COL  = HIDDEN_START_COL + 4;    // 4 cols
    var HIDDEN_COLS_NEEDED = 8;

    ensureHiddenColumnCapacity(sheet, HIDDEN_START_COL, HIDDEN_COLS_NEEDED);

    // Month windows
    var today = getMonthStart_(new Date());
    var pastStart = getMonthStart_(new Date(today));
    pastStart.setMonth(pastStart.getMonth() - DC.PAST_MONTHS_COUNT);

    var upcomingEnd = getMonthStart_(new Date(today));
    upcomingEnd.setMonth(upcomingEnd.getMonth() + DC.UPCOMING_MONTHS_COUNT);

    // Build filtered rows inline
    var pastData = [];
    var upcomingData = [];
    for (var i = 0; i < n; i++) {
      var m = months[i];
      var d = dashboardData[i]; // [total, upcoming, overdue, approved]
      var row = [m, d[2], d[1], d[0]];
      if (m >= pastStart && m < today) pastData.push(row);
      else if (m >= today && m < upcomingEnd) upcomingData.push(row);
    }

    var DATA_START_ROW = 2;
    var HEADER = [['Month', 'Overdue', 'Upcoming', 'Total']];

    // Ensure we have enough rows for writes
    var neededRows = Math.max(
      DATA_START_ROW + 1 + pastData.length,
      DATA_START_ROW + 1 + upcomingData.length,
      20
    );
    ensureRowCapacity(sheet, neededRows);

    // Use persistent stored counts to clear precisely
    var prevPast = getStoredCount(sheet, PAST_COL);
    var prevUpc  = getStoredCount(sheet, UPC_COL);
    var rowsToClear = Math.max(pastData.length, upcomingData.length, prevPast, prevUpc) + 2;

    // Clear old blocks, then write headers and data
    clearHiddenBlock(sheet, DATA_START_ROW, PAST_COL, rowsToClear, 4);
    clearHiddenBlock(sheet, DATA_START_ROW, UPC_COL,  rowsToClear, 4);

    sheet.getRange(DATA_START_ROW, PAST_COL, 1, 4).setValues(HEADER);
    sheet.getRange(DATA_START_ROW, UPC_COL,  1, 4).setValues(HEADER);

    if (pastData.length > 0) {
      sheet.getRange(DATA_START_ROW + 1, PAST_COL, pastData.length, 4).setValues(pastData);
      sheet.getRange(DATA_START_ROW + 1, PAST_COL, pastData.length, 1).setNumberFormat(MONTH_FMT);
    }
    if (upcomingData.length > 0) {
      sheet.getRange(DATA_START_ROW + 1, UPC_COL, upcomingData.length, 4).setValues(upcomingData);
      sheet.getRange(DATA_START_ROW + 1, UPC_COL, upcomingData.length, 1).setNumberFormat(MONTH_FMT);
    }

    // Store current counts for next run
    setStoredCount(sheet, PAST_COL, pastData.length);
    setStoredCount(sheet, UPC_COL,  upcomingData.length);

    // Chart builder
    var buildChart = function(title, leftCol, rowsCount, anchorRow) {
      if (rowsCount <= 0) return null;
      var range = sheet.getRange(DATA_START_ROW, leftCol, rowsCount + 1, 4); // includes header
      return sheet.newChart()
        .asColumnChart()
        .addRange(range)
        .setNumHeaders(1)
        .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_ROWS)
        .setOption('title', title)
        .setOption('width', DC.CHART_WIDTH)
        .setOption('height', DC.CHART_HEIGHT)
        .setOption('colors', [COLORS.overdue, COLORS.upcoming, COLORS.total])
        .setOption('legend', { position: 'top' })
        .setOption('isStacked', STACKED)
        .setPosition(anchorRow, DL.CHART_ANCHOR_COL, 0, 0)
        .build();
    };

    // Insert or placeholders
    if (pastData.length > 0) {
      var c1 = buildChart('Past ' + pastData.length + ' Months: Overdue, Upcoming, Total',
                          PAST_COL, pastData.length, DL.CHART_START_ROW);
      if (c1) sheet.insertChart(c1);
      Logger.log('Created past months chart with ' + pastData.length + ' data points.');
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL,
        'No project data found for the past ' + DC.PAST_MONTHS_COUNT + ' months.');
      Logger.log('Skipping past months chart: No data available.');
    }

    if (upcomingData.length > 0) {
      var c2 = buildChart('Next ' + upcomingData.length + ' Months: Overdue, Upcoming, Total',
                          UPC_COL, upcomingData.length, DL.CHART_START_ROW + DC.ROW_SPACING);
      if (c2) sheet.insertChart(c2);
      Logger.log('Created upcoming months chart with ' + upcomingData.length + ' data points.');
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW + DC.ROW_SPACING, DL.CHART_ANCHOR_COL,
        'No project data found for the next ' + DC.UPCOMING_MONTHS_COUNT + ' months.');
      Logger.log('Skipping upcoming months chart: No data available.');
    }

    // Sanity checks
    assertCondition(sheet.getMaxColumns() >= (HIDDEN_START_COL + HIDDEN_COLS_NEEDED - 1), 'Hidden column capacity insufficient post-build.');
    assertCondition(sheet.getMaxRows() >= neededRows, 'Row capacity insufficient post-build.');

  } catch (error) {
    Logger.log('ERROR in createOrUpdateDashboardCharts: ' + error.message + '\n' + error.stack);
    // Do not throw. Keep dashboard usable.
    displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL,
      'Chart creation failed. Check logs for details.'
    );
  }
}

// =================================================================
// ==================== DATE HELPERS ===============================
// =================================================================

/**
 * Start of month for a given date.
 */
function getMonthStart_(d) {
  var x = new Date(d);
  x.setDate(1);
  x.setHours(0, 0, 0, 0);
  return x;
}

/**
 * Generate first-of-month Date objects from start to end inclusive.
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