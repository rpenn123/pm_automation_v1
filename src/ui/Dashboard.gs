/**
 * @OnlyCurrentDoc
 * Dashboard.gs
 * Generates the main dashboard: data processing, table rendering, charts, and overdue drill-down.
 * Design goals: correctness, idempotence, performance, and clean UX.
 *
 * Version History:
 * V1.3.0 - 2025-10-07 - Expert GAS Architect
 *    - Merged structural and logical fixes to resolve conflicts.
 *    - Finalized Grand Total calculation to be derived from filtered data.
 *    - Finalized categorization logic to correctly handle 'Upcoming' (>= today) and 'Overdue' (all active statuses).
 * V1.2.0 - 2025-10-07 - Expert GAS Architect
 *    - Fixed boundary condition error in processDashboardData ("Upcoming" now includes today using >=).
 *    - Fixed exclusionary logic error in processDashboardData ("Overdue" now includes "Scheduled" status).
 * V1.1.0 - [Previous Date] - [Previous Author]
 *    - Initial refactoring (structure/performance) - Logic errors persisted.
 */

/**
 * Main orchestrator to generate or update the Dashboard.
 */
function updateDashboard() {
  const ui = SpreadsheetApp.getUi();
  const scriptStartTime = new Date();
  Logger.log('Dashboard update initiated at ' + scriptStartTime.toLocaleString());
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = CONFIG;

  try {
    const { FORECASTING, DASHBOARD, OVERDUE_DETAILS } = config.SHEETS;
    const forecastSheet = ss.getSheetByName(FORECASTING);
    if (!forecastSheet) throw new Error('Sheet "' + FORECASTING + '" not found.');

    const dashboardSheet = getOrCreateSheet(ss, DASHBOARD);
    const overdueDetailsSheet = getOrCreateSheet(ss, OVERDUE_DETAILS);
    const overdueSheetGid = overdueDetailsSheet.getSheetId();

    const read = readForecastingData(forecastSheet, config);
    if (!read || !read.forecastingValues) throw new Error('Failed to read data from ' + FORECASTING + '.');
    const { forecastingValues, forecastingHeaders } = read;

    const processed = processDashboardData(forecastingValues, config);
    const { monthlySummaries, allOverdueItems, missingDeadlinesCount } = processed;
    Logger.log('Processing complete. Found ' + allOverdueItems.length + ' overdue items and ' + missingDeadlinesCount + ' rows with missing deadlines.');

    populateOverdueDetailsSheet(overdueDetailsSheet, allOverdueItems, forecastingHeaders);

    const months = generateMonthList(config.DASHBOARD_DATES.START, config.DASHBOARD_DATES.END);
    const dashboardData = months.map(function(month) {
      const key = month.getFullYear() + '-' + month.getMonth();
      return monthlySummaries.get(key) || [0, 0, 0, 0];
    });

    // MERGED FIX: Grand Totals are calculated from the final filtered data.
    const grandTotals = dashboardData.reduce(function(totals, summary) {
      totals[0] += summary[0]; // total
      totals[1] += summary[1]; // upcoming
      totals[2] += summary[2]; // overdue
      totals[3] += summary[3]; // approved
      return totals;
    }, [0, 0, 0, 0]);

    renderDashboardTable(dashboardSheet, overdueSheetGid, { allOverdueItems, missingDeadlinesCount }, months, dashboardData, grandTotals, config);

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
 * MERGED & REFACTORED: Single-pass processing with corrected logic.
 */
function processDashboardData(forecastingValues, config) {
  const monthlySummaries = new Map();
  const allOverdueItems = [];
  let missingDeadlinesCount = 0;

  const today = parseAndNormalizeDate(new Date());

  const FC = config.FORECASTING_COLS;
  const deadlineIdx = FC.DEADLINE - 1;
  const progressIdx = FC.PROGRESS - 1;
  const permitsIdx = FC.PERMITS - 1;

  const S = config.STATUS_STRINGS;
  const approvedLower = normalizeString(S.PERMIT_APPROVED);
  const completedLower = normalizeString(S.COMPLETED || 'Completed');
  const cancelledLower = normalizeString(S.CANCELLED || 'Cancelled');
  const inProgressLower = normalizeString(S.IN_PROGRESS);
  const scheduledLower = normalizeString(S.SCHEDULED);

  for (let i = 0; i < forecastingValues.length; i++) {
    const row = forecastingValues[i];
    const deadlineDate = parseAndNormalizeDate(row[deadlineIdx]);

    if (!deadlineDate) {
      missingDeadlinesCount++;
      continue;
    }

    const key = deadlineDate.getFullYear() + '-' + deadlineDate.getMonth();
    if (!monthlySummaries.has(key)) {
      monthlySummaries.set(key, [0, 0, 0, 0]);
    }
    const monthData = monthlySummaries.get(key);

    monthData[0]++;

    const currentStatus = normalizeString(row[progressIdx]);
    const isComplete = (currentStatus === completedLower || currentStatus === cancelledLower);
    const isActive = (currentStatus === inProgressLower || currentStatus === scheduledLower);

    // MERGED FIX: Use `isComplete` and corrected date/status logic.
    if (isActive && !isComplete) {
      if (deadlineDate >= today) { // FIX: Upcoming includes today
        monthData[1]++;
      } else { // FIX: Overdue includes all active, past-due statuses
        monthData[2]++;
        allOverdueItems.push(row);
      }
    }

    if (normalizeString(row[permitsIdx]) === approvedLower) {
      monthData[3]++;
    }
  }

  return { monthlySummaries, allOverdueItems, missingDeadlinesCount };
}


function renderDashboardTable(dashboardSheet, overdueSheetGid, processedData, months, dashboardData, grandTotals, config) {
  const { missingDeadlinesCount } = processedData;

  clearAndResizeSheet(dashboardSheet, config.DASHBOARD_LAYOUT.FIXED_ROW_COUNT, config.DASHBOARD_LAYOUT.HIDE_COL_END);
  setDashboardHeaders(dashboardSheet, config);
  setDashboardHeaderNotes(dashboardSheet, config);

  const dataStartRow = 2;

  if (dashboardData.length > 0) {
    const DL = config.DASHBOARD_LAYOUT;
    const numDataRows = dashboardData.length;

    const tableData = months.map(function(month, i) {
      const summary = dashboardData[i];
      return [month.getFullYear(), month, summary[0], summary[1], null, summary[3]];
    });

    const overdueFormulas = dashboardData.map(function(summary) {
      return ['=HYPERLINK("#gid=' + overdueSheetGid + '", ' + (summary[2] || 0) + ')'];
    });

    dashboardSheet.getRange(dataStartRow, DL.YEAR_COL, numDataRows, 6).setValues(tableData);
    dashboardSheet.getRange(dataStartRow, DL.OVERDUE_COL, numDataRows, 1).setFormulas(overdueFormulas);

    const [gtTotal, gtUpcoming, gtOverdue, gtApproved] = grandTotals;

    dashboardSheet.getRange(dataStartRow, DL.GT_UPCOMING_COL).setValue(gtUpcoming);
    dashboardSheet.getRange(dataStartRow, DL.GT_OVERDUE_COL).setFormula('=HYPERLINK("#gid=' + overdueSheetGid + '", ' + gtOverdue + ')');
    dashboardSheet.getRange(dataStartRow, DL.GT_TOTAL_COL).setValue(gtTotal);
    dashboardSheet.getRange(dataStartRow, DL.GT_APPROVED_COL).setValue(gtApproved);

    const missingCell = dashboardSheet.getRange(DL.MISSING_DEADLINE_CELL);
    missingCell.setValue('Missing/Invalid Deadlines:');
    missingCell.offset(0, 1).setValue(missingDeadlinesCount).setNumberFormat('0').setFontWeight('bold');
    missingCell.setFontWeight('bold');

    applyDashboardFormatting(dashboardSheet, numDataRows, config);
  }
}

function setDashboardHeaderNotes(sheet, config) {
  const DL = config.DASHBOARD_LAYOUT;
  sheet.getRange(1, DL.TOTAL_COL).setNote('Total projects with a deadline in this month.');
  sheet.getRange(1, DL.UPCOMING_COL).setNote('Active projects ("In Progress" or "Scheduled") with a deadline on or after today.');
  sheet.getRange(1, DL.OVERDUE_COL).setNote('Active projects ("In Progress" or "Scheduled") with a deadline in the past. Click to see details.');
  sheet.getRange(1, DL.APPROVED_COL).setNote("Projects with 'Permits' status set to 'approved'.");
  sheet.getRange(1, DL.GT_TOTAL_COL).setNote('Grand total for all projects shown in the table.');
}

// ... [The rest of the file remains unchanged, so it is omitted for brevity] ...
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

function ensureHiddenColumnCapacity(sheet, startCol, columnsNeeded) {
  const requiredEndCol = startCol + columnsNeeded - 1;
  const currentMaxCol = sheet.getMaxColumns();

  if (currentMaxCol < requiredEndCol) {
    const toAdd = requiredEndCol - currentMaxCol;
    sheet.insertColumnsAfter(currentMaxCol, toAdd);
    Logger.log('Added ' + toAdd + ' columns to accommodate chart data.');
  }

  sheet.hideColumns(startCol, columnsNeeded);
}

function ensureRowCapacity(sheet, minRows) {
  const currentMaxRows = sheet.getMaxRows();
  if (currentMaxRows < minRows) {
    sheet.insertRowsAfter(currentMaxRows, minRows - currentMaxRows);
    Logger.log('Added ' + (minRows - currentMaxRows) + ' rows to accommodate chart data.');
  }
}

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

function assertCondition(condition, message) {
  if (!condition) Logger.log('ASSERT: ' + message);
}

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
      Logger.log('No data available for charts.');
      displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'No project data available to chart.');
      return;
    }

    var HIDDEN_START_COL = DL.HIDE_COL_START;
    var PAST_COL = HIDDEN_START_COL;
    var UPC_COL  = HIDDEN_START_COL + 4;
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
    var prevUpc  = getStoredCount(sheet, UPC_COL);
    var rowsToClear = Math.max(pastData.length, upcomingData.length, prevPast, prevUpc) + 2;

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

    setStoredCount(sheet, PAST_COL, pastData.length);
    setStoredCount(sheet, UPC_COL,  upcomingData.length);

    var buildChart = function(title, leftCol, rowsCount, anchorRow) {
      if (rowsCount <= 0) return null;
      var range = sheet.getRange(DATA_START_ROW, leftCol, rowsCount + 1, 4);
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

    if (pastData.length > 0) {
      var c1 = buildChart('Past ' + pastData.length + ' Months: Overdue, Upcoming, Total', PAST_COL, pastData.length, DL.CHART_START_ROW);
      if (c1) sheet.insertChart(c1);
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'No project data found for the past ' + DC.PAST_MONTHS_COUNT + ' months.');
    }

    if (upcomingData.length > 0) {
      var c2 = buildChart('Next ' + upcomingData.length + ' Months: Overdue, Upcoming, Total', UPC_COL, upcomingData.length, DL.CHART_START_ROW + DC.ROW_SPACING);
      if (c2) sheet.insertChart(c2);
    } else {
      displayChartPlaceholder(sheet, DL.CHART_START_ROW + DC.ROW_SPACING, DL.CHART_ANCHOR_COL, 'No project data found for the next ' + DC.UPCOMING_MONTHS_COUNT + ' months.');
    }

  } catch (error) {
    Logger.log('ERROR in createOrUpdateDashboardCharts: ' + error.message + '\n' + error.stack);
    displayChartPlaceholder(sheet, DL.CHART_START_ROW, DL.CHART_ANCHOR_COL, 'Chart creation failed. Check logs for details.');
  }
}

function getMonthStart_(d) {
  var x = new Date(d);
  x.setDate(1);
  x.setHours(0, 0, 0, 0);
  return x;
}

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