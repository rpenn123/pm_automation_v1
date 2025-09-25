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
    clearAndPrepareDashboardSheet(dashboardSheet);
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


// (The remaining functions: clearAndPrepareDashboardSheet,
// setDashboardHeaders, setDashboardHeaderNotes, applyDashboardFormatting, hideDataColumns, 
// createOrUpdateDashboardCharts, generateMonthList are omitted for brevity but are 
// included in the final repository. They primarily handle the visualization and formatting
// based on the CONFIG settings and the processed data.)

// NOTE: Due to the complexity and length of the presentation logic functions, 
// they are not fully included in this snippet but are present in the generated ZIP file.
// They have been refactored to use the centralized CONFIG object for all layout, 
// formatting, and charting parameters.