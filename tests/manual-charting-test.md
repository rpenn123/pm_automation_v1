# Testing Plan for Dashboard Charting

This document outlines the testing plan to verify the refactored dashboard charting functionality. It includes sample data and expected outcomes.

**Assumptions:**
- The current date is around **late September 2025**.
- The `CONFIG.DASHBOARD_CHARTING` settings are:
  - `PAST_MONTHS_COUNT: 3`
  - `UPCOMING_MONTHS_COUNT: 6`

---

## 1. Sample Data for "Forecasting" Sheet

To test the charting logic, clear the existing data in the "Forecasting" sheet (from row 2 downwards) and insert the following 3 rows.

**Important:** The column letters correspond to the default layout. Ensure they match your `CONFIG.FORECASTING_COLS`. The key columns for this test are `PROJECT_NAME` (A), `PROGRESS` (F), and `DEADLINE` (I).

| A (Project Name) | ... | F (Progress) | ... | I (Deadline) |
| :--------------- | :-- | :----------- | :-- | :----------- |
| Project Alpha    | ... | In Progress  | ... | 8/15/2025    |
| Project Bravo    | ... | Scheduled    | ... | 11/20/2025   |
| Project Charlie  | ... | In Progress  | ... | 1/10/2025    |


---

## 2. Test Execution Steps

1.  **Prepare the Sheet**: Navigate to the "Forecasting" sheet in your Google Sheet.
2.  **Clear Old Data**: Delete all data rows below the header row (i.e., from row 2 to the bottom).
3.  **Insert Sample Data**: Copy and paste the 3 rows from the table above into the "Forecasting" sheet, starting at row 2.
4.  **Run the Script**: From the Google Apps Script editor, select the `updateDashboard` function from the dropdown menu and click **Run**.
5.  **Observe the Dashboard**: Once the script finishes, switch to the "Dashboard" sheet to observe the results.

---

## 3. Expected Outcomes

Based on the sample data and assuming a run date in late September 2025, here is what you should see on the "Dashboard" sheet:

### Scenario A: Both Charts Generated Successfully

This is the primary test case using the provided sample data.

*   **"Past Months" Chart**:
    *   A column chart titled something like "Past 3 Months: Overdue vs. Total" should be visible starting around row 2, column K.
    *   It should display data for **August 2025**, showing **1 Overdue** project (Project Alpha) and **1 Total** project for that month.

*   **"Upcoming Months" Chart**:
    *   A second column chart titled something like "Next 6 Months: Upcoming vs. Total" should be visible below the first chart (around row 17).
    *   It should display data for **November 2025**, showing **1 Upcoming** project (Project Bravo) and **1 Total** project for that month.

*   **Data Table**:
    *   The main data table will show rows for all months, but only August and November will have non-zero values based on this data. Project Charlie's deadline (Jan 2025) is outside the 3-month past window and will not generate a chart point, but it will be counted in the grand totals.

### Scenario B: Test for Placeholder Messages (No Data)

This test verifies the user feedback when no relevant data is found.

1.  **Execution**: Follow the "Test Execution Steps" above, but **do not add any data** to the "Forecasting" sheet after clearing it. Or, for a more robust test, use only "Project Charlie" from the sample data, as its deadline is outside the chart's date range.
2.  **Run `updateDashboard`**.

*   **Expected Outcome**:
    *   **No charts** should be generated.
    *   Where the "Past Months" chart would be, you should see a merged cell with the italicized message: *"No project data found for the past 3 months."*
    *   Where the "Upcoming Months" chart would be, you should see a merged cell with the italicized message: *"No project data found for the next 6 months."*

---