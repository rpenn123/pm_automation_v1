### Dashboard Overdue Logic and Data Flow

This document outlines the updated logic for calculating overdue projects and how data flows from the `Forecasting` sheet to the `Dashboard`.

#### Overdue Project Definition

A project from the **Forecasting** sheet is considered **Overdue** if it meets both of the following criteria:

1.  **Deadline:** The date in the `Deadline` column (J) is on or before today's date. The comparison is performed in the `America/New_York` timezone, normalizing both dates to midnight to ensure accuracy.
2.  **Status:** The project's status in the `Progress` column (G) is **not** one of the following terminal or non-active statuses:
    *   `Done`
    *   `Canceled`
    *   `On Hold`
    *   `Stuck`
    *   `Completed` (from config)
    *   `Cancelled` (from config)

Projects with empty or invalid deadline cells are ignored.

#### Data Flow and Hover Notes

The dashboard generation process is orchestrated by the `updateDashboard` function in `src/ui/Dashboard.gs` and follows these steps:

1.  **Read Data:** The script reads all project data from the `Forecasting` sheet.
2.  **Process Data:** The `processDashboardData` function iterates through each project, calculating monthly summaries for `Total`, `Upcoming`, `Overdue`, and `Approved` projects based on the rules defined in the script.
3.  **Build Hover Notes:** During processing, if a project is identified as overdue, its name and deadline are collected. These details are grouped by month.
4.  **Render Table:** The `renderDashboardTable` function populates the main dashboard grid. For each month, it sets the overdue count in **Column E**. It then generates a hover note for that cell, containing a list of the contributing overdue projects (`Project Name — Deadline`). If a month has more than 20 overdue projects, the note will show the first 20 followed by a `(+N more)` line.
5.  **Delete Old Sheet:** The script now programmatically deletes the "Overdue Details" sheet, as it has been fully replaced by the hover notes.

#### How to Run the Verification Function

A lightweight verification test has been added to ensure the dashboard's overdue totals are accurate.

1.  Open the Google Apps Script editor for the project.
2.  From the function dropdown list, select `runDashboardVerification`.
3.  Click the **Run** button.
4.  Open the execution logs by navigating to **View > Logs** (or `Ctrl+Enter` / `Cmd+Enter`).
5.  The logs will show the result of the verification, including a check of the grand total and a spot-check of a monthly total against its hover note. A success message indicates that the dashboard's totals match the direct calculation.