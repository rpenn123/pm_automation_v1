# Manual Smoke Test Plan

This document provides a step-by-step manual testing plan to verify the core functionality of the project management automation system. It should be executed in the **TEST** environment after every deployment.

## 1. Prerequisites

Before starting, ensure you have a clean state for testing:

-   You have at least one project row in the `Forecasting` sheet that does **not** yet exist in the `Upcoming`, `Inventory_Elevators`, or `Framing` sheets.
-   This test project should have a unique `Project Name` and, ideally, a unique `SFID`.
-   For the sync test, ensure you have one project that exists in **both** the `Forecasting` and `Upcoming` sheets.

**Example Test Row in `Forecasting`:**

| SFID | Project Name | Progress | Permits | Delivered | ... |
| :--- | :----------- | :------- | :------ | :-------- | :-- |
| SMOKE-001 | Smoke Test Project | Scheduled | Pending | `FALSE` | ... |

---

## 2. Test Cases

Execute the following test cases in order. After each edit, wait a few moments for the script to run, then check the relevant sheets for the expected outcome.

### Test Case 1: Permit Transfer (`Forecasting` → `Upcoming`)

1.  **Action:** In the `Forecasting` sheet, find your test project row. Change the value in the `Permits` column to `approved`.
2.  **Verification:**
    -   Navigate to the `Upcoming` sheet.
    -   **Expected Result:** The "Smoke Test Project" row has been copied to the `Upcoming` sheet with all the correctly mapped columns (Project Name, Deadline, etc.).

### Test Case 2: Delivery Transfer (`Forecasting` → `Inventory`)

1.  **Action:** In the `Forecasting` sheet, find your test project row. Check the box in the `Delivered` column (setting its value to `TRUE`).
2.  **Verification:**
    -   Navigate to the `Inventory_Elevators` sheet.
    -   **Expected Result:** The "Smoke Test Project" row has been copied to the `Inventory_Elevators` sheet.

### Test Case 3: Framing Transfer (`Forecasting` → `Framing`)

1.  **Action:** In the `Forecasting` sheet, find your test project row. Change the value in the `Progress` column to `In Progress`.
2.  **Verification:**
    -   Navigate to the `Framing` sheet.
    -   **Expected Result:** The "Smoke Test Project" row has been copied to the `Framing` sheet.

### Test Case 4: Bi-Directional Progress Sync

This test requires a project that exists on **both** the `Forecasting` and `Upcoming` sheets.

1.  **Action (Part A):** In the `Forecasting` sheet, change the `Progress` of the synced project to a new value (e.g., "Site Prep").
2.  **Verification (Part A):**
    -   Navigate to the `Upcoming` sheet.
    -   **Expected Result:** The `Progress` column for that same project has automatically updated to "Site Prep".

3.  **Action (Part B):** In the `Upcoming` sheet, change the `Progress` back to another value (e.g., "Scheduled").
4.  **Verification (Part B):**
    -   Navigate back to the `Forecasting` sheet.
    -   **Expected Result:** The `Progress` column for that project has automatically updated back to "Scheduled".

### Test Case 5: "Last Edit" Tracking

1.  **Action:** Perform any of the data-modifying actions from the test cases above.
2.  **Verification:**
    -   In the sheet where you made the edit, scroll to the `Last Edit` column for the row you modified.
    -   **Expected Result:** The cell should display a fresh, relative timestamp, such as "just now" or "1 min. ago".

### Test Case 6: Audit Logging

1.  **Action:** After completing all other test cases, open the external log spreadsheet. (You can get the link from the script's properties or by running the setup routine again).
2.  **Verification:**
    -   Navigate to the sheet for the current month (e.g., `2025-10`).
    -   **Expected Result:** You should see new rows at the top of the log corresponding to each action you performed (e.g., `SyncFtoU`, `Upcoming Transfer (Permits)`). The `Result` column for these entries should be `success` or another expected outcome like `skipped-duplicate`.

---

If all test cases pass, the core functionality of the system is verified.