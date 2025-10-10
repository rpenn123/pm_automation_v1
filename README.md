# Google Sheets Project Management Automation & Dashboard

[![Validate Deploy Configs](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml/badge.svg)](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml)

This repository contains the Google Apps Script (GAS) codebase for a powerful project management automation and dashboard system. It is designed to be tightly integrated with Google Sheets, providing a centralized platform for tracking project lifecycles from forecasting to completion.

## 1. Purpose of the Project

The primary purpose of this system is to automate repetitive data management tasks and provide clear, high-level visibility into project statuses. It solves common challenges in spreadsheet-based project management, such as:

-   Manually copying and pasting data between sheets.
-   Keeping project statuses in sync across different views.
-   Lack of a centralized dashboard for key performance indicators (KPIs).
-   Absence of a clear audit trail for data modifications.

By automating these processes, the system saves time, reduces human error, and empowers teams to make data-driven decisions more effectively.

## 2. Core Features

-   **Automated Data Transfers:** Automatically moves project rows between sheets based on status changes (e.g., from `Forecasting` to `Upcoming` when permits are approved).
-   **Bi-Directional Data Sync:** Keeps the `Progress` status synchronized between the `Forecasting` and `Upcoming` sheets, allowing updates from either sheet to be reflected in the other.
-   **Dynamic Dashboard:** A comprehensive, auto-generating dashboard that summarizes project data by month, including totals, upcoming deadlines, overdue items, and configurable charts.
-   **"Last Edit" Tracking:** Automatically records and displays a human-readable timestamp (e.g., "5 min. ago") for the last modification to any row in a tracked sheet.
-   **Robust Audit Logging:** Captures every significant action (e.g., data transfer, sync, error) in a separate, dedicated log spreadsheet, organized by month for easy review.
-   **Resilient Error Handling:** Includes a centralized error handler, a retry mechanism for transient network issues, and automatic email notifications for critical failures.

## 3. How It Works

The system is event-driven and modular, built around a few key components. Every significant operation is assigned a unique `correlationId` that is passed through all functions and logs, allowing for end-to-end tracing of an entire automation flow.

1.  **The `onEdit` Trigger (`src/core/Automations.gs`):** This is the heart of the system. Every time a user edits a cell in the spreadsheet, this trigger fires. It acts as a router, checking if the edit occurred in a relevant location (e.g., the 'Progress' column of the 'Forecasting' sheet).

2.  **Rule-Based Logic (`src/core/Automations.gs`):** The `onEdit` function uses a `rules` array to determine what to do. Each rule is an object specifying a sheet, a column, an optional value to check for, and a handler function. If an edit matches a rule's conditions, the corresponding handler is executed. This design makes the system highly extensible.

3.  **The Transfer Engine (`src/core/TransferEngine.gs`):** This is a generic, reusable engine that handles all data transfers between sheets. It takes a detailed configuration object that defines the source, destination, column mappings, and sophisticated duplicate-checking logic, making it extremely flexible.

4.  **Services (`src/services/`):**
    -   **`LoggerService.gs`:** Manages all audit logging to an external spreadsheet and sends formatted error notification emails. The logging is self-healing; if the log spreadsheet is deleted, it will be automatically recreated.
    -   **`LastEditService.gs`:** Manages the "Last Edit" columns, updating the timestamp and relative time formula whenever a row is modified.
    -   **`ErrorService.gs`:** Provides custom error types and a centralized handler (`handleError`) that standardizes how errors are logged and reported.

5.  **The Dashboard (`src/ui/Dashboard.gs`):** This script reads data from the `Forecasting` sheet, performs all the necessary calculations and aggregations in memory, and then renders the final tables and charts on the `Dashboard` sheet. This process is idempotent, meaning running it multiple times will produce the same result without creating duplicate data.

### Dashboard Overdue Logic

The dashboard provides a high-level summary of project timelines. A project from the **Forecasting** sheet is considered **Overdue** if it meets **both** of the following criteria:

1.  **Deadline:** The date in the `Deadline` column is on or before today's date.
2.  **Status:** The project's status in the `Progress` column is **exactly** "In Progress".

Projects with any other status (e.g., "Scheduled", "Completed", "Cancelled") are **not** counted as overdue, regardless of their deadline.

## 4. Setup and Usage

Follow these steps to get the project cloned, deployed, and running.

### Step 1: System Prerequisites

-   **Node.js:** The latest Long-Term Support (LTS) version is recommended.
-   **Git:** For version control.
-   **`clasp` Login:** Authenticate `clasp` with your Google account by running `npx clasp login`. You only need to do this once.
-   **Google Apps Script API:** Ensure the API is enabled for your Google account. You can do so [here](https://script.google.com/home/usersettings).

### Step 2: Installation & Deployment

1.  **Clone the repository:**
    ```sh
    git clone https://github.com/rpenn123/pm_automation_v1.git
    cd pm_automation_v1
    ```

2.  **Install dependencies:**
    ```sh
    npm install
    ```

3.  **Deploy to an environment:**
    -   **To TEST:** `npm run deploy:test`
    -   **To PRODUCTION:** `npm run deploy:prod`

    These commands copy the correct `.clasp.[env].json` configuration from the `config/` directory and push the `src/` folder to the corresponding Apps Script project.

### Step 3: First-Time In-Sheet Setup

**This is a mandatory one-time setup inside your Google Sheet.**

1.  After a successful deployment, open the Google Sheet linked to your Apps Script project.
2.  A new menu named **"🚀 Project Actions"** should appear in the menu bar. If it doesn't appear after a minute, try reloading the page.
3.  Navigate to **🚀 Project Actions > ⚙️ Setup & Configuration > Run Full Setup (Install Triggers)**.
4.  A dialog box will appear asking for authorization. Click "Continue" and select your Google account. You will see a warning that "Google hasn't verified this app." Click "Advanced," then "Go to [Project Name] (unsafe)." Grant the requested permissions. This is necessary for the script to manage triggers and edit sheets on your behalf.
5.  The setup script will then:
    -   Install the necessary `onEdit` trigger that powers all automations.
    -   Create the "Last Edit" tracking columns on all relevant sheets.
    -   Initialize the external audit logging system.
6.  Next, navigate to **🚀 Project Actions > ⚙️ Setup & Configuration > Set Error Notification Email** and enter the email address where you wish to receive error alerts.

## 5. Project Structure

The repository is organized to separate concerns, making it easier to navigate and maintain. All functions are documented using JSDoc-style comments.

```
.
├── config/       # Environment-specific .clasp.json files (DO NOT COMMIT SENSITIVE IDs).
├── docs/         # Additional project documentation (e.g., RFCs).
├── scripts/      # Node.js scripts for deployment and validation.
├── src/          # Google Apps Script source code.
│   ├── core/     # Core automation logic.
│   │   ├── Automations.gs     # Main onEdit trigger and rule-based routing.
│   │   ├── TransferEngine.gs  # Generic engine for all data transfers.
│   │   └── Utilities.gs       # Shared helper functions (normalization, lookups, etc.).
│   ├── services/ # Background services.
│   │   ├── ErrorService.gs    # Custom error classes and centralized handler.
│   │   ├── LastEditService.gs # Manages "Last Edit" timestamp columns.
│   │   └── LoggerService.gs   # Handles audit logging and error email notifications.
│   └── ui/       # User-facing elements.
│   │   ├── Dashboard.gs       # Logic for generating the dashboard sheet.
│   │   └── Setup.gs           # Creates the custom menu (onOpen) and setup routines.
│   └── Config.gs # Central configuration for the entire application.
├── tests/        # Manual test plans and local unit test suites.
├── .gitignore    # Specifies files for git to ignore.
├── package.json  # Defines scripts and dependencies.
└── README.md     # This file.
```

## 6. Configuration

The entire application is controlled by the `CONFIG` object in **`src/Config.gs`**. This file is the single source of truth for all settings. To modify the script's behavior, you will likely need to edit this file.

Key areas include:
-   `SHEETS`: The names of all sheets used in the automation.
-   `STATUS_STRINGS`: The text values for statuses that trigger actions (e.g., "In Progress", "approved"). **Note:** These are case-insensitive in the script logic.
-   `*_COLS`: The 1-indexed column numbers for all data fields in each sheet. If you add, remove, or move a column in your sheet, you **must** update the corresponding mapping here.
-   `LAST_EDIT.TRACKED_SHEETS`: An array of sheet names that should have the "Last Edit" feature.
-   `DASHBOARD_LAYOUT` & `DASHBOARD_FORMATTING`: Defines the structure and appearance of the Dashboard sheet.

**Example:** If you move the 'Deadline' column in the 'Forecasting' sheet from column J (10) to column K (11), you would update `src/Config.gs` as follows:
```javascript
// Before
DEADLINE: 10, // J

// After
DEADLINE: 11, // K
```

## 7. Testing

### Local Unit Testing
The project includes a lightweight, custom unit testing framework that allows you to run tests for your `.gs` files locally using Node.js, without needing to deploy them. This provides a much faster feedback loop for development.

**How to Run the Tests:**
From the root of the project directory, run:
```sh
node run_test.js
```
The script will execute all test suites defined in `tests/` and log the results to the console.

**Adding a New Test File:**
To add a new test suite (e.g., `tests/my_new_feature.test.gs`):
1.  Create your new test file in `tests/`. It should contain one or more test functions and a main runner function (e.g., `runMyNewFeatureTests()`).
2.  Open `run_test.js` and make three changes:
    a.  **Load the file:** Add a new `fs.readFileSync` line to load your new test file.
    b.  **Evaluate the file:** Add an `eval()` line for your new test file's content.
    c.  **Call the runner:** Add a call to your main test runner function inside the `try` block at the end of the script.

### Manual Smoke Testing
A detailed manual test plan is located in `tests/smoke-test.md`. This should be run against the TEST environment after every deployment to verify core functionality.

### Built-in Verification
The script includes a built-in verification function to audit the Dashboard's calculations against the source data.
To run it:
1.  Open the Apps Script editor in your spreadsheet (**Extensions > Apps Script**).
2.  From the function dropdown list, select `runDashboardVerification`.
3.  Click **Run**.
4.  View the results in the **Execution Log** (`Ctrl+Enter` or `Cmd+Enter`).

## 8. Troubleshooting

-   **"🚀 Project Actions" menu does not appear:**
    -   Wait a minute and reload the Google Sheet. Simple triggers can sometimes be delayed.
    -   Ensure you have run a successful deployment (`npm run deploy:test`).
    -   Check the Apps Script editor for any errors that may have occurred during the `onOpen` execution.

-   **Automations are not working:**
    -   Verify that you have run the **"Run Full Setup"** menu item to install the `onEdit` trigger. You can check existing triggers under **"Triggers"** (the clock icon) in the Apps Script editor. There should only be one `onEdit` trigger.
    -   Check that the status strings in your sheet (e.g., "In Progress") exactly match the values in `src/Config.gs` (though the check is case-insensitive).
    -   View the **Executions** log in the Apps Script editor to see if the `onEdit` function is running and if it's throwing any errors.

-   **Dashboard is not updating or shows errors:**
    -   Run the update manually from the menu: **🚀 Project Actions > Update Dashboard Now**.
    -   Check the `Forecasting` sheet for any data that might be in an unexpected format, especially in the `Deadline` column.
    -   Check the logs for detailed error messages.

-   **Deployment with `clasp` fails:**
    -   Ensure you are logged in. Run `npx clasp login`.
    -   Ensure the Google Apps Script API is enabled for your account (see Step 1).
    -   Check that your `config/.clasp.[env].json` files contain the correct `scriptId`.

---
*Owner: Ryan (rpenn@mobility123.com)*