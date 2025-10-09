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
-   **Dynamic Dashboard:** A comprehensive, auto-generating dashboard that summarizes project data by month, including totals, upcoming deadlines, and overdue items.
-   **"Last Edit" Tracking:** Automatically records and displays a human-readable timestamp (e.g., "5 min. ago") for the last modification to any row in a tracked sheet.
-   **Robust Audit Logging:** Captures every significant action (e.g., data transfer, sync, error) in a separate, dedicated log spreadsheet, organized by month.
-   **Error Notifications:** Automatically sends detailed email alerts to a configured address upon encountering a critical error.

## 3. How It Works

The system is event-driven and modular, built around a few key components:

1.  **The `onEdit` Trigger (`src/core/Automations.gs`):** This is the heart of the system. Every time a user edits a cell in the spreadsheet, this trigger fires. It acts as a router, checking if the edit occurred in a relevant location.

2.  **Rule-Based Logic (`src/core/Automations.gs`):** The `onEdit` function uses a `rules` array to determine what to do. Each rule specifies a sheet, a column, and a handler function. If an edit matches a rule's conditions, the corresponding handler is executed. This makes the system easy to extend with new automations.

3.  **The Transfer Engine (`src/core/TransferEngine.gs`):** This is a generic, reusable engine that handles all data transfers between sheets. It takes a configuration object that defines the source, destination, column mappings, and duplicate-checking logic, making it highly flexible.

4.  **Services (`src/services/`):**
    -   **`LoggerService.gs`:** Manages all audit logging to an external spreadsheet and sends formatted error notification emails.
    -   **`LastEditService.gs`:** Manages the "Last Edit" columns, updating the timestamp and relative time formula whenever a row is modified.

5.  **The Dashboard (`src/ui/Dashboard.gs`):** This script reads data from the `Forecasting` sheet, performs all the necessary calculations and aggregations in memory, and then renders the final tables and charts on the `Dashboard` sheet.

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
2.  A new menu named **"🚀 Project Actions"** should appear in the menu bar.
3.  Navigate to **🚀 Project Actions > ⚙️ Setup & Configuration > Run Full Setup (Install Triggers)**.
4.  Follow the authorization prompts. This step:
    -   Installs the necessary `onEdit` trigger that powers all automations.
    -   Creates the "Last Edit" tracking columns on all relevant sheets.
    -   Initializes the external audit logging system.
5.  Next, navigate to **🚀 Project Actions > ⚙️ Setup & Configuration > Set Error Notification Email** and enter the email address where you wish to receive error alerts.

## 5. Project Structure and Documentation

The repository is organized to separate concerns, making it easier to navigate and maintain. All functions are documented using JSDoc-style comments.

```
.
├── config/       # Environment-specific .clasp.json files
├── docs/         # Additional project documentation
├── scripts/      # Node.js scripts for deployment and validation
├── src/          # Google Apps Script source code
│   ├── core/     # Core automation logic
│   │   ├── Automations.gs     # Main onEdit trigger and rule-based routing.
│   │   ├── TransferEngine.gs  # Generic engine for all data transfers.
│   │   └── Utilities.gs       # Shared helper functions (normalization, lookups, etc.).
│   ├── services/ # Background services
│   │   ├── LastEditService.gs # Manages "Last Edit" timestamp columns.
│   │   └── LoggerService.gs   # Handles audit logging and error email notifications.
│   └── ui/       # User-facing elements
│   │   ├── Dashboard.gs       # Logic for generating the dashboard sheet.
│   │   └── Setup.gs           # Creates the custom menu (onOpen) and setup routines.
│   └── Config.gs # Central configuration for the entire application.
├── tests/        # Manual testing plans
├── package.json  # Defines scripts and dependencies
└── README.md     # This file
```

## 6. Configuration

The entire application is controlled by the `CONFIG` object in **`src/Config.gs`**. This file is the central hub for all settings. To modify the script's behavior, you will likely need to edit this file.

Key areas include:
-   `SHEETS`: The names of all sheets used in the automation.
-   `STATUS_STRINGS`: The text values for statuses that trigger actions (e.g., "In Progress", "approved").
-   `*_COLS`: The 1-indexed column numbers for all data fields in each sheet. If you add, remove, or move a column in your sheet, you **must** update the corresponding mapping here.
-   `DASHBOARD_LAYOUT` & `DASHBOARD_FORMATTING`: Defines the structure and appearance of the Dashboard sheet.

**Example:** If you move the 'Deadline' column in the 'Forecasting' sheet from column J (10) to column K (11), you would update `src/Config.gs` as follows:
```javascript
// Before
DEADLINE: 10, // J

// After
DEADLINE: 11, // K
```

## 7. Testing

### Manual Smoke Testing
A detailed manual test plan is located in `tests/smoke-test.md`. This should be run against the TEST environment after every deployment to verify core functionality.

### Automated Verification
The script includes a built-in verification function to audit the Dashboard's calculations. To run it:
1. Open the Apps Script editor in your spreadsheet.
2. From the function dropdown, select `runDashboardVerification` and click **Run**.
3. Check the **Execution Log** at the bottom of the screen to see a report comparing the dashboard's totals against a direct recalculation from the source data.

## 8. CI/Deployment

-   **CI:** A GitHub Actions workflow runs on every pull request to `main`. It executes `npm run validate-config` to ensure that the `clasp` configuration files in `config/` are valid.
-   **Deployment:** Deployments are handled via the `npm run deploy:*` commands as described in Step 2. This process is designed to be run from a developer's local machine.

---
*Owner: Ryan (rpenn@mobility123.com)*