# Google Sheets Project Management Automation & Dashboard

[![Validate Deploy Configs](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml/badge.svg)](https://github.com/rpenn123/pm_automation_v1/actions/workflows/validate-deploy.yml)

This repository contains the Google Apps Script (GAS) codebase for a powerful project management automation and dashboard system. It is designed to be tightly integrated with Google Sheets, providing a centralized platform for tracking project lifecycles from forecasting to completion.

The system automates the flow of data between different project-tracking sheets and provides a high-level visual dashboard for monitoring key metrics like project timelines, overdue tasks, and permit statuses.

## 1. Core Features

-   **Automated Data Transfers:** Automatically moves project rows between sheets based on status changes. For example, a project is moved from `Forecasting` to `Upcoming` when its permits are approved.
-   **Bi-Directional Data Sync:** Keeps the `Progress` status synchronized between the `Forecasting` and `Upcoming` sheets, allowing updates from either sheet to be reflected in the other.
-   **Dynamic Dashboard:** A comprehensive, auto-generating dashboard that summarizes project data by month, including totals, upcoming deadlines, and overdue items. It features clickable drill-downs for overdue projects.
-   **"Last Edit" Tracking:** Automatically records and displays a human-readable timestamp (e.g., "5 min. ago") for the last modification to any row in a tracked sheet, providing clear visibility into recent activity.
-   **Robust Audit Logging:** Captures every significant action (e.g., data transfer, sync, error) in a separate, dedicated log spreadsheet, organized by month for easy review and debugging.
-   **Error Notifications:** Automatically sends detailed email alerts to a configured address upon encountering a critical error.

## 2. Setup and Usage

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

## 3. Project Structure

The repository is organized to separate concerns, making it easier to navigate and maintain.

```
.
├── .github/      # CI workflows for validating clasp configs
├── config/       # Environment-specific .clasp.json files
├── docs/         # Additional project documentation
├── scripts/      # Node.js scripts for deployment and validation
├── src/          # Google Apps Script source code
│   ├── core/     # Core automation logic (onEdit router, TransferEngine, Utilities)
│   ├── services/ # Background services (LoggerService, LastEditService)
│   └── ui/       # User-facing elements (Dashboard generator, Setup menus)
├── tests/        # Manual testing plans
├── package.json  # Defines scripts and dependencies
└── README.md     # This file
```

## 4. Configuration

The entire application is controlled by the `CONFIG` object in **`src/Config.gs`**. This file is the central hub for all settings.

To modify the script's behavior, you will likely need to edit this file. Key areas include:
-   `SHEETS`: The names of all sheets used in the automation.
-   `STATUS_STRINGS`: The text values for statuses that trigger actions (e.g., "In Progress", "approved"). These now include
    keys for `COMPLETED` and `CANCELLED` so you can align dashboard terminology with your team's language without digging into
    business logic.
-   `*_COLS`: The 1-indexed column numbers for all data fields in each sheet. If you add, remove, or move a column in your sheet, you **must** update the corresponding mapping here.
-   `DASHBOARD_LAYOUT`: Defines the structure and appearance of the Dashboard sheet.

### Dashboard setup notes

The dashboard and automations both read directly from `CONFIG.STATUS_STRINGS`. If your sheet labels use different words for
completed or cancelled work, update the `COMPLETED` and `CANCELLED` values there so the charts and aggregations stay in sync
with on-sheet terminology.

## 5. Testing

Currently, testing is performed manually. The smoke test plan provides a consistent way to verify core functionality after a deployment.

-   **Smoke Test:** A detailed plan is located in `tests/smoke-test.md`. This should be run against the TEST environment after every deployment.

## 6. CI/Deployment

-   **CI:** A GitHub Actions workflow runs on every pull request to `main`. It executes `npm run validate-config` to ensure that the `clasp` configuration files in `config/` are valid and correctly formatted.
-   **Deployment:** Deployments are handled via the `npm run deploy:*` commands as described in Step 2. This process is designed to be run from a developer's local machine.

## 7. Contribution Guide

1.  Create a new feature branch from `main`.
2.  Make your code changes within the `src/` directory.
3.  If you modify deployment or validation logic, update the scripts in `scripts/`.
4.  Run `npm run deploy:test` to deploy your changes to the TEST environment.
5.  Perform the smoke test in `tests/smoke-test.md` to verify functionality.
6.  Open a pull request to `main`.

## 8. Rollback and Revert Procedure

If a deployment introduces a critical issue, the safest way to roll back is to revert the pull request that contained the breaking change.

1.  In GitHub, find and revert the PR.
2.  Check out the `main` branch locally and pull the revert commit.
3.  Deploy the stable version to the affected environment: `npm run deploy:prod`.

---
*Owner: Ryan (rpenn@mobility123.com)*