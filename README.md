# Google Sheets Automation & Dashboard

This repository contains the Google Apps Script (GAS) codebase for a powerful project management automation and dashboard system. It is designed to be synced with two Google Sheets (TEST and PROD) and includes a streamlined deployment process using `@google/clasp` and one-click Windows batch files.

## 1. Core Features

This system provides a suite of features to automate project tracking and visualize progress:

*   **Two-Way Progress Sync:** Automatically synchronizes the `Progress` status of a project between the `Forecasting` and `Upcoming` sheets.
*   **Automated Row Transfers:**
    *   Moves projects from `Forecasting` to `Upcoming` when their `Permits` status is marked as "approved".
    *   Moves projects from `Forecasting` to `Inventory_Elevators` when their `Delivered` status is set to `TRUE`.
    *   Copies projects from `Forecasting` to `Framing` when their `Progress` status is set to "In Progress".
*   **Automated Tracking and Logging:**
    *   **Last Edit Time:** Automatically records and displays a human-readable timestamp (e.g., "5 minutes ago") for the last modification to any row in a tracked sheet.
    *   **External Audit Logs:** Maintains a detailed, monthly log of all automated actions in a separate Google Spreadsheet for auditing and debugging.
    *   **Email Error Alerts:** Sends detailed email notifications to a designated address upon encountering any script errors.
*   **Dynamic Dashboard:**
    *   A menu-driven dashboard that provides a high-level overview of project statuses.
    *   Includes monthly summaries, grand totals, and a clickable drill-down view for overdue projects.
    *   Features auto-generated charts to visualize past and upcoming project loads.

## 2. System Architecture

The system is built on the following components:

*   **Google Apps Script (GAS):** The core logic is written in Apps Script, a JavaScript-based language that runs on Google's servers and can interact with Google Workspace applications.
*   **Google Sheets:** The system is bound to two Google Sheets that serve as the database and user interface.
*   **`clasp`:** This is a command-line tool from Google that allows for local development of Apps Script projects. It handles pulling and pushing code between the local repository and the scripts bound to the Google Sheets.

## 3. Ops & Deployment

This section covers the environments, prerequisites, and deployment procedures.

### Environments

The system operates in two distinct environments, each with its own Google Sheet and bound Apps Script project.

*   **TEST Environment:**
    *   **Purpose:** For development, testing new features, and verifying changes before they go live.
    *   **Script ID:** `1vYXETLX8I3HICSveg7FmLKWbZiwToKicThNeOxm_maQgnQ97AwDmz7iX`
*   **PRODUCTION Environment:**
    *   **Purpose:** The live system used for daily operations.
    *   **Script ID:** `15_PrYM6MxfCbA1bXt0deEGI7cXf74B_KlJY7Ydw59uVrmHZn3IEKFGPJ`

> **Deployment Control:** The active deployment target is determined by the contents of the `.clasp.json` file at the root of the repository. The deployment scripts (`update_test.bat`, `update_production.bat`) work by overwriting this file with the correct configuration from either `.clasp.test.json` or `.clasp.prod.json` before pushing the code.

### Prerequisites

Before you can work with this repository, you need the following:

*   **Node.js:** The latest Long-Term Support (LTS) version is recommended.
*   **`@google/clasp`:** Install it globally via npm by running: `npm i -g @google/clasp`.
*   **`clasp` Login:** Authenticate `clasp` with your Google account by running: `clasp login`.
*   **Git:** For version control.
*   **Google Account:** You must have editor access to both the TEST and PROD Google Sheets.
*   **Google Apps Script API:** You must enable the Apps Script API for your Google account. You can do so [here](https://script.google.com/home/usersettings).

### Deployment Process

The repository includes simple batch scripts for one-click deployments on Windows.

1.  **Deploy to TEST:**
    *   Double-click the `update_test.bat` file.
    *   In a terminal like PowerShell, you can run: `.\update_test.bat`.
    *   This script will update `.clasp.json` with the TEST configuration and push the `src` directory to the TEST Apps Script project.

2.  **Deploy to PRODUCTION:**
    *   Double-click the `update_production.bat` file.
    *   This script will update `.clasp.json` with the PROD configuration and push the `src` directory to the PROD Apps Script project.

### Post-Deployment Steps

After every deployment, it is crucial to perform the following steps in the target Google Sheet:

1.  **Refresh the Sheet:** Reload the Google Sheet in your browser.
2.  **Run Full Setup (If Necessary):** If you have made changes that add new services or require new permissions (e.g., adding external logging for the first time), you must run the initial setup.
    *   Go to the custom menu: **🚀 Project Actions → Run Full Setup (Install Triggers & Logging)**.
    *   This ensures that the `onEdit` trigger is correctly installed and all necessary initializations are complete.

## 4. Development Workflow

### Daily Use & Smoke Test

To ensure a deployment was successful, perform the following 2-minute smoke test:

1.  **Sync Test:** In the `Forecasting` sheet, edit the `Progress` for a project that also exists in the `Upcoming` sheet. Verify that the change is reflected in `Upcoming`.
2.  **Permit Transfer Test:** In the `Forecasting` sheet, set the `Permits` status for a project to `approved`. Verify that the project row is correctly transferred to the `Upcoming` sheet.
3.  **Delivery Transfer Test:** In the `Forecasting` sheet, set the `Delivered` status for a project to `TRUE`. Verify that the project is transferred to the `Inventory_Elevators` sheet.
4.  **Framing Transfer Test:** In the `Forecasting` sheet, set the `Progress` status for a project to `In Progress`. Verify that the project is copied to the `Framing` sheet.
5.  **Check Executions:** Open the Apps Script editor (**Extensions → Apps Script**) and go to the **Executions** tab to check for any failed runs.

### Codebase Overview

The core logic is located in the `src/` directory.

*   `Config.gs`: A centralized configuration file for all settings, including sheet names, column mappings, and status strings. **All application settings should be changed here.**
*   `Automations.gs`: Contains the main `onEdit` trigger, which routes edits to the appropriate handler functions. It also defines the logic for data synchronization and the configurations for data transfers.
*   `TransferEngine.gs`: A generic, reusable engine that executes the data transfers based on the configurations provided by `Automations.gs`. It handles locking, duplicate checking, and data mapping.
*   `Dashboard.gs`: Contains all the logic for generating the dynamic dashboard, including data processing, formatting, and chart creation.
*   `LastEditService.gs`: Manages the "Last Edit" tracking columns, including their creation and automatic updates.
*   `LoggerService.gs`: Handles the external audit logging to a separate spreadsheet and the email notifications for errors.
*   `Setup.gs`: Manages the creation of the custom UI menu (`onOpen`) and the one-time setup routine for installing triggers and initializing services.
*   `Utilities.gs`: A collection of shared helper functions used throughout the project for tasks like data normalization, lookups, and date formatting.
*   `BugFixTest.gs`: Contains a unit test for a specific bug fix related to the `TransferEngine.gs` to ensure regressions do not occur.

### Code Documentation

All source files in the `src/` directory are documented using JSDoc style comments. Each function has a block comment explaining its purpose, parameters, and return values. This documentation is intended to make the codebase easier to understand and maintain.

## 5. Continuous Integration (CI)

This repository has a GitHub Actions workflow (`.github/workflows/validate-deploy.yml`) that runs on every commit and pull request. This CI check validates the deployment configuration files (`.clasp.*.json`) to prevent common deployment errors.

## 6. Rollback Procedure

To roll back to a previous version:

1.  Check out a known-good commit from the Git history.
2.  Deploy to the **TEST** environment and perform the smoke test to validate.
3.  Once validated, deploy to the **PROD** environment.

## 7. Recommended Hardening

The following are recommendations for improving the security and stability of the repository:

1.  **Branch Protection:** Enable branch protection on the `main` branch to require pull requests and successful CI checks before merging.
2.  **CODEOWNERS:** Add a `CODEOWNERS` file to automatically request reviews from designated owners on pull requests.
3.  **Release Checklist:** Create a `docs/RELEASE.md` file with the smoke test checklist to ensure consistent validation for every release.

---
*Owner: Ryan (rpenn@mobility123.com)*