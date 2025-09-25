# Optimized Google Sheets Automation & Dashboard Repository

This repository contains the optimized, merged, and restructured Google Apps Script (GAS) code for managing project forecasting, automations, and dashboard reporting within a Google Sheet.

The architecture has been modernized to follow best practices, making it ideal for version control via GitHub and deployment using Google's `clasp` tool.

## Features

1.  **Data Automation (onEdit Triggers):**
    *   Two-way synchronization of "Progress" status between `Forecasting` and `Upcoming` sheets.
    *   Automatic data transfers from `Forecasting` to `Upcoming` (Permits="approved"), `Inventory_Elevators` (Delivered=TRUE), and `Framing` (Progress="In Progress").
    *   "Last Edit" tracking (hidden timestamp and visible relative time) on key sheets.
    *   Robust external audit logging (monthly logs in a separate workbook).
    *   Email notifications for critical script errors.
2.  **Dashboard Reporting (Menu Trigger):**
    *   High-performance summary dashboard (Total, Upcoming, Overdue, Approved projects).
    *   Clickable drill-down sheet for `Overdue Details`.
    *   Charts visualizing trends.
    *   Report on missing or invalid deadlines.

## Architecture Highlights

The project utilizes a modular architecture for improved maintainability. Each file in the `/src` directory has a specific responsibility:

-   `/src/Config.gs`: Centralized configuration object (Single Source of Truth) for all settings, including sheet names, column mappings, status strings, and dashboard layout properties.
-   `/src/Utilities.gs`: A collection of shared helper functions for common tasks like data normalization, date manipulation, sheet lookups, and object creation.
-   `/src/LoggerService.gs`: Encapsulated service for handling all logging. This includes sending detailed email notifications on critical errors and writing to a persistent, external audit log.
-   `/src/LastEditService.gs`: Manages the "Last Edit" feature, including the creation, update, and initialization of timestamp and relative-time columns on tracked sheets.
-   `/src/TransferEngine.gs`: A generic, reusable engine for executing configuration-based data transfers between sheets. It handles locking, duplicate checking, and post-transfer actions like sorting.
-   `/src/Automations.gs`: The core of the automation logic. Contains the main `onEdit` trigger handler, which routes edits to the appropriate sync or transfer functions based on a set of rules.
-   `/src/Dashboard.gs`: Contains all logic for generating the project dashboard, including data aggregation, summary calculations, chart creation, and formatting.
-   `/src/Setup.gs`: Handles user-facing setup tasks, including creating the custom UI menu (`onOpen`) and the one-time installation routine for triggers and logging.

## Developer Documentation

This project adheres to a high standard of code documentation. All functions across all `.gs` files are documented using **JSDoc**.

For a detailed understanding of any specific function, its parameters, return values, and purpose, please refer directly to the source code comments within the `/src` directory. The JSDoc comments provide a complete reference for developers looking to understand or extend the codebase.

## Setup Instructions (Using clasp)

These instructions guide you through deploying this repository to your Google Sheet.

### Prerequisites

1.  **Node.js and npm:** Install from [nodejs.org](https://nodejs.org/).
2.  **Google Clasp:** Install globally:
    ```bash
    npm install -g @google/clasp
    ```
3.  **Enable Google Apps Script API:** Go to [script.google.com/home/usersettings](https://script.google.com/home/usersettings) and enable the API.

### Deployment Steps

1.  **Download and Unzip:** Download this repository to your local machine.
2.  **Authenticate Clasp:** Navigate to the repository root in your terminal and log in:
    ```bash
    clasp login
    ```

3.  **Connect to Your Sheet:**
    *   Open your target Google Sheet.
    *   Go to `Extensions` > `Apps Script`.
    *   In the Apps Script editor, click `Project Settings` (gear icon).
    *   Copy the `Script ID`.

4.  **Configure `.clasp.json`:**
    *   Open the `.clasp.json` file in the repository root.
    *   Replace `"YOUR_SCRIPT_ID_HERE"` with the Script ID you copied.

    ```json
    {
        "scriptId":"PASTE_YOUR_ID_HERE",
        "rootDir": "./src"
    }
    ```

5.  **Deploy the Code:**
    Push the code to the Google Apps Script project. **Warning: This overwrites existing code in the target project.**

    ```bash
    clasp push
    ```

### Finalizing Setup in Google Sheets

1.  **Refresh the Google Sheet.**
2.  **Run Setup:**
    *   A new menu `🚀 Project Actions` will appear.
    *   Click `🚀 Project Actions` > `Run Full Setup (Install Triggers & Logging)`.
3.  **Authorize:**
    *   You must authorize the script. This is required to install triggers, send emails, and access/create the external log spreadsheet.
    *   Review the permissions and click `Allow`.

The setup is complete. Automations will run on edits, and the dashboard can be updated via the menu.

## Maintenance

-   To make changes, edit the files in `/src`.
-   To deploy changes, run `clasp push`.