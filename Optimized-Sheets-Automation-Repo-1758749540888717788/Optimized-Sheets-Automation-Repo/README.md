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

The project utilizes a modular architecture for improved maintainability:

-   `/src/Config.gs`: Centralized settings (Single Source of Truth).
-   `/src/Utilities.gs`: Shared helper functions (normalization, lookups, dates).
-   `/src/LoggerService.gs`: Encapsulated logging and error handling.
-   `/src/LastEditService.gs`: Manages "Last Edit" columns.
-   `/src/TransferEngine.gs`: Generic, reusable data transfer logic.
-   `/src/Automations.gs`: `onEdit` routing, synchronization, and transfer definitions.
-   `/src/Dashboard.gs`: Dashboard generation logic.
-   `/src/Setup.gs`: Menu UI and installation routines.

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