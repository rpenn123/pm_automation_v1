/**
 * @OnlyCurrentDoc
 * Config.gs
 * Centralized configuration for the entire project (Automations and Dashboard).
 * Standardized on 1-based indexing for clarity (matching the spreadsheet view).
 * Code MUST subtract 1 when accessing array indices (0-based).
 */

/**
 * Global configuration object for the entire application.
 * This object centralizes all settings, including sheet names, column mappings,
 * status strings, and dashboard layout, to make maintenance and updates easier.
 * All column numbers are 1-indexed to match the user-facing spreadsheet view.
 *
 * @const {object}
 */
const CONFIG = {
  APP_NAME: "Sheet Automations & Dashboard",
  
  // --- General Settings ---
  // --- Sheet Names ---
  SHEETS: {
    FORECASTING: "Forecasting",
    UPCOMING: "Upcoming",
    INVENTORY: "Inventory_Elevators",
    FRAMING: "Framing",
    DASHBOARD: "Dashboard"
  },

  // --- Status Strings (used in both Automations and Dashboard) ---
  STATUS_STRINGS: {
    IN_PROGRESS: "In Progress",
    PERMIT_APPROVED: "approved",
    SCHEDULED: "Scheduled",
    COMPLETED: "Completed",
    CANCELLED: "Cancelled"
  },

  // --- Forecasting Columns (1-indexed) ---
  FORECASTING_COLS: {
    SFID: 1, // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2, // B - The official name of the project
    CREATED: 3, // C - The date the project record was created
    DETAILS: 4, // D - Detailed description, used for Inventory transfer
    PRIORITY: 5, // E - Priority level of the project
    EQUIPMENT: 6, // F - Specific equipment being used
    PROGRESS: 7, // G - Current status, synced with Upcoming sheet, triggers Framing automation
    PERMITS: 8, // H - Permit status, triggers transfer to Upcoming sheet
    ARCHITECT: 9, // I - Architect contact, used for Framing transfer
    DEADLINE: 10, // J - The project's final deadline
    SHIPPING: 11, // K - Shipping status for materials/equipment
    DELIVERED: 12, // L - Delivery confirmation, triggers Inventory transfer
    PERMITS_APPROVED: 13, // M - Date when permits were officially approved
    START_DATE: 14, // N - The scheduled start date for construction
    SITE_PREP: 15, // O - Status of the site preparation
    LOCATION: 16, // P - Physical address, used for Upcoming transfer
    OWNER: 17, // Q - The assigned project owner or manager
    LAST_EDIT_AT_HIDDEN: 18, // R - Timestamp of the last edit (hidden from users)
    LAST_EDIT: 20 // T - A user-friendly representation of the last edit time
  },

  // --- Upcoming Columns (1-indexed) ---
  UPCOMING_COLS: {
    SFID: 1, // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2, // B - The official name of the project
    CONSTRUCTION_START: 3, // C - The scheduled start date for construction
    DEADLINE: 4, // D - The project's final deadline
    PROGRESS: 5, // E - Current status of the project
    EQUIPMENT: 6, // F - Specific equipment being used
    PERMITS: 7, // G - Current status of the permits
    CONSTRUCTION: 8, // H - Current status of the construction phase
    // Column I ("Tr") is intentionally unmapped as it's not part of any automation.
    LOCATION: 10, // J - Physical address of the project site
    NOTES: 11 // K - Any relevant notes about the project
  },

  // --- Framing Columns (1-indexed) ---
  FRAMING_COLS: {
    SFID: 1, // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2, // B - The official name of the project
    DEADLINE: 5, // E - The project's final deadline
    ARCHITECT: 7, // G - The architect assigned to the project
    EQUIPMENT: 9 // I - Specific equipment being used
  },

  // --- Inventory Columns (1-indexed) ---
  INVENTORY_COLS: {
    PROJECT_NAME: 2, // B
    PROGRESS: 5,     // E
    EQUIPMENT: 7,    // G
    DETAILS: 13,     // M
  },

  // --- Last Edit Configuration ---
  LAST_EDIT: {
    AT_HEADER: "Last Edit At (hidden)",
    REL_HEADER: "Last Edit",
    // Sheets that should have "Last Edit" tracking
    TRACKED_SHEETS: ["Forecasting", "Upcoming", "Framing"]
  },

  // --- External Logging Configuration ---
  LOGGING: {
    SPREADSHEET_NAME: "Sheet Automations Logs",
    SPREADSHEET_ID_PROP: "LOG_SPREADSHEET_ID", // Script property key
    ERROR_EMAIL_PROP: "ERROR_NOTIFICATION_EMAIL" // Script property key
  },

  // ================= DASHBOARD SPECIFIC CONFIG =================

  // --- Date Range for Dashboard ---
  DASHBOARD_DATES: {
    START: new Date(2024, 0, 1), // Jan 1, 2024 (Month is 0-indexed)
    END: new Date(2027, 11, 1)   // Dec 1, 2027
  },

  // --- Dashboard Layout (1-indexed) ---
  DASHBOARD_LAYOUT: {
    YEAR_COL: 1,        // A
    MONTH_COL: 2,       // B
    TOTAL_COL: 3,       // C
    UPCOMING_COL: 4,    // D
    OVERDUE_COL: 5,     // E
    APPROVED_COL: 6,    // F
    GT_UPCOMING_COL: 7, // G
    GT_OVERDUE_COL: 8,  // H
    GT_TOTAL_COL: 9,    // I
    GT_APPROVED_COL: 10,// J
    MISSING_DEADLINE_CELL: "M1",
    FIXED_ROW_COUNT: 150,
    HIDE_COL_START: 12, // L
    HIDE_COL_END: 27,   // AA
    CHART_START_ROW: 2,
    CHART_ANCHOR_COL: 11, // K
    TEMP_DATA_START_COL: 13 // M
  },

  // --- Formatting (SLEEK THEME) ---
  DASHBOARD_FORMATTING: {
    HEADER_BACKGROUND: "#4A6572",
    HEADER_FONT_COLOR: "#FFFFFF",
    TOTALS_BACKGROUND: "#F5F5F5",
    BANDING_COLOR_EVEN: "#FFFFFF",
    BANDING_COLOR_ODD: "#F8F9FA",
    BORDER_COLOR: "#CCCCCC",
    OVERDUE_ROW_HIGHLIGHT: "#FFEBEE",
    MONTH_FORMAT: "mmmm",
    COUNT_FORMAT: "0",
    CHART_COLORS: {
        overdue: '#D32F2F',
        upcoming: '#1976D2',
        total: '#616161'
    }
  },

  // --- Charting ---
  DASHBOARD_CHARTING: {
    ENABLED: true,
    CHART_HEIGHT: 260,
    CHART_WIDTH: 475,
    ROW_SPACING: 15,
    PAST_MONTHS_COUNT: 3,
    UPCOMING_MONTHS_COUNT: 6,
    STACKED: false,
    SERIES_PAST: ['Month', 'Overdue', 'Total'],
    SERIES_UPCOMING_NO_OVERDUE: ['Month', 'Upcoming', 'Total']
  }
};