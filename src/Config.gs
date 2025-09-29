/**
 * @OnlyCurrentDoc
 * Config.gs
 * Centralized configuration for the entire project (Automations and Dashboard).
 * Standardized on 1-based indexing for clarity (matching the spreadsheet view).
 * Code MUST subtract 1 when accessing array indices (0-based).
 */

const CONFIG = {
  APP_NAME: "Sheet Automations & Dashboard",
  
  // --- General Settings ---
  ERROR_NOTIFICATION_EMAIL: "rpenn@mobility123.com",

  // --- Sheet Names ---
  SHEETS: {
    FORECASTING: "Forecasting",
    UPCOMING: "Upcoming",
    INVENTORY: "Inventory_Elevators",
    FRAMING: "Framing",
    DASHBOARD: "Dashboard",
    OVERDUE_DETAILS: "Overdue Details"
  },

  // --- Status Strings (used in both Automations and Dashboard) ---
  STATUS_STRINGS: {
    IN_PROGRESS: "In Progress",
    PERMIT_APPROVED: "approved",
    SCHEDULED: "Scheduled"
  },

  // --- Forecasting Columns (1-indexed) ---
  FORECASTING_COLS: {
    SFID: 1,          // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2,  // B
    DETAILS: 4,       // D (Used for Inventory transfer)
    EQUIPMENT: 6,     // F
    PROGRESS: 7,      // G (Synced with Upcoming F, Triggers Framing)
    PERMITS: 8,       // H (Triggers Upcoming transfer)
    ARCHITECT: 9,     // I (Used for Framing transfer)
    DEADLINE: 10,     // J
    DELIVERED: 12,    // L (Triggers Inventory transfer)
    LOCATION: 16,     // P (Used for Upcoming transfer)
  },

  // --- Upcoming Columns (1-indexed) ---
  UPCOMING_COLS: {
    SFID: 1,               // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2,       // B
    CONSTRUCTION_START: 3, // C
    DEADLINE: 4,           // D
    PROGRESS: 5,           // E
    EQUIPMENT: 6,          // F
    PERMITS: 7,            // G
    CONSTRUCTION: 8,       // H
    // Column I ("Tr") is intentionally unmapped as it's not part of any automation.
    LOCATION: 10,          // J
    NOTES: 11,             // K
  },

  // --- Framing Columns (1-indexed) ---
  FRAMING_COLS: {
    SFID: 1,        // A - Salesforce ID for unique record syncing
    PROJECT_NAME: 2,// B
    DEADLINE: 5,    // E
    ARCHITECT: 7,   // G
    EQUIPMENT: 9,   // I
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
    SPREADSHEET_ID_PROP: "LOG_SPREADSHEET_ID" // Script property key
  },

  // ================= DASHBOARD SPECIFIC CONFIG =================

  // --- Date Range for Dashboard ---
  DASHBOARD_DATES: {
    START: new Date(2024, 0, 1), // Jan 1, 2024 (Month is 0-indexed)
    END: new Date(2027, 11, 1)   // Dec 1, 2027
  },

  // Columns from Forecasting (by key) to display on the Overdue Details sheet
  OVERDUE_DETAILS_DISPLAY_KEYS: ['PROJECT_NAME', 'DEADLINE', 'PROGRESS', 'ARCHITECT', 'EQUIPMENT'],

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
    MONTH_FORMAT: "mmmm yyyy",
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