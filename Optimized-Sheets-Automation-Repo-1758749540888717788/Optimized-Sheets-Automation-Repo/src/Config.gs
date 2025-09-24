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
    PROJECT_NAME: 1,  // A
    DETAILS: 3,       // C (Used for Inventory transfer)
    EQUIPMENT: 5,     // E
    PROGRESS: 6,      // F (Synced with Upcoming E, Triggers Framing)
    PERMITS: 7,       // G (Triggers Upcoming transfer)
    ARCHITECT: 8,     // H (Used for Framing transfer)
    DEADLINE: 9,      // I
    DELIVERED: 11,    // K (Triggers Inventory transfer)
    LOCATION: 15,     // O (Used for Upcoming transfer)
  },

  // --- Upcoming Columns (1-indexed) ---
  UPCOMING_COLS: {
    PROJECT_NAME: 1,       // A
    CONSTRUCTION_START: 2, // B
    INSTALL_DATE: 3,       // C
    DEADLINE: 4,           // D
    PROGRESS: 5,           // E
    EQUIPMENT: 6,          // F
    PERMITS: 7,            // G
    CONSTRUCTION: 8,       // H
    LOCATION: 9,           // I
    NOTES: 10,             // J
  },

  // --- Framing Columns (1-indexed) ---
  FRAMING_COLS: {
    PROJECT_NAME: 1, // A
    DEADLINE: 4,     // D
    ARCHITECT: 6,    // F
    EQUIPMENT: 8,    // H
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
    MONTH_COL: 1,       // A
    TOTAL_COL: 2,       // B
    UPCOMING_COL: 3,    // C
    OVERDUE_COL: 4,     // D
    APPROVED_COL: 5,    // E
    GT_UPCOMING_COL: 6, // F
    GT_OVERDUE_COL: 7,  // G
    GT_TOTAL_COL: 8,    // H
    GT_APPROVED_COL: 9, // I
    MISSING_DEADLINE_CELL: "L1",
    FIXED_ROW_COUNT: 150,
    HIDE_COL_START: 11, // K
    HIDE_COL_END: 26,   // Z
    CHART_START_ROW: 2,
    CHART_ANCHOR_COL: 10, // J
    TEMP_DATA_START_COL: 12 // L
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
    SERIES_PAST: ['Month', 'Overdue', 'Total'],
    SERIES_UPCOMING_NO_OVERDUE: ['Month', 'Upcoming', 'Total']
  }
};