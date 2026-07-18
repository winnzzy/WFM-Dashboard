/**
 * ==========================================================
 * WFM DASHBOARD
 * Utilities.gs — Shared constants and helper functions
 * ==========================================================
 */

/**
 * Daily Operations column indexes (0-based for array access).
 * Maps to columns A–R in the "Daily Operations" sheet.
 */
const COL = {
  DATE: 0,              // A
  AGENT: 1,             // B
  TEAM: 2,              // C
  SUPERVISOR: 3,        // D
  SHIFT: 4,             // E
  QUEUE: 5,             // F
  BREAK_SLOT: 6,        // G
  SCHEDULED_OUT: 7,     // H
  SCHEDULED_BACK: 8,    // I
  LOGIN: 9,             // J
  ACTUAL_OUT: 10,       // K
  ACTUAL_BACK: 11,      // L
  BREAK_USED: 12,       // M
  VARIANCE: 13,         // N
  STATUS: 14,           // O
  REMARKS: 15,          // P
  OVERRIDE: 16,         // Q
  OVERRIDE_TIME: 17     // R
};

/**
 * Total number of columns in the Daily Operations sheet.
 */
const OPS_COL_COUNT = 18;

/**
 * Queue names — single source of truth.
 */
const QUEUES = [
  "Call",
  "Email",
  "Clara",
  "Ebanqo"
];

/**
 * Agent status values — single source of truth.
 */
const STATUS = {
  ON_QUEUE: "On Queue",
  ON_BREAK: "On Break",
  BREAK_OVERDUE: "Break Overdue"
};

/**
 * Dashboard section ranges (on the "Dashboard" sheet).
 */
const DASHBOARD = {
  BREAK_MONITOR: "A24:H200",
  QUEUE_SHARE: "J23:N30",
  LATE_RETURNS: "A45:H70"
};

/**
 * Shared color constants.
 */
const COLORS = {
  HEADER_BG: "#0F4C81",
  HEADER_TEXT: "#FFFFFF",
  ALT_ROW: "#F8F9FA",
  BODY_BG: "#FFFFFF",
  SECTION_HEADER_BG: "#D9EAD3",
  LATE_HEADER_BG: "#C62828",
  LATE_HEADER_SUB: "#F4CCCC",
  OVERRIDE_HEADER_BG: "#F57C00",
  OVERRIDE_HEADER_SUB: "#FCE5CD"
};

/**
 * Sheet name constants — single source of truth for sheet references.
 */
const SHEETS = {
  DASHBOARD: "Dashboard",
  DAILY_OPS: "Daily Operations",
  AGENT_ROSTER: "Agent Roster",
  DAILY_SCHEDULE: "Daily Schedule",
  SETTINGS: "Settings"
};

/**
 * Maximum number of agent rows in Daily Operations.
 */
const MAX_OPS_ROWS = 500;

/**
 * Data starts at this row in Daily Operations (1-based).
 */
const OPS_DATA_START_ROW = 6;

/**
 * Data starts at this row in Agent Roster (1-based).
 */
const ROSTER_DATA_START_ROW = 2;

/**
 * Maximum roster rows to scan.
 */
const MAX_ROSTER_ROWS = 300;

/**
 * Converts a Date to minutes since midnight.
 * @param {Date} value
 * @returns {number|null}
 */
function toMinutes(value) {
  if (!(value instanceof Date)) return null;
  return value.getHours() * 60 + value.getMinutes();
}

/**
 * Checks if a value is a valid Date.
 * @param {*} value
 * @returns {boolean}
 */
function isValidDate(value) {
  return value instanceof Date && !isNaN(value);
}

/**
 * Returns a formatted status string with emoji for the dashboard.
 * @param {string} status
 * @returns {string}
 */
function formatStatus(status) {
  switch (status) {
    case STATUS.ON_QUEUE:
      return "🟢 On Queue";
    case STATUS.ON_BREAK:
      return "🟡 On Break";
    case STATUS.BREAK_OVERDUE:
      return "🔴 Break Overdue";
    default:
      return status;
  }
}

/**
 * Returns a sheet by name, or throws if not found.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {string} name — use SHEETS constant
 * @returns {SpreadsheetApp.Sheet}
 */
function getSheetOrThrow(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('Sheet "' + name + '" not found.');
  }
  return sheet;
}