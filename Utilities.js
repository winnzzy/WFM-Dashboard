/**
 * ==========================================================
 * WFM DASHBOARD
 * Utilities.gs
 * ==========================================================
 */

/**
 * Daily Operations column indexes (0-based)
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
  STATUS: 14            // O

};


/**
 * Queue Names
 */
const QUEUES = [
  "Call",
  "Email",
  "Clara",
  "Ebanqo"
];


/**
 * Agent Statuses
 */
const STATUS = {

  ON_QUEUE: "On Queue",

  ON_BREAK: "On Break",

  BREAK_OVERDUE: "Break Overdue"

};


/**
 * Dashboard ranges
 */
const DASHBOARD = {

  BREAK_MONITOR: "A24:H200",

  QUEUE_SHARE: "J23:N30",

  LATE_RETURNS: "A45:H70"

};


/**
 * Returns minutes since midnight.
 */
function toMinutes(value) {

  if (!(value instanceof Date)) return null;

  return value.getHours() * 60 + value.getMinutes();

}


/**
 * Checks if a value is a valid Date.
 */
function isValidDate(value) {

  return value instanceof Date && !isNaN(value);

}


/**
 * Formats dashboard status.
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