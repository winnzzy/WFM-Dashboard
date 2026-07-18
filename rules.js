/**
 * ==========================================================
 * BUSINESS RULES
 * ==========================================================
 * Break window definitions per shift.
 * Used by Automation.js to determine if an agent is within
 * the approved break window.
 */

/** Break windows by shift (24-hour format, start inclusive, end exclusive). */
var BREAK_WINDOWS = {
  Morning:    { start: 12, end: 14 },
  Afternoon:  { start: 15, end: 17 },
  Night:      { start: 3,  end: 6  }
};

/**
 * Checks if the current time is within the approved break
 * window for the given shift.
 * @param {string} shift — "Morning", "Afternoon", or "Night"
 * @returns {boolean}
 */
function isWithinBreakWindow(shift) {

  var rule = BREAK_WINDOWS[shift];

  if (!rule) return false;

  var now = new Date();
  var hour = now.getHours();

  return hour >= rule.start && hour < rule.end;
}