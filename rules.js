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
  // Morning widened 2026-11-10 for the Morning Call second
  // break: first breaks run 11:00 AM–1:00 PM and second
  // breaks 1:30–3:00 PM, so a scheduled second break must
  // not demand a supervisor override to start. Standard
  // Morning breaks (12:00 PM–2:00 PM) are inside it too.
  Morning:    { start: 11, end: 15 },
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