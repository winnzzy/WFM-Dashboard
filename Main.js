/**
 * ==========================================================
 * MAIN — INSTALL
 * ==========================================================
 * Master installer for the Team Lead Dashboard.
 * Run this function once to create all sheets, apply
 * formatting, validation, formulas, and named ranges.
 *
 * Execution order matters — Settings and Roster must be
 * built before Daily Operations, and Daily Operations
 * before Dashboard.
 */

/**
 * Installs the complete Team Lead Dashboard.
 * Run from the Apps Script editor or via the custom menu.
 */
function INSTALL() {

  buildSettings();

  buildRoster();

  buildDailyOperations();

  buildWeeklySchedule();

  buildDashboardSheet();

  applyDropdowns();

  applyAgentDropdown();

  applyRosterValidation();

  // Notification Log sheet (created on first send if missing).
  if (typeof ensureNotificationLogSheet === "function") {
    ensureNotificationLogSheet();
  }

  applyFormatting();

  applyFormulas();

  createNamedRanges();

  updateBreakMonitor();

  updateQueueShare();

  updateLateReturns();

  refreshDashboard();

  onOpen();
}
