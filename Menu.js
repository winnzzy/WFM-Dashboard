/**
 * ==========================================================
 * TEAM LEAD MENU
 * ==========================================================
 * Custom menu for the Team Lead Dashboard.
 */

/**
 * Creates the custom "Team Lead" menu on spreadsheet open.
 */
function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("👨‍💼 Team Lead")

    .addItem("🔄 Refresh Dashboard", "refreshDashboard")

    .addSeparator()

    .addItem("☕ Start Break", "startBreak")
    .addItem("✅ Return From Break", "returnFromBreak")

    .addSeparator()

    .addItem("📤 Generate Daily Report", "generateDailyReport")
    .addItem("📋 Export Queue Share", "exportQueueShare")

    .addSeparator()

    .addItem("🗓 Generate Daily Schedule", "generateDailySchedule")
    .addItem("♻ Reset Today", "resetToday")

    .addToUi();
}


/**
 * Generates a summary report of today's operations.
 * Placeholder — to be implemented with specific reporting logic.
 */
function generateDailyReport() {
  SpreadsheetApp.getUi().alert(
    "Daily Report",
    "This feature is not yet implemented.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Exports the current Queue Share data to a new sheet or file.
 * Placeholder — to be implemented with specific export logic.
 */
function exportQueueShare() {
  SpreadsheetApp.getUi().alert(
    "Export Queue Share",
    "This feature is not yet implemented.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * Resets today's Daily Operations data.
 * Clears agent assignments, break data, and statuses
 * while preserving the sheet structure.
 */
function resetToday() {

  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    "Reset Today",
    "This will clear all agent data for today. Continue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActive();
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Clear data rows only (preserve headers in rows 1–5)
  ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).clearContent();

  // Clear override columns
  ops.getRange(
    OPS_DATA_START_ROW,
    COL.OVERRIDE + 1,
    MAX_OPS_ROWS,
    2
  ).clearContent();

  refreshDashboard();
}