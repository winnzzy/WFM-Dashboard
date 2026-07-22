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

    .addItem("📅 Generate Today's Schedule", "generateTodayFromWeekly")
    .addItem("📆 Generate Tomorrow's Schedule", "generateTomorrowFromWeekly")
    .addItem("🗑 Clear Today's Schedule", "clearTodaySchedule")

    .addSeparator()

    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("🔧 System")
        .addItem("🩺 Health Check", "systemHealthCheck")
        .addItem("🔐 Verify Permissions", "verifyPermissions")
        .addItem("🔧 Rebuild Dashboard", "refreshDashboard")
        .addItem("⚡ Reinstall Triggers", "reinstallTriggers")
        .addItem("📊 System Report", "showSystemReport")
    )

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
 * Clears today's generated schedule from Daily Schedule
 * and Daily Operations. Does NOT touch Weekly Schedule,
 * Agent Roster, Settings, Dashboard layout, Named Ranges,
 * dropdowns, or validation.
 *
 * Batch operations for performance:
 *   - Single clearContent() on Daily Schedule data area
 *   - Single getValues() + targeted column clear + single
 *     setValues() on Daily Operations data area
 */
function clearTodaySchedule() {

  var ui = SpreadsheetApp.getUi();

  var response = ui.alert(
    "Clear Today's Schedule",
    "Are you sure you want to clear today's generated schedule?\n\n" +
    "This will erase today's generated Daily Schedule and " +
    "Daily Operations only.\n\n" +
    "Weekly Schedule and Agent Roster will NOT be affected.",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActive();

  // =========================================================
  // 1. CLEAR DAILY SCHEDULE — batch clearContent
  // =========================================================
  var scheduleSheet = getSheetOrThrow(ss, SHEETS.DAILY_SCHEDULE);
  scheduleSheet.getRange("A2:J500").clearContent();

  // =========================================================
  // 2. CLEAR DAILY OPERATIONS — batch read + targeted clear
  // =========================================================
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  var dataRange = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  );

  var data = dataRange.getValues();

  // Columns to clear (0-based indexes):
  //   5  = Queue           (F)
  //   6  = Break Slot      (G)
  //   7  = Scheduled Out   (H)
  //   8  = Scheduled Back  (I)
  //   9  = Login           (J)
  //   10 = Actual Out      (K)
  //   11 = Actual Back     (L)
  //   12 = Break Used      (M)
  //   13 = Variance        (N)
  //   14 = Status          (O)
  //   15 = Remarks         (P)
  //   16 = Override        (Q)
  //   17 = Override Time   (R)
  var colsToClear = [5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17];

  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < colsToClear.length; j++) {
      data[i][colsToClear[j]] = "";
    }
  }

  dataRange.setValues(data);

  SpreadsheetApp.flush();

  // =========================================================
  // 3. REFRESH DASHBOARD
  // =========================================================
  refreshDashboard();

  ui.alert(
    "Schedule Cleared",
    "Today's Daily Schedule and Daily Operations have been " +
    "cleared successfully.\n\n" +
    "Weekly Schedule and Agent Roster remain untouched.",
    ui.ButtonSet.OK
  );
}
