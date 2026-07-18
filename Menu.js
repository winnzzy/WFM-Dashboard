/**
 * ==========================================================
 * TEAM LEAD MENU
 * ==========================================================
 */

function onOpen() {

  SpreadsheetApp.getUi()

    .createMenu("👨‍💼 Team Lead")

    .addItem("🔄 Refresh Dashboard","refreshDashboard")

    .addSeparator()

    .addItem("☕ Start Break","startBreak")

    .addItem("✅ Return From Break","returnFromBreak")

    .addSeparator()

    .addItem("📤 Generate Daily Report","generateDailyReport")

    .addItem("📋 Export Queue Share","exportQueueShare")

    .addSeparator()
    .addItem("🗓 Generate Daily Schedule","generateDailySchedule")
    .addItem("♻ Reset Today","resetToday")

    .addToUi();

}