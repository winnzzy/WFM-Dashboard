/**
 * ==========================================================
 * SETUP
 * ==========================================================
 * Builds the "Daily Operations" sheet with headers,
 * column widths, and frozen rows.
 */

/**
 * Builds the Daily Operations sheet structure.
 * Uses shared SHEETS and COLORS constants from Utilities.js.
 */
function buildDailyOperations() {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(SHEETS.DAILY_OPS);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.DAILY_OPS);
  }

  sh.clear();

  // --- Title ---
  sh.getRange("A1:R1")
    .merge()
    .setValue("WFM DASHBOARD — OPERATIONS")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setFontSize(14)
    .setHorizontalAlignment("center");

  // --- Date ---
  sh.getRange("A2:R2")
    .merge()
    .setValue("Date: " + Utilities.formatDate(
      new Date(),
      "America/Los_Angeles",
      "MM-dd-yyyy"
    ))
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold")
    .setFontSize(12)
    .setHorizontalAlignment("center");

  // --- Legend ---
  sh.getRange("A3:R3")
    .merge()
    .setValue(
      "LEGEND:    " +
      "🟢 On Queue    |    " +
      "🟡 On Break    |    " +
      "🔴 Break Overdue    |    " +
      "⚪ Other"
    )
    .setBackground("#D9D9D9")
    .setFontSize(10)
    .setHorizontalAlignment("center");

  // --- Section headers ---
  var sections = [
    { col: 1, end: 1, label: "AGENT INFO" },
    { col: 2, end: 2, label: "" },
    { col: 3, end: 3, label: "" },
    { col: 4, end: 4, label: "" },
    { col: 5, end: 5, label: "" },
    { col: 6, end: 6, label: "" },
    { col: 7, end: 9, label: "SCHEDULE" },
    { col: 10, end: 14, label: "ACTUALS" },
    { col: 15, end: 15, label: "" },
    { col: 16, end: 16, label: "" },
    { col: 17, end: 18, label: "OVERRIDE" }
  ];

  sections.forEach(function(sec) {
    sh.getRange(4, sec.col, 1, sec.end - sec.col + 1)
      .merge()
      .setValue(sec.label)
      .setBackground("#D9D9D9")
      .setFontWeight("bold")
      .setFontSize(10)
      .setHorizontalAlignment("center");
  });

  // --- Column headers ---
  var headers = [
    "Date",
    "Agent",
    "Center",
    "Supervisor",
    "Shift",
    "Queue",
    "Break Slot",
    "Scheduled Out",
    "Scheduled Back",
    "Actual Login",
    "Actual Out",
    "Actual Back",
    "Break Used (min)",
    "Variance",
    "Status",
    "Remarks",
    "Override",
    "Override Time"
  ];

  sh.getRange(5, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setFontSize(10)
    .setHorizontalAlignment("center");

  // --- Column widths ---
  var widths = [
    120, 250, 120, 180, 100, 100,
    100, 140, 140, 140, 140, 140,
    150, 180, 130, 150, 130, 150
  ];

  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  // --- Freeze rows ---
  sh.setFrozenRows(5);
}