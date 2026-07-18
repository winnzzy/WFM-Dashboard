/**
 * ==========================================================
 * DASHBOARD
 * ==========================================================
 * Builds the main Dashboard sheet with KPI cards,
 * live break monitor, queue share, late returns,
 * and supervisor override sections.
 */

/**
 * Builds the complete Dashboard sheet layout.
 * All section ranges and colors use shared constants.
 */
function buildDashboardSheet() {

  var ss = SpreadsheetApp.getActive();

  var dash = ss.getSheetByName(SHEETS.DASHBOARD);

  if (!dash) {
    dash = ss.insertSheet(SHEETS.DASHBOARD);
  }

  dash.clear();

  // --- KPI Cards Row 1 ---
  createKPICard(
    dash, "A1", "B3",
    "Total Active Agents Today",
    "0",
    COLORS.HEADER_BG,
    COLORS.HEADER_TEXT
  );

  createKPICard(
    dash, "D1", "E3",
    "Currently On Queue",
    "0",
    "#1B5E20",
    COLORS.HEADER_TEXT
  );

  createKPICard(
    dash, "G1", "H3",
    "Currently On Break",
    "0",
    "#E65100",
    COLORS.HEADER_TEXT
  );

  createKPICard(
    dash, "J1", "K3",
    "Break Overdue",
    "0",
    COLORS.LATE_HEADER_BG,
    COLORS.HEADER_TEXT
  );

  // --- KPI Cards Row 2 ---
  createKPICard(
    dash, "A5", "B7",
    "Breaks Started On Time",
    "0",
    COLORS.SECTION_HEADER_BG,
    "#000000"
  );

  createKPICard(
    dash, "D5", "E7",
    "Breaks Started Late",
    "0",
    COLORS.OVERRIDE_HEADER_SUB,
    "#000000"
  );

  createKPICard(
    dash, "G5", "H7",
    "Avg Break Duration (min)",
    "0",
    COLORS.SECTION_HEADER_BG,
    "#000000"
  );

  createKPICard(
    dash, "J5", "K7",
    "Supervisor Overrides Today",
    "0",
    COLORS.OVERRIDE_HEADER_BG,
    COLORS.HEADER_TEXT
  );

  // --- KPI Cards Row 3 ---
  createKPICard(
    dash, "A9", "B11",
    "Agents Absent Today",
    "0",
    COLORS.LATE_HEADER_SUB,
    "#000000"
  );

  createKPICard(
    dash, "D9", "E11",
    "Agents Logged Out",
    "0",
    COLORS.LATE_HEADER_SUB,
    "#000000"
  );

  createKPICard(
    dash, "G9", "H11",
    "Break Compliance %",
    "0%",
    COLORS.SECTION_HEADER_BG,
    "#000000"
  );

  createKPICard(
    dash, "J9", "K11",
    "Shifts Covered",
    "0",
    COLORS.SECTION_HEADER_BG,
    "#000000"
  );

  // --- Break Slots Section ---
  dash.getRange("A13:H13")
    .merge()
    .setValue("TODAY'S BREAK SLOTS")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var breakHeaders = [
    "Agent",
    "Queue",
    "Break Slot",
    "Scheduled Out",
    "Scheduled Back",
    "Actual Login",
    "Actual Out",
    "Actual Back"
  ];

  dash.getRange("A14:H14").setValues([breakHeaders]);

  dash.getRange("A14:H14")
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  // --- Live Break Monitor Section ---
  dash.getRange("A22:H22")
    .merge()
    .setValue("LIVE BREAK MONITOR")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var monitorHeaders = [
    "Agent",
    "Queue",
    "Break Slot",
    "Actual Out",
    "Expected Back",
    "Minutes Left",
    "Status",
    "Variance"
  ];

  dash.getRange("A23:H23").setValues([monitorHeaders]);

  dash.getRange("A23:H23")
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  // --- Queue Share Section ---
  dash.getRange("J13:N13")
    .merge()
    .setValue("QUEUE SHARE")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var queueHeaders = [
    "Queue",
    "Total",
    "On Queue",
    "On Break",
    "Break Overdue"
  ];

  dash.getRange("J14:N14").setValues([queueHeaders]);

  dash.getRange("J14:N14")
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  // --- Break Overdue Section ---
  dash.getRange("J20:N20")
    .merge()
    .setValue("BREAK OVERDUE")
    .setBackground(COLORS.LATE_HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var overdueHeaders = [
    "Agent",
    "Queue",
    "Expected Back",
    "Minutes Overdue",
    "Status"
  ];

  dash.getRange("J21:N21").setValues([overdueHeaders]);

  dash.getRange("J21:N21")
    .setBackground(COLORS.LATE_HEADER_SUB)
    .setFontWeight("bold");

  // --- Supervisor Overrides Section ---
  dash.getRange("J37:N37")
    .merge()
    .setValue("SUPERVISOR OVERRIDES")
    .setBackground(COLORS.OVERRIDE_HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var overrideHeaders = [
    "Agent",
    "Queue",
    "Override Time",
    "Reason",
    "Supervisor"
  ];

  dash.getRange("J38:N38").setValues([overrideHeaders]);

  dash.getRange("J38:N38")
    .setBackground(COLORS.OVERRIDE_HEADER_SUB)
    .setFontWeight("bold");

  // --- Agent Status Section ---
  dash.getRange("A36:H36")
    .merge()
    .setValue("AGENT STATUS")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  var statusHeaders = [
    "Agent",
    "Queue",
    "Shift",
    "Break Slot",
    "Status",
    "Break Used (min)",
    "Variance",
    "Remarks"
  ];

  dash.getRange("A37:H37").setValues([statusHeaders]);

  dash.getRange("A37:H37")
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  // --- System Log Section ---
  dash.getRange("A70:H70")
    .merge()
    .setValue("SYSTEM LOG")
    .setBackground("#D9D9D9")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  dash.getRange("A71").setValue(
    "Log entries will appear here after daily schedule generation."
  );

  // --- Column widths ---
  var widths = [
    250, 120, 100, 180, 100, 100, 140, 140, 140,
    250, 140, 100, 150, 180
  ];
  widths.forEach(function(w, i) {
    dash.setColumnWidth(i + 1, w);
  });
}


/**
 * Creates a KPI card on the dashboard.
 *
 * @param {SpreadsheetApp.Sheet} sheet — The Dashboard sheet
 * @param {string} titleCell — Top-left cell of the title area (e.g. "A1")
 * @param {string} valueCell — Top-left cell of the value area (e.g. "A2")
 * @param {string} title — KPI title text
 * @param {string} value — Initial value to display
 * @param {string} bgColor — Background hex color
 * @param {string} textColor — Text hex color
 */
function createKPICard(sheet, titleCell, valueCell, title, value, bgColor, textColor) {

  // Title cell
  sheet.getRange(titleCell)
    .setValue(title)
    .setBackground(bgColor)
    .setFontColor(textColor)
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center");

  // Value cell
  sheet.getRange(valueCell)
    .setValue(value)
    .setBackground(bgColor)
    .setFontColor(textColor)
    .setFontSize(28)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}