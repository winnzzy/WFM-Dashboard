/**
 * ==========================================================
 * SETTINGS SHEET
 * ==========================================================
 * Creates and configures the Settings sheet with dropdown
 * source data for Centers, Shifts, Supervisors, and Statuses.
 */

/**
 * Builds the Settings sheet with all dropdown source values.
 * This is the single source of truth for validation dropdowns.
 */
function buildSettings() {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(SHEETS.SETTINGS);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.SETTINGS);
  }

  sh.clear();

  // --- Headers ---
  var headers = [
    "Queues",
    "Shifts",
    "Status",
    "Centers",
    "Supervisors"
  ];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // --- Queues (Column A) ---
  sh.getRange("A2:A5").setValues([
    ["Call"],
    ["Email"],
    ["Clara"],
    ["Ebanqo"]
  ]);

  // --- Shifts (Column B) ---
  sh.getRange("B2:B5").setValues([
    ["Morning"],
    ["Afternoon"],
    ["Night"],
    ["OFF"]
  ]);

  // --- Status (Column C) ---
  sh.getRange("C2:C11").setValues([
    ["On Queue"],
    ["On Break"],
    ["Break Overdue"],
    ["Absent"],
    ["Logged Out"],
    ["Meeting"],
    ["Training"],
    ["Coaching"],
    ["System Issue"],
    ["Leave"]
  ]);

  // --- Centers (Column D) ---
  sh.getRange("D2:D6").setValues([
    ["Lagos"],
    ["Abuja"],
    ["PH"],
    ["Kano"],
    ["Ibadan"]
  ]);

  // --- Supervisors (Column E) ---
  sh.getRange("E2:E7").setValues([
    ["Supervisor A"],
    ["Supervisor B"],
    ["Supervisor C"],
    ["Supervisor D"],
    ["Supervisor E"],
    ["Supervisor F"]
  ]);

  // --- Formatting ---
  var widths = [150, 150, 150, 150, 150];

  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.setFrozenRows(1);
}