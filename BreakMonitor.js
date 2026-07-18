/**
 * ==========================================================
 * LIVE BREAK MONITOR
 * ==========================================================
 * Populates the Live Break Monitor section on the Dashboard
 * with agents currently on break or overdue.
 */

/**
 * Scans Daily Operations for agents on break/overdue
 * and writes them to the Dashboard break monitor section.
 */
function updateBreakMonitor() {

  var ss = SpreadsheetApp.getActive();

  var dash = getSheetOrThrow(ss, SHEETS.DASHBOARD);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Clear previous monitor data
  dash.getRange("A22:H200").clearContent();

  // Section title
  dash.getRange("A22:H22")
    .merge()
    .setValue("LIVE BREAK MONITOR")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Column headers
  dash.getRange("A23:H23").setValues([[
    "Agent",
    "Queue",
    "Break Slot",
    "Actual Out",
    "Expected Back",
    "Minutes Left",
    "Status",
    "Variance"
  ]]);

  dash.getRange("A23:H23")
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  // Read agent data
  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var output = [];
  var now = new Date();

  for (var i = 0; i < data.length; i++) {

    var row = data[i];
    var status = row[COL.STATUS];

    if (status !== STATUS.ON_BREAK && status !== STATUS.BREAK_OVERDUE) {
      continue;
    }

    var minutesLeft = "";

    if (isValidDate(row[COL.SCHEDULED_BACK])) {
      minutesLeft = Math.round(
        (row[COL.SCHEDULED_BACK].getTime() - now.getTime()) / 60000
      );
    }

    output.push([
      row[COL.AGENT],
      row[COL.QUEUE],
      row[COL.BREAK_SLOT],
      row[COL.ACTUAL_OUT],
      row[COL.SCHEDULED_BACK],
      minutesLeft,
      formatStatus(status),
      row[COL.VARIANCE]
    ]);
  }

  if (output.length > 0) {

    dash.getRange(24, 1, output.length, 8).setValues(output);

    // Format time columns D & E (Actual Out, Expected Back)
    dash.getRange(24, 4, output.length, 2)
      .setNumberFormat("h:mm AM/PM");

  } else {

    dash.getRange("A24")
      .setValue("No agents are currently on break.");

  }
}