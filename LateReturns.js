/**
 * ==========================================================
 * LATE RETURNS
 * ==========================================================
 * Populates the Late Returns section on the Dashboard.
 */

/**
 * Scans Daily Operations for overdue breaks and writes
 * them to the Dashboard's Late Returns section.
 */
function updateLateReturns() {

  var ss = SpreadsheetApp.getActive();

  var dash = getSheetOrThrow(ss, SHEETS.DASHBOARD);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Clear previous late returns
  dash.getRange("A45:H60").clearContent();

  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var output = [];

  for (var i = 0; i < data.length; i++) {

    var r = data[i];

    // Skip non-overdue agents
    if (r[COL.STATUS] !== STATUS.BREAK_OVERDUE) continue;

    output.push([
      r[COL.AGENT],          // Agent
      r[COL.QUEUE],          // Queue
      r[COL.SCHEDULED_BACK], // Expected Back
      "",                    // Actual Back (blank)
      "",                    // Variance (blank)
      r[COL.SUPERVISOR],     // Supervisor
      "",                    // Override
      "Overdue"              // Remarks
    ]);
  }

  if (output.length > 0) {
    dash.getRange(
      OPS_DATA_START_ROW + 39, // Row 45
      1,
      output.length,
      8
    ).setValues(output);
  }
}