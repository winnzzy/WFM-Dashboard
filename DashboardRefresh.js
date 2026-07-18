/**
 * ==========================================================
 * DASHBOARD REFRESH
 * ==========================================================
 * Updates all live dashboard sections with current data
 * from the Daily Operations sheet.
 *
 * Called by:
 *   - updateAgentStatuses() after status recalculation
 *   - Menu: Refresh Dashboard
 *   - resetToday() after clearing data
 *   - startBreak() and returnFromBreak() after mutations
 */

/**
 * Refreshes all dashboard sections in a single pass.
 * Reads Daily Operations data once, then distributes
 * to all dashboard sub-sections.
 */
function refreshDashboard() {

  var ss = SpreadsheetApp.getActive();

  var dash = getSheetOrThrow(ss, SHEETS.DASHBOARD);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // --- Single batch read of all agent data ---
  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  // --- Counters ---
  var totalAgents = 0;
  var onQueue = 0;
  var onBreak = 0;
  var breakOverdue = 0;
  var absent = 0;
  var loggedOut = 0;
  var overrideCount = 0;

  var now = new Date();

  // --- Scan data ---
  for (var i = 0; i < data.length; i++) {

    var row = data[i];

    if (!row[COL.AGENT]) continue;

    totalAgents++;

    var status = row[COL.STATUS];

    if (status === STATUS.ON_QUEUE) {
      onQueue++;
    } else if (status === STATUS.ON_BREAK) {
      onBreak++;
    } else if (status === STATUS.BREAK_OVERDUE) {
      breakOverdue++;
    } else if (status === STATUS.ABSENT) {
      absent++;
    } else if (status === STATUS.LOGGED_OUT) {
      loggedOut++;
    }

    if (row[COL.OVERRIDE]) {
      overrideCount++;
    }
  }

  // --- Update KPI cards ---
  dash.getRange("B3").setValue(totalAgents);
  dash.getRange("E3").setValue(onQueue);
  dash.getRange("H3").setValue(onBreak);
  dash.getRange("K3").setValue(breakOverdue);

  // --- Break compliance ---
  var totalWithBreaks = totalAgents - absent - loggedOut;
  var compliance = totalWithBreaks > 0
    ? Math.round((onQueue / totalWithBreaks) * 100) + "%"
    : "0%";
  dash.getRange("H11").setValue(compliance);

  // --- Break Overdue section ---
  dash.getRange("J22:N60").clearContent();

  var overdueOutput = [];

  for (var j = 0; j < data.length; j++) {

    var r = data[j];

    if (!r[COL.AGENT]) continue;
    if (r[COL.STATUS] !== STATUS.BREAK_OVERDUE) continue;

    var minutesOverdue = "";

    if (isValidDate(r[COL.SCHEDULED_BACK])) {
      minutesOverdue = Math.round(
        (now.getTime() - r[COL.SCHEDULED_BACK].getTime()) / 60000
      );
    }

    overdueOutput.push([
      r[COL.AGENT],
      r[COL.QUEUE],
      r[COL.SCHEDULED_BACK],
      minutesOverdue,
      formatStatus(r[COL.STATUS])
    ]);
  }

  if (overdueOutput.length > 0) {
    dash.getRange(22, 10, overdueOutput.length, 5).setValues(overdueOutput);
    dash.getRange(22, 12, overdueOutput.length, 1)
      .setNumberFormat("h:mm AM/PM");
  }

  // --- Agent Status section ---
  dash.getRange("A38:H100").clearContent();

  var statusOutput = [];

  for (var k = 0; k < data.length; k++) {

    var row2 = data[k];

    if (!row2[COL.AGENT]) continue;

    statusOutput.push([
      row2[COL.AGENT],
      row2[COL.QUEUE],
      row2[COL.SHIFT],
      row2[COL.BREAK_SLOT],
      formatStatus(row2[COL.STATUS]),
      row2[COL.BREAK_USED],
      row2[COL.VARIANCE],
      row2[COL.REMARKS]
    ]);
  }

  if (statusOutput.length > 0) {
    dash.getRange(38, 1, statusOutput.length, 8).setValues(statusOutput);
  }

  // --- Also update Queue Share and Late Returns ---
  updateQueueShare();
  updateLateReturns();
}