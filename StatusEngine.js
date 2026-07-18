/**
 * ==========================================================
 * STATUS ENGINE
 * ==========================================================
 * Runs on a 1-minute time-driven trigger.
 * Evaluates each agent's scheduled break times and sets
 * the appropriate status in Daily Operations column O.
 *
 * Business logic preserved exactly as original:
 *   - Before scheduled out  → On Queue
 *   - During break window   → On Break
 *   - After scheduled back  → Break Overdue
 */

/**
 * Recalculates agent statuses based on current time.
 * Also triggers a dashboard refresh after writing.
 */
function updateAgentStatuses() {

  var ss = SpreadsheetApp.getActive();

  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  var range = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  );

  var data = range.getValues();

  var nowMinutes = toMinutes(new Date());

  for (var i = 0; i < data.length; i++) {

    var row = data[i];

    if (!row[COL.AGENT]) continue;

    var scheduledOut = row[COL.SCHEDULED_OUT];
    var scheduledBack = row[COL.SCHEDULED_BACK];

    if (!isValidDate(scheduledOut) || !isValidDate(scheduledBack)) {
      row[COL.STATUS] = "";
      continue;
    }

    var out = toMinutes(scheduledOut);
    var back = toMinutes(scheduledBack);

    if (nowMinutes < out) {
      row[COL.STATUS] = STATUS.ON_QUEUE;
    } else if (nowMinutes < back) {
      row[COL.STATUS] = STATUS.ON_BREAK;
    } else {
      row[COL.STATUS] = STATUS.BREAK_OVERDUE;
    }
  }

  range.setValues(data);

  refreshDashboard();
}