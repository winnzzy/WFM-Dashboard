/**
 * ==========================================================
 * QUEUE SHARE
 * ==========================================================
 * Populates the Queue Share KPI table on the Dashboard
 * with live agent counts per queue.
 */

/**
 * Counts agents assigned to each queue and writes
 * the results to the Dashboard queue share section.
 */
function updateQueueShare() {

  var ss = SpreadsheetApp.getActive();

  var dash = getSheetOrThrow(ss, SHEETS.DASHBOARD);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Clear previous queue share data
  dash.getRange("J23:N30").clearContent();

  // Header
  dash.getRange("J23:N23").setValues([[
    "Queue",
    "Total",
    "On Queue",
    "On Break",
    "Break Overdue"
  ]]);

  dash.getRange("J23:N23")
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold");

  // Read agent data
  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  // Count agents per queue per status
  var counts = {};

  QUEUES.forEach(function(queue) {
    counts[queue] = {
      total: 0,
      onQueue: 0,
      onBreak: 0,
      overdue: 0
    };
  });

  for (var i = 0; i < data.length; i++) {

    var row = data[i];

    if (!row[COL.AGENT]) continue;

    var queue = row[COL.QUEUE];
    var status = row[COL.STATUS];

    if (!counts[queue]) {
      counts[queue] = {
        total: 0,
        onQueue: 0,
        onBreak: 0,
        overdue: 0
      };
    }

    counts[queue].total++;

    if (status === STATUS.ON_QUEUE) {
      counts[queue].onQueue++;
    } else if (status === STATUS.ON_BREAK) {
      counts[queue].onBreak++;
    } else if (status === STATUS.BREAK_OVERDUE) {
      counts[queue].overdue++;
    }
  }

  // Write to dashboard
  var output = [];

  QUEUES.forEach(function(queue) {
    var c = counts[queue];
    output.push([
      queue,
      c.total,
      c.onQueue,
      c.onBreak,
      c.overdue
    ]);
  });

  if (output.length > 0) {
    dash.getRange(24, 10, output.length, 5).setValues(output);
  }
}