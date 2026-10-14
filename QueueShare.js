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
  var queueShareRows = Math.max(QUEUES.length + 2, 1);

  // Clear previous queue share data
  dash.getRange(15, 10, queueShareRows, 5).breakApart().clearContent();

  // Header
  dash.getRange("J14:N14").setValues([[
    "Queue",
    "Total",
    "On Queue",
    "On Break",
    "Break Overdue"
  ]]);

  dash.getRange("J14:N14")
    .setBackground(COLORS.SECTION_HEADER_BG)
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
    dash.getRange(15, 10, output.length, 5).setValues(output);
  }
}

/**
 * Exports the current Queue Share table to a new sheet.
 * Keeps all existing dashboard logic unchanged.
 */
function exportQueueShareToSheet() {

  var ss = SpreadsheetApp.getActive();
  var scheduleSheet = getSheetOrThrow(ss, SHEETS.DAILY_SCHEDULE);

  var lastRow = Math.max(scheduleSheet.getLastRow(), 1);
  var lastCol = Math.max(scheduleSheet.getLastColumn(), 1);
  var values = scheduleSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  var hasDataRows = false;
  for (var i = 2; i < values.length; i++) {
    if (String(values[i][1] || "").trim() !== "") {
      hasDataRows = true;
      break;
    }
  }

  if (!hasDataRows) {
    SpreadsheetApp.getUi().alert(
      "Export Queue Share",
      "No Daily Schedule data available to export.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  var exportSheet = ss.insertSheet(getUniqueQueueShareExportSheetName_(ss));

  exportSheet.getRange(1, 1).setValue("Daily Schedule Export").setFontWeight("bold");
  exportSheet.getRange(2, 1).setValue("Generated At").setFontWeight("bold");
  exportSheet.getRange(2, 2)
    .setValue(new Date())
    .setNumberFormat("dd-MMM-yyyy h:mm AM/PM");

  exportSheet.getRange(4, 1, values.length, values[0].length).setValues(values);
  exportSheet.getRange(4, 1, 1, values[0].length)
    .setBackground(COLORS.SECTION_HEADER_BG)
    .setFontWeight("bold");

  exportSheet.setFrozenRows(4);
  exportSheet.autoResizeColumns(1, values[0].length);

  SpreadsheetApp.getUi().alert(
    "Export Queue Share",
    "Daily Schedule exported to sheet: " + exportSheet.getName(),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Builds a unique export sheet name.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @returns {string}
 */
function getUniqueQueueShareExportSheetName_(ss) {
  var tz = ss.getSpreadsheetTimeZone() || "Africa/Lagos";
  var stamp = Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmm");
  var base = "Daily Schedule Export " + stamp;
  var name = base;
  var suffix = 2;

  while (ss.getSheetByName(name)) {
    name = base + " (" + suffix + ")";
    suffix++;
  }

  return name;
}
