/**
 * ==========================================================
 * DAILY REPORT
 * ==========================================================
 * Replaces the old placeholder ("This feature is not yet
 * implemented") with a real end-of-day summary built from
 * Daily Operations — the same sheet the scheduler writes.
 *
 * Team Lead -> Generate Daily Report writes a "Daily Report"
 * sheet and shows a summary alert. Nothing is emailed
 * automatically; the team lead decides what to forward.
 *
 * How each agent is classified (from Daily Operations):
 *   • Scheduled — a row exists for today.
 *   • Break taken — Actual Out is filled.
 *   • No break recorded — Actual Out is blank at report time.
 *   • Late return — Variance says "Late by X mins".
 *   • On time — Variance says "On Time".
 *   • Early return — Variance says "Early by X mins".
 *   • Overdue now — Status is Break Overdue, or Actual Out
 *     is filled, Actual Back is blank, and now is past
 *     Scheduled Back. (The status engine alone would flag
 *     everyone past their back time, returned or not, so
 *     the Actual Back check is what makes this real.)
 *
 * This report is deliberately the raw material for the
 * weekly business report: per-agent scheduled vs actual,
 * late returns and adherence, in one place.
 */

/**
 * Builds the Daily Report sheet from Daily Operations and
 * shows a summary alert. Returns the summary counts.
 */
function generateDailyReport() {

  var ss = SpreadsheetApp.getActive();
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  var data = ops.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var reportDate = new Date();
  var rows = [];
  var scheduled = 0;
  var taken = 0;
  var noBreak = 0;
  var late = 0;
  var onTime = 0;
  var early = 0;
  var overdue = 0;
  var lateAgents = [];
  var overdueAgents = [];
  var now = new Date();

  for (var i = 0; i < data.length; i++) {
    var r = data[i];
    if (!r[COL.AGENT]) continue;

    scheduled++;
    if (isValidDate(r[COL.DATE])) reportDate = r[COL.DATE];

    var hasOut = isValidDate(r[COL.ACTUAL_OUT]);
    var hasBack = isValidDate(r[COL.ACTUAL_BACK]);
    var variance = String(r[COL.VARIANCE] || "");

    if (hasOut) taken++; else noBreak++;

    if (variance.indexOf("Late by") === 0) {
      late++;
      lateAgents.push(r[COL.AGENT] + " (" + variance + ")");
    } else if (variance === "On Time") {
      onTime++;
    } else if (variance.indexOf("Early by") === 0) {
      early++;
    }

    var isOverdue = false;
    if (hasOut && !hasBack && isValidDate(r[COL.SCHEDULED_BACK]) &&
        now.getTime() > r[COL.SCHEDULED_BACK].getTime()) {
      isOverdue = true;
    } else if (r[COL.STATUS] === STATUS.BREAK_OVERDUE && hasOut && !hasBack) {
      isOverdue = true;
    }
    if (isOverdue) {
      overdue++;
      overdueAgents.push(r[COL.AGENT]);
    }

    rows.push([
      r[COL.AGENT],
      r[COL.SHIFT],
      r[COL.QUEUE],
      r[COL.BREAK_SLOT],
      r[COL.SCHEDULED_OUT],
      r[COL.SCHEDULED_BACK],
      r[COL.ACTUAL_OUT],
      r[COL.ACTUAL_BACK],
      r[COL.BREAK_USED],
      variance,
      isOverdue ? "Overdue" : r[COL.STATUS]
    ]);
  }

  // --- Write the Daily Report sheet ---
  var sh = ss.getSheetByName(SHEETS.DAILY_REPORT);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.DAILY_REPORT);
  }
  sh.clear();

  var dateLabel = formatDateKey(reportDate);
  sh.getRange(1, 1).setValue("DAILY REPORT — " + dateLabel)
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setFontSize(14);

  var summary = [
    ["Date", dateLabel],
    ["Agents scheduled", scheduled],
    ["Breaks taken (Actual Out recorded)", taken],
    ["No break recorded yet", noBreak],
    ["On time returns", onTime],
    ["Late returns", late],
    ["Early returns", early],
    ["Overdue right now", overdue],
    ["Late agents", lateAgents.length ? lateAgents.join("; ") : "None"],
    ["Overdue agents", overdueAgents.length ? overdueAgents.join("; ") : "None"]
  ];
  sh.getRange(2, 1, summary.length, 2).setValues(summary);
  sh.getRange(2, 1, summary.length, 1).setFontWeight("bold");

  var headerRow = summary.length + 3;
  var headers = [
    "Agent", "Shift", "Queue", "Break Slot",
    "Scheduled Out", "Scheduled Back",
    "Actual Out", "Actual Back",
    "Break Used (min)", "Variance", "Status"
  ];
  sh.getRange(headerRow, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  if (rows.length > 0) {
    sh.getRange(headerRow + 1, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(headerRow + 1, 5, rows.length, 4).setNumberFormat("h:mm AM/PM");
  }

  var widths = [250, 100, 140, 100, 140, 140, 140, 140, 150, 180, 150];
  widths.forEach(function(w, idx) {
    sh.setColumnWidth(idx + 1, w);
  });
  sh.setFrozenRows(headerRow);

  SpreadsheetApp.flush();

  var summaryText =
    "Daily Report — " + dateLabel + "\n\n" +
    "Scheduled: " + scheduled + "\n" +
    "Breaks taken: " + taken + "  •  No break recorded: " + noBreak + "\n" +
    "On time: " + onTime + "  •  Late: " + late + "  •  Early: " + early + "\n" +
    "Overdue right now: " + overdue +
    (lateAgents.length ? "\n\nLate:\n• " + lateAgents.join("\n• ") : "") +
    (overdueAgents.length ? "\n\nOverdue:\n• " + overdueAgents.join("\n• ") : "") +
    "\n\nFull table written to the 'Daily Report' sheet.";

  SpreadsheetApp.getUi().alert("Daily Report", summaryText, SpreadsheetApp.getUi().ButtonSet.OK);

  return {
    dateLabel: dateLabel,
    scheduled: scheduled,
    taken: taken,
    noBreak: noBreak,
    late: late,
    onTime: onTime,
    early: early,
    overdue: overdue
  };
}
