/**
 * ==========================================================
 * LIVE BREAK MONITOR
 * ==========================================================
 */

function updateBreakMonitor() {

  const ss = SpreadsheetApp.getActive();

  const dash = ss.getSheetByName("Dashboard");
  const ops = ss.getSheetByName("Daily Operations");

  if (!dash || !ops) {
    throw new Error("Dashboard or Daily Operations sheet not found.");
  }

  // Clear previous monitor
  dash.getRange("A22:H200").clearContent();

  // Title
  dash.getRange("A22:H22")
    .merge()
    .setValue("LIVE BREAK MONITOR")
    .setBackground("#0F4C81")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Headers
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
    .setBackground("#D9EAD3")
    .setFontWeight("bold");

  const data = ops.getRange("A6:R500").getValues();
  const output = [];
  const now = new Date();

  data.forEach(row => {

    const status = row[COL.STATUS];

    if (
      status !== STATUS.ON_BREAK &&
      status !== STATUS.BREAK_OVERDUE
    ) return;

    let minutesLeft = "";

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

  });

  if (output.length) {

    dash.getRange(24, 1, output.length, 8).setValues(output);

    dash.getRange(24, 4, output.length, 2)
      .setNumberFormat("h:mm AM/PM");

  } else {

    dash.getRange("A24")
      .setValue("No agents are currently on break.");

  }

}