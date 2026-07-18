/**
 * ==========================================================
 * WFM DASHBOARD
 * StatusEngine.gs
 * ==========================================================
 */

function updateAgentStatuses() {

  const ss = SpreadsheetApp.getActive();
  const ops = ss.getSheetByName("Daily Operations");

  if (!ops) {
    throw new Error("Daily Operations sheet not found.");
  }

  const range = ops.getRange("A6:O500");
  const data = range.getValues();

  const nowMinutes = toMinutes(new Date());

  data.forEach(row => {

    if (!row[COL.AGENT]) return;

    const scheduledOut = row[COL.SCHEDULED_OUT];
    const scheduledBack = row[COL.SCHEDULED_BACK];

    if (!isValidDate(scheduledOut) || !isValidDate(scheduledBack)) {
      row[COL.STATUS] = "";
      return;
    }

    const out = toMinutes(scheduledOut);
    const back = toMinutes(scheduledBack);

    if (nowMinutes < out) {

      row[COL.STATUS] = STATUS.ON_QUEUE;

    } else if (nowMinutes < back) {

      row[COL.STATUS] = STATUS.ON_BREAK;

    } else {

      row[COL.STATUS] = STATUS.BREAK_OVERDUE;

    }

  });

  range.setValues(data);

  refreshDashboard();

}