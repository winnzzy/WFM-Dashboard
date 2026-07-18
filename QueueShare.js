/**
 * ==========================================================
 * WFM DASHBOARD
 * QueueShare.gs
 * ==========================================================
 */

function updateQueueShare() {

  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName("Dashboard");
  const ops = ss.getSheetByName("Daily Operations");

  if (!dash || !ops) return;

  dash.getRange(DASHBOARD.QUEUE_SHARE).clearContent();

  const data = ops.getRange("A6:O500").getValues();

  const summary = {};

  QUEUES.forEach(queue => {
    summary[queue] = {
      assigned: 0,
      onQueue: 0,
      onBreak: 0
    };
  });

  data.forEach(row => {

    const agent = String(row[COL.AGENT]).trim();
    const queue = String(row[COL.QUEUE]).trim();
    const status = String(row[COL.STATUS]).trim();

    if (!agent) return;
    if (!summary[queue]) return;

    summary[queue].assigned++;

    if (status === STATUS.ON_QUEUE) {
      summary[queue].onQueue++;
    }

    if (
      status === STATUS.ON_BREAK ||
      status === STATUS.BREAK_OVERDUE
    ) {
      summary[queue].onBreak++;
    }

  });

  const output = [];

  QUEUES.forEach(queue => {

    const q = summary[queue];

    let coverage = "🟢 OK";

    if (q.onQueue === 0) {
      coverage = "🔴 CRITICAL";
    } else if (q.onQueue === 1) {
      coverage = "🟡 LOW";
    }

    output.push([
      queue,
      q.assigned,
      q.onQueue,
      q.onBreak,
      coverage
    ]);

  });

  dash
    .getRange(DASHBOARD.QUEUE_SHARE)
    .offset(0, 0, output.length, 5)
    .setValues(output);

}