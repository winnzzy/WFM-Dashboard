/**
 * ==========================================================
 * BREAK AUTOMATION V2
 * ==========================================================
 * Manual break start/return triggered by Team Lead menu.
 *
 * NOTE: COL constants are 0-based (array index).
 *       Sheet getRange() is 1-based, so we use COL.xxx + 1.
 */

/**
 * Starts a break for the currently selected agent row.
 */
function startBreak() {

  var ss = SpreadsheetApp.getActive();
  var sh = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  var row = sh.getActiveCell().getRow();

  if (row < OPS_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("Select an agent row (row 6 or below).");
    return;
  }

  var statusCell = sh.getRange(row, COL.STATUS + 1).getValue();

  if (statusCell === STATUS.ON_BREAK) {
    SpreadsheetApp.getUi().alert("Agent is already on break.");
    return;
  }

  // --- Check break window ---
  var shift = sh.getRange(row, COL.SHIFT + 1).getValue();

  if (!isWithinBreakWindow(shift)) {

    var ui = SpreadsheetApp.getUi();

    var response = ui.alert(
      "Supervisor Override",
      "This agent is outside the approved break window.\n\nContinue anyway?",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return;
    }

    // Record override
    sh.getRange(row, COL.OVERRIDE + 1).setValue("YES");

    sh.getRange(row, COL.OVERRIDE_TIME + 1)
      .setValue(new Date())
      .setNumberFormat("h:mm AM/PM");

    sh.getRange(row, COL.REMARKS + 1)
      .setValue("Outside approved break window");
  }

  // --- Check same-queue conflict + Call queue coverage rule ---
  var queue = sh.getRange(row, COL.QUEUE + 1).getValue();
  var queueKey = normalizeQueueName(queue);

  var data = sh.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var sameQueue = false;
  var otherCallOnQueue = 0;

  for (var i = 0; i < data.length; i++) {

    var r = data[i];
    var currentRow = OPS_DATA_START_ROW + i;

    if (!r[COL.AGENT]) continue;

    var rowQueueKey = normalizeQueueName(r[COL.QUEUE]);

    if (rowQueueKey === queueKey && r[COL.STATUS] === STATUS.ON_BREAK) {
      sameQueue = true;
    }

    // For Call queue, require at least one OTHER Call agent to remain
    // actively working on queue while this agent is on break.
    if (
      queueKey === "call" &&
      currentRow !== row &&
      rowQueueKey === "call" &&
      r[COL.STATUS] === STATUS.ON_QUEUE
    ) {
      otherCallOnQueue++;
    }
  }

  if (shouldEnforceQueueCoverage(shift) && sameQueue) {
    SpreadsheetApp.getUi().alert(
      "Another agent on '" + queue + "' is already on break."
    );
    return;
  }

  if (queueKey === "call" && otherCallOnQueue < 1) {
    SpreadsheetApp.getUi().alert(
      "Call queue coverage rule: at least one other Call agent must remain On Queue before this break can start."
    );
    return;
  }

  // --- Set break start ---
  var now = new Date();

  sh.getRange(row, COL.ACTUAL_OUT + 1)
    .setValue(now)
    .setNumberFormat("h:mm AM/PM");

  sh.getRange(row, COL.STATUS + 1)
    .setValue(STATUS.ON_BREAK);

  refreshDashboard();
}


/**
 * Returns the currently selected agent from break.
 */
function returnFromBreak() {

  var ss = SpreadsheetApp.getActive();
  var sh = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  var row = sh.getActiveCell().getRow();

  if (row < OPS_DATA_START_ROW) {
    SpreadsheetApp.getUi().alert("Select an agent row (row 6 or below).");
    return;
  }

  var statusCell = sh.getRange(row, COL.STATUS + 1).getValue();

  if (statusCell !== STATUS.ON_BREAK) {
    SpreadsheetApp.getUi().alert("Agent is not on break.");
    return;
  }

  var actualBack = new Date();

  sh.getRange(row, COL.ACTUAL_BACK + 1)
    .setValue(actualBack)
    .setNumberFormat("h:mm AM/PM");

  sh.getRange(row, COL.STATUS + 1)
    .setValue(STATUS.ON_QUEUE);

  refreshDashboard();
}