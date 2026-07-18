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

  // --- Check max 2 agents on break + same-queue conflict ---
  var queue = sh.getRange(row, COL.QUEUE + 1).getValue();

  var data = sh.getRange(
    OPS_DATA_START_ROW,
    1,
    MAX_OPS_ROWS,
    OPS_COL_COUNT
  ).getValues();

  var agentsOnBreak = 0;
  var sameQueue = false;

  for (var i = 0; i < data.length; i++) {

    var r = data[i];

    if (r[COL.STATUS] !== STATUS.ON_BREAK) continue;

    agentsOnBreak++;

    if (r[COL.QUEUE] === queue) {
      sameQueue = true;
    }
  }

  if (agentsOnBreak >= 2) {
    SpreadsheetApp.getUi().alert(
      "Maximum of 2 agents are already on break."
    );
    return;
  }

  if (sameQueue) {
    SpreadsheetApp.getUi().alert(
      "Another agent on '" + queue + "' is already on break."
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

  var actualOut = sh.getRange(row, COL.ACTUAL_OUT + 1).getValue();

  var scheduledBack = sh.getRange(row, COL.SCHEDULED_BACK + 1).getValue();

  var breakMinutes = Math.round(
    (actualBack - actualOut) / 60000
  );

  sh.getRange(row, COL.BREAK_USED + 1)
    .setValue(breakMinutes);

  var variance = Math.round(
    (actualBack - scheduledBack) / 60000
  );

  if (variance < 0) {
    sh.getRange(row, COL.VARIANCE + 1)
      .setValue("Early by " + Math.abs(variance) + " mins");
  } else if (variance === 0) {
    sh.getRange(row, COL.VARIANCE + 1)
      .setValue("On Time");
  } else {
    sh.getRange(row, COL.VARIANCE + 1)
      .setValue("Late by " + variance + " mins");
  }

  sh.getRange(row, COL.STATUS + 1)
    .setValue(STATUS.ON_QUEUE);

  refreshDashboard();
}