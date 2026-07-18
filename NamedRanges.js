/**
 * ==========================================================
 * NAMED RANGES
 * ==========================================================
 * Creates named ranges for all Daily Operations columns.
 * These are for user convenience in the spreadsheet UI.
 */

/**
 * Creates named ranges mapping column letters to descriptive names.
 * Uses SHEETS constant from Utilities.js.
 */
function createNamedRanges() {

  var ss = SpreadsheetApp.getActive();
  var sh = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Row 6 = data start, row 500 = max data row
  var range = function(col) {
    return sh.getRange(
      OPS_DATA_START_ROW,
      col,
      MAX_OPS_ROWS - OPS_DATA_START_ROW + 1,
      1
    );
  };

  ss.setNamedRange("DateColumn",           range(1));   // A
  ss.setNamedRange("AgentColumn",          range(2));   // B
  ss.setNamedRange("CenterColumn",         range(3));   // C
  ss.setNamedRange("SupervisorColumn",     range(4));   // D
  ss.setNamedRange("ShiftColumn",          range(5));   // E
  ss.setNamedRange("QueueColumn",          range(6));   // F
  ss.setNamedRange("BreakSlotColumn",      range(7));   // G
  ss.setNamedRange("ScheduledOutColumn",   range(8));   // H
  ss.setNamedRange("ScheduledBackColumn",  range(9));   // I
  ss.setNamedRange("LoginColumn",          range(10));  // J
  ss.setNamedRange("ActualOutColumn",      range(11));  // K
  ss.setNamedRange("ActualBackColumn",     range(12));  // L
  ss.setNamedRange("BreakUsedColumn",      range(13));  // M
  ss.setNamedRange("VarianceColumn",       range(14));  // N
  ss.setNamedRange("StatusColumn",         range(15));  // O
  ss.setNamedRange("RemarksColumn",        range(16));  // P
  ss.setNamedRange("OverrideColumn",       range(17));  // Q
  ss.setNamedRange("OverrideTimeColumn",   range(18));  // R
}