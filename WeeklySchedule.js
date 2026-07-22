/**
 * ==========================================================
 * WEEKLY SCHEDULE
 * ==========================================================
 * Builds the Weekly Schedule sheet where Team Leads enter
 * only shift assignments (Morning / Afternoon / Night / OFF / Leave)
 * for each agent per weekday.
 *
 * The Weekly Schedule is the primary input for daily generation.
 * Everything else (queue, break slot, times) is auto-generated.
 */

/**
 * Builds the Weekly Schedule sheet with headers,
 * agent names from roster, and dropdown validation.
 */
function buildWeeklySchedule() {

  var ss = SpreadsheetApp.getActive();

  var ws = ss.getSheetByName(SHEETS.WEEKLY_SCHEDULE);

  if (!ws) {
    ws = ss.insertSheet(SHEETS.WEEKLY_SCHEDULE);
  }

  ws.clear();

  // --- Headers ---
  var headers = [
    "Agent Name",
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday"
  ];

  ws.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  ws.setFrozenRows(1);
  ws.setFrozenColumns(1);

  // --- Populate agent names from Agent Roster ---
  var roster = ss.getSheetByName(SHEETS.AGENT_ROSTER);

  if (roster) {

    var agentNames = roster.getRange(
      ROSTER_DATA_START_ROW,
      ROSTER_COL.NAME + 1,
      MAX_ROSTER_ROWS,
      1
    ).getValues();

    var names = [];

    for (var i = 0; i < agentNames.length; i++) {
      if (agentNames[i][0]) {
        names.push([agentNames[i][0]]);
      }
    }

    if (names.length > 0) {
      ws.getRange(2, 1, names.length, 1).setValues(names);
    }
  }

  // --- Apply shift dropdown validation (columns B–H) ---
  var shiftRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      WEEKLY_SHIFTS,
      true
    )
    .setAllowInvalid(false)
    .build();

  ws.getRange(
    WEEKLY_DATA_START_ROW,
    2,
    MAX_WEEKLY_ROWS,
    7
  )
    .clearDataValidations()
    .setDataValidation(shiftRule);

  // --- Column widths ---
  var widths = [250, 120, 120, 120, 120, 120, 120, 120];
  widths.forEach(function(w, idx) {
    ws.setColumnWidth(idx + 1, w);
  });
}