/**
 * ==========================================================
 * WEEKLY SCHEDULE SHORTCUTS
 * ==========================================================
 * The Weekly Schedule sheet is a recurring Mon–Sun template
 * (one row per agent, columns B–H), not a dated week — so
 * "copy last week" does not apply here. What does save time
 * is filling a whole row at once instead of cell by cell.
 *
 * All three shortcuts act on the agent in the currently
 * selected row of the Weekly Schedule sheet:
 *   • Set Selected Agent's Whole Week… — prompt for one
 *     shift and fill Mon–Sun with it.
 *   • Set Selected Agent's Week from Roster — fill Mon–Sun
 *     with that agent's Roster shift (useful for a newly
 *     added agent: add them in Roster, sync, then one tap).
 *   • Clear Selected Agent's Week — blank Mon–Sun, after
 *     confirmation. Nothing else on the row is touched.
 *
 * Roster rows are never changed by these shortcuts, and
 * other agents' rows are never touched.
 */

/**
 * Returns the agent name in the active Weekly Schedule row,
 * or "" with an alert if the selection is not usable.
 */
function getSelectedWeeklyAgent() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var activeSheet = ss.getActiveSheet();

  if (!activeSheet || activeSheet.getName() !== SHEETS.WEEKLY_SCHEDULE) {
    ui.alert("Weekly Shortcut", "Select a row on the Weekly Schedule sheet first.", ui.ButtonSet.OK);
    return "";
  }
  var ws = activeSheet;
  var row = ss.getActiveCell().getRow();
  if (row < WEEKLY_DATA_START_ROW) {
    ui.alert("Weekly Shortcut", "Select an agent row (row 2 or below) on Weekly Schedule.", ui.ButtonSet.OK);
    return "";
  }
  var name = String(ws.getRange(row, 1).getValue() || "").trim();
  if (!name) {
    ui.alert("Weekly Shortcut", "The selected row has no agent name in column A.", ui.ButtonSet.OK);
    return "";
  }
  return name;
}

function setSelectedAgentWeek() {
  var ui = SpreadsheetApp.getUi();
  var name = getSelectedWeeklyAgent();
  if (!name) return;

  var response = ui.prompt(
    "Set Whole Week — " + name,
    "Enter one shift for Mon–Sun:\nMorning, Afternoon, Night, OFF, or Leave",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  var shift = String(response.getResponseText() || "").trim();
  // Accept case-insensitive input, store the canonical value.
  var canonical = "";
  WEEKLY_SHIFTS.forEach(function(s) {
    if (s.toLowerCase() === shift.toLowerCase()) canonical = s;
  });
  if (!canonical) {
    ui.alert("Weekly Shortcut", "Invalid shift '" + shift + "'. Use Morning, Afternoon, Night, OFF, or Leave.", ui.ButtonSet.OK);
    return;
  }

  var ws = SpreadsheetApp.getActive().getSheetByName(SHEETS.WEEKLY_SCHEDULE);
  var row = ws.getActiveCell().getRow();
  ws.getRange(row, 2, 1, 7).setValues([[canonical, canonical, canonical, canonical, canonical, canonical, canonical]]);
  ui.alert("Weekly Shortcut", name + ": Mon–Sun set to " + canonical + ".", ui.ButtonSet.OK);
}

function setSelectedAgentWeekFromRoster() {
  var ui = SpreadsheetApp.getUi();
  var name = getSelectedWeeklyAgent();
  if (!name) return;

  var contacts = getRosterMap();
  var agent = contacts[name];
  if (!agent) {
    ui.alert("Weekly Shortcut", name + " is not on the Agent Roster. Add them to the Roster first, then sync.", ui.ButtonSet.OK);
    return;
  }
  var shift = String(agent.shift || "").trim();
  if (WEEKLY_SHIFTS.indexOf(shift) === -1) {
    ui.alert("Weekly Shortcut", name + "'s Roster shift ('" + shift + "') is not a Weekly shift. Set it in the Roster to Morning, Afternoon, Night, OFF, or Leave first.", ui.ButtonSet.OK);
    return;
  }

  var ws = SpreadsheetApp.getActive().getSheetByName(SHEETS.WEEKLY_SCHEDULE);
  var row = ws.getActiveCell().getRow();
  ws.getRange(row, 2, 1, 7).setValues([[shift, shift, shift, shift, shift, shift, shift]]);
  ui.alert("Weekly Shortcut", name + ": Mon–Sun set to Roster shift (" + shift + "). Change individual days afterwards if needed.", ui.ButtonSet.OK);
}

function clearSelectedAgentWeek() {
  var ui = SpreadsheetApp.getUi();
  var name = getSelectedWeeklyAgent();
  if (!name) return;

  var response = ui.alert(
    "Clear Week — " + name,
    "Clear Mon–Sun for " + name + "? This blanks the 7 day cells only.",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  var ws = SpreadsheetApp.getActive().getSheetByName(SHEETS.WEEKLY_SCHEDULE);
  var row = ws.getActiveCell().getRow();
  ws.getRange(row, 2, 1, 7).clearContent();
  ui.alert("Weekly Shortcut", name + ": Mon–Sun cleared.", ui.ButtonSet.OK);
}
