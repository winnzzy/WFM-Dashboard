/**
 * ==========================================================
 * ROSTER ↔ WEEKLY SCHEDULE SYNC
 * ==========================================================
 * This is the "no-code add/remove" bridge.
 *
 * ADD: type the agent in Agent Roster (name + shift + queue
 *   + email + phone), then run "Sync Roster → Weekly
 *   Schedule" — or just run "Generate Today's Schedule",
 *   which now does the same add-missing step first. The
 *   agent appears in Weekly Schedule with blank day cells;
 *   fill in their Mon–Sun shifts there, then generate.
 *
 * REMOVE, two equivalent ways (pick one):
 *   a) Set Employment Status to "Inactive" in the Roster.
 *      Generation skips them immediately, even if Weekly
 *      still lists a shift for them. Nothing is deleted.
 *   b) Delete the Roster row AND run "Remove Weekly Rows
 *      Not on Roster" (with confirmation) to drop their
 *      Weekly row too. Generation also skips agents who
 *      are in Weekly but no longer on the Roster, with a
 *      warning listing their names, instead of throwing.
 *
 * Nothing here deletes Roster rows, and nothing here
 * overwrites shifts you already set in Weekly Schedule.
 */

/**
 * Appends every Roster agent missing from Weekly Schedule
 * as a new row with blank Mon–Sun cells. Existing Weekly
 * rows and shifts are untouched. Returns the names added.
 */
function ensureWeeklyHasRosterAgents() {

  var ss = SpreadsheetApp.getActive();
  var ws = getSheetOrThrow(ss, SHEETS.WEEKLY_SCHEDULE);

  var weeklyData = ws.getRange(
    WEEKLY_DATA_START_ROW,
    1,
    MAX_WEEKLY_ROWS,
    8
  ).getValues();

  var weeklyNames = {};
  var lastRow = WEEKLY_DATA_START_ROW - 1;
  for (var i = 0; i < weeklyData.length; i++) {
    var n = String(weeklyData[i][0] || "").trim();
    if (n) {
      weeklyNames[n] = true;
      lastRow = WEEKLY_DATA_START_ROW + i;
    }
  }

  var added = [];
  getRosterAgents().forEach(function(agent) {
    if (weeklyNames[agent.name]) return;
    lastRow++;
    ws.getRange(lastRow, 1, 1, 8).setValues([[agent.name, "", "", "", "", "", "", ""]]);
    weeklyNames[agent.name] = true;
    added.push(agent.name);
  });

  if (added.length > 0) {
    // New rows need the same Mon–Sun shift dropdowns.
    var shiftRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(WEEKLY_SHIFTS, true)
      .setAllowInvalid(false)
      .build();
    ws.getRange(WEEKLY_DATA_START_ROW, 2, MAX_WEEKLY_ROWS, 7)
      .setDataValidation(shiftRule);
  }

  return added;
}

/**
 * Menu wrapper: add missing Roster agents to Weekly, then
 * report roster problems (duplicates, bad emails) if any.
 */
function syncRosterToWeeklyMenu() {

  var added = ensureWeeklyHasRosterAgents();
  var problems = validateRoster();

  var msg = "Roster → Weekly Schedule sync complete.\n\n";
  msg += added.length > 0
    ? "Added to Weekly Schedule:\n• " + added.join("\n• ") + "\n\nFill in their Mon–Sun shifts, then generate the day's schedule."
    : "No new agents to add — Weekly Schedule already covers every Roster name.";
  if (problems.length > 0) {
    msg += "\n\nRoster issues to fix (scheduling still works, notifications may skip these):\n• " + problems.join("\n• ");
  }

  SpreadsheetApp.getUi().alert("Roster Sync", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Deletes Weekly Schedule rows whose agent is no longer on
 * the Roster at all. Inactive agents are NOT deleted by this
 * (use status for a reversible removal). Asks first.
 * Returns the names removed.
 */
function pruneWeeklyFromRosterMenu() {

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  var ws = getSheetOrThrow(ss, SHEETS.WEEKLY_SCHEDULE);

  var rosterNames = {};
  getRosterAgents().forEach(function(a) { rosterNames[a.name] = true; });

  var weeklyData = ws.getRange(
    WEEKLY_DATA_START_ROW,
    1,
    MAX_WEEKLY_ROWS,
    8
  ).getValues();

  var rowsToDelete = [];
  for (var i = 0; i < weeklyData.length; i++) {
    var name = String(weeklyData[i][0] || "").trim();
    if (name && !rosterNames[name]) {
      rowsToDelete.push({ name: name, row: WEEKLY_DATA_START_ROW + i });
    }
  }

  if (rowsToDelete.length === 0) {
    ui.alert("Weekly Cleanup", "Nothing to remove — every Weekly Schedule agent is on the Roster.", ui.ButtonSet.OK);
    return [];
  }

  var response = ui.alert(
    "Remove Weekly Rows Not on Roster?",
    "These agents are in Weekly Schedule but NOT in Agent Roster:\n\n• " +
    rowsToDelete.map(function(r) { return r.name; }).join("\n• ") +
    "\n\nDelete their Weekly Schedule rows? Their Roster row is already gone, so this only removes the leftover Weekly row.",
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return [];

  // Delete bottom-up so row numbers stay valid.
  rowsToDelete.sort(function(a, b) { return b.row - a.row; });
  rowsToDelete.forEach(function(r) { ws.deleteRow(r.row); });

  ui.alert(
    "Weekly Cleanup",
    "Removed from Weekly Schedule:\n• " +
    rowsToDelete.map(function(r) { return r.name; }).join("\n• "),
    ui.ButtonSet.OK
  );
  return rowsToDelete.map(function(r) { return r.name; });
}
