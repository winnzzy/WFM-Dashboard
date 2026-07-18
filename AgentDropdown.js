/**
 * ==========================================================
 * AGENT DROPDOWN
 * ==========================================================
 * Applies agent name validation dropdown to the Agent column
 * in Daily Operations. Uses shared constants from Utilities.js.
 */

/**
 * Applies agent name dropdown validation to Daily Operations
 * column B, sourced from the Agent Roster sheet.
 */
function applyAgentDropdown() {

  var ss = SpreadsheetApp.getActive();

  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);
  var ops = getSheetOrThrow(ss, SHEETS.DAILY_OPS);

  // Source range: Agent Roster column A
  var names = roster.getRange("A2:A" + MAX_ROSTER_ROWS);

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(names, true)
    .setAllowInvalid(false)
    .build();

  // Apply to Daily Operations column B (Agent)
  ops.getRange(
    "B" + OPS_DATA_START_ROW + ":B" + MAX_OPS_ROWS
  ).setDataValidation(rule);
}