/**
 * ==========================================
 * AGENT DROPDOWN
 * ==========================================
 */

function applyAgentDropdown(){

  const ss = SpreadsheetApp.getActive();

  const roster = ss.getSheetByName("Agent Roster");

  const sh = ss.getSheetByName("Daily Operations");

  const names = roster.getRange("A2:A100");

  const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(names, true)
      .setAllowInvalid(false)
      .build();

  sh.getRange("B6:B155")
    .setDataValidation(rule);

}
