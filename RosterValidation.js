/**
 * ==========================================================
 * ROSTER VALIDATION
 * ==========================================================
 * Applies data validation dropdowns to the Agent Roster sheet.
 * Uses shared SHEETS constant from Utilities.js.
 */

/**
 * Applies all validation rules to Agent Roster columns F–I.
 */
function applyRosterValidation() {

  var ss = SpreadsheetApp.getActive();

  var roster = getSheetOrThrow(ss, SHEETS.AGENT_ROSTER);

  // Shift Dropdown (Column F)
  var shiftRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Morning", "Afternoon", "Night", "OFF"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("F2:F300").setDataValidation(shiftRule);

  // Employment Status (Column G)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Active", "Inactive", "Leave"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("G2:G300").setDataValidation(statusRule);

  // Queue Preference (Column H)
  var queueRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Auto", "Call", "Email", "Clara", "Ebanqo"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("H2:H300").setDataValidation(queueRule);

  // Break Group (Column I)
  var breakRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Auto", "A", "B", "C", "D"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("I2:I300").setDataValidation(breakRule);
}