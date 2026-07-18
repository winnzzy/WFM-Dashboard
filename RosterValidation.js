/**
 * ===================================================
 * ROSTER VALIDATION
 * ===================================================
 */

function applyRosterValidation() {

  const ss = SpreadsheetApp.getActive();

  const roster = ss.getSheetByName("Agent Roster");

  // Shift Dropdown
  const shiftRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Morning","Afternoon","Night","OFF"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("F2:F300").setDataValidation(shiftRule);

  // Employment Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Active","Inactive","Leave"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("G2:G300").setDataValidation(statusRule);

  // Queue Preference
  const queueRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Auto","Call","Email","Clara","Ebanqo"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("H2:H300").setDataValidation(queueRule);

  // Break Group
  const breakRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Auto","A","B","C","D"],
      true
    )
    .setAllowInvalid(false)
    .build();

  roster.getRange("I2:I300").setDataValidation(breakRule);

}