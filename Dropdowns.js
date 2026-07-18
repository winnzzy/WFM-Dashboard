function applyDropdowns() {

  const ss = SpreadsheetApp.getActive();

  const sh = ss.getSheetByName("Daily Operations");
  const settings = ss.getSheetByName("Settings");

  // Center
  const centerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange("D2:D3"), true)
    .setAllowInvalid(false)
    .build();

  // Shift
  const shiftRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange("B2:B5"), true)
    .setAllowInvalid(false)
    .build();

  // Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange("C2:C10"), true)
    .setAllowInvalid(false)
    .build();

  // Supervisor
  const supervisorRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(settings.getRange("E2:E5"), true)
    .setAllowInvalid(false)
    .build();

  // Queue (Fixed List)
  const queueRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      [
        "Call",
        "Email",
        "Clara",
        "Ebanqo",
        "Auto"
      ],
      true
    )
    .setAllowInvalid(false)
    .build();

  // Apply validations

  // Center
  sh.getRange("C6:C500").clearDataValidations().setDataValidation(centerRule);

  // Supervisor
  sh.getRange("D6:D500").clearDataValidations().setDataValidation(supervisorRule);

  // Shift
  sh.getRange("E6:E500").clearDataValidations().setDataValidation(shiftRule);

  // Queue
  sh.getRange("F6:F500").clearDataValidations().setDataValidation(queueRule);

  // Status (Column O, not N)
  sh.getRange("O6:O500").clearDataValidations().setDataValidation(statusRule);

}