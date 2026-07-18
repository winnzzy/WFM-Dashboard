/**
 * ==========================================================
 * DROPDOWNS
 * ==========================================================
 * Applies data validation dropdowns to the Daily Operations
 * sheet. Validation ranges are generated dynamically from
 * the Settings sheet to automatically expand when new values
 * are added.
 */

/**
 * Applies all data validation rules to the Daily Operations sheet.
 * Ranges are determined dynamically — never hardcoded.
 */
function applyDropdowns() {

  var ss = SpreadsheetApp.getActive();

  var sh = getSheetOrThrow(ss, SHEETS.DAILY_OPS);
  var settings = getSheetOrThrow(ss, SHEETS.SETTINGS);

  // --- Helper: find the last non-empty row in a column ---
  function getLastPopulatedRow(sheet, col) {
    var values = sheet.getRange(1, col, sheet.getLastRow(), 1).getValues();
    for (var i = values.length - 1; i >= 0; i--) {
      if (values[i][0] !== "") return i + 1; // 1-based row
    }
    return 1; // header row only
  }

  // --- Dynamic ranges from Settings sheet ---
  var lastStatusRow = getLastPopulatedRow(settings, 3);  // Column C = Status
  var lastCenterRow = getLastPopulatedRow(settings, 4);  // Column D = Centers
  var lastShiftRow = getLastPopulatedRow(settings, 2);   // Column B = Shifts
  var lastSuperRow = getLastPopulatedRow(settings, 5);   // Column E = Supervisors

  // --- Build validation rules ---

  // Center
  var centerRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      settings.getRange(2, 4, lastCenterRow - 1, 1),
      true
    )
    .setAllowInvalid(false)
    .build();

  // Shift
  var shiftRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      settings.getRange(2, 2, lastShiftRow - 1, 1),
      true
    )
    .setAllowInvalid(false)
    .build();

  // Status
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      settings.getRange(2, 3, lastStatusRow - 1, 1),
      true
    )
    .setAllowInvalid(false)
    .build();

  // Supervisor
  var supervisorRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(
      settings.getRange(2, 5, lastSuperRow - 1, 1),
      true
    )
    .setAllowInvalid(false)
    .build();

  // Queue (fixed list — includes "Auto" for Night shift)
  var queueRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(
      ["Call", "Email", "Clara", "Ebanqo", "Auto"],
      true
    )
    .setAllowInvalid(false)
    .build();

  // --- Apply validations to Daily Operations ---
  // Center (Column C)
  sh.getRange("C" + OPS_DATA_START_ROW + ":C" + MAX_OPS_ROWS)
    .clearDataValidations()
    .setDataValidation(centerRule);

  // Supervisor (Column D)
  sh.getRange("D" + OPS_DATA_START_ROW + ":D" + MAX_OPS_ROWS)
    .clearDataValidations()
    .setDataValidation(supervisorRule);

  // Shift (Column E)
  sh.getRange("E" + OPS_DATA_START_ROW + ":E" + MAX_OPS_ROWS)
    .clearDataValidations()
    .setDataValidation(shiftRule);

  // Queue (Column F)
  sh.getRange("F" + OPS_DATA_START_ROW + ":F" + MAX_OPS_ROWS)
    .clearDataValidations()
    .setDataValidation(queueRule);

  // Status (Column O)
  sh.getRange("O" + OPS_DATA_START_ROW + ":O" + MAX_OPS_ROWS)
    .clearDataValidations()
    .setDataValidation(statusRule);
}