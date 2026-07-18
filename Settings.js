/**
 * ===============================================
 * SETTINGS
 * ===============================================
 */

function buildSettings() {

  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName("Settings");

  if (!sh) {
    sh = ss.insertSheet("Settings");
  }

  sh.clear();

  // QUEUES
  sh.getRange("A1").setValue("Queues").setFontWeight("bold");
  sh.getRange("A2:A5").setValues([
    ["Email"],
    ["Clara"],
    ["Ebanqo"],
    ["Call"]
  ]);

  // SHIFTS
  sh.getRange("B1").setValue("Shifts").setFontWeight("bold");
  sh.getRange("B2:B5").setValues([
    ["Morning"],
    ["Afternoon"],
    ["Night"],
    ["OFF"]
  ]);

  // STATUS
  sh.getRange("C1").setValue("Status").setFontWeight("bold");
  sh.getRange("C2:C11").setValues([
    ["On Queue"],
    ["On Break"],
    ["Break Overdue"],
    ["Meeting"],
    ["Coaching"],
    ["Training"],
    ["Offline"],
    ["Logged Out"],
    ["OFF"],
    ["Leave"]
  ]);

  // CENTERS
  sh.getRange("D1").setValue("Centers").setFontWeight("bold");
  sh.getRange("D2:D3").setValues([
    ["Abuja"],
    ["Lagos"]
  ]);

  // SUPERVISORS
  sh.getRange("E1").setValue("Supervisors").setFontWeight("bold");
  sh.getRange("E2:E5").setValues([
    ["Winner Nnamuah"],
    ["Grace"],
    ["Team Lead"],
    ["Manager"]
  ]);

  sh.hideSheet();

}