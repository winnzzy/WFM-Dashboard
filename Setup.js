const CONFIG = {
  HEADER_COLOR: "#0F4C81",
  HEADER_TEXT: "#FFFFFF",
  BODY_COLOR: "#FFFFFF",
  ALT_ROW: "#F8F9FA",
  ROWS: 150,
  TEAM_LEAD: "Winner Nnamuah"
};

function buildDailyOperations() {

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Daily Operations");

  sh.clear();

  buildDailyHeader(sh);

 const headers = [
  "No",
  "Agent",
  "Center",
  "Supervisor",
  "Shift",
  "Queue",
  "Break Slot",
  "Scheduled Out",
  "Scheduled Back",
  "Login",
  "Actual Out",
  "Actual Back",
  "Break Used (mins)",
  "Variance",
  "Status",
  "Remarks",
"Override",
"Override Time"
];

  sh.getRange(5, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(CONFIG.HEADER_COLOR)
    .setFontColor(CONFIG.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.setFrozenRows(5);

  // Remove existing filter if present
if (sh.getFilter()) {
  sh.getFilter().remove();
}

// Create a new filter
sh.getRange("A5:R155").createFilter();

  const widths = [
  50,   // No
  220,  // Agent
  90,   // Center
  150,  // Supervisor
  80,   // Shift
  100,  // Queue
  90,   // Break Slot
  120,  // Scheduled Out
  120,  // Scheduled Back
  90,   // Login
  120,  // Actual Out
  120,  // Actual Back
  120,  // Break Used
  120,  // Variance
  120,  // Status
  250,  // Remarks
  120,  // Override
  140   // Override Time
];

widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

}

function buildDailyHeader(sh) {

  sh.getRange("A1:O1").merge();

  sh.getRange("A1")
    .setValue("CONTACT CENTER TEAM LEAD DAILY OPERATIONS")
    .setBackground(CONFIG.HEADER_COLOR)
    .setFontColor("white")
    .setFontSize(18)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("A2").setValue("Date");
  sh.getRange("B2").setFormula("=TODAY()");

  sh.getRange("D2").setValue("Team Lead");
  sh.getRange("E2").setValue(CONFIG.TEAM_LEAD);

  sh.getRange("H2").setValue("Center");
  sh.getRange("I2").setValue("Abuja");

  sh.getRange("K2").setValue("Generated");
  sh.getRange("L2").setFormula("=NOW()");

}