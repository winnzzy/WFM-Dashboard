/**
 * ==========================================================
 * AGENT ROSTER
 * ==========================================================
 * Creates and configures the Agent Roster sheet.
 */

/**
 * Builds the Agent Roster sheet with headers, default values,
 * and sample agent data. Uses shared constants from Utilities.js.
 */
function buildRoster() {

  var ss = SpreadsheetApp.getActive();

  var sh = ss.getSheetByName(SHEETS.AGENT_ROSTER);

  if (!sh) {
    sh = ss.insertSheet(SHEETS.AGENT_ROSTER);
  }

  sh.clear();

  // --- Headers ---
  var headers = [
    "Agent Name",
    "Center",
    "Supervisor",
    "Phone",
    "Email",
    "Shift",
    "Employment Status",
    "Queue Preference",
    "Break Group",
    "Remarks"
  ];

  sh.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground(COLORS.HEADER_BG)
    .setFontColor(COLORS.HEADER_TEXT)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // --- Default Employment Status (Active) ---
  var statusDefaults = [];
  for (var i = 0; i < 11; i++) {
    statusDefaults.push(["Active"]);
  }
  sh.getRange(2, 7, statusDefaults.length, 1).setValues(statusDefaults);

  // --- Default Queue Preference (Auto) ---
  var queueDefaults = [];
  for (var j = 0; j < 11; j++) {
    queueDefaults.push(["Auto"]);
  }
  sh.getRange(2, 8, queueDefaults.length, 1).setValues(queueDefaults);

  // --- Clear Break Group and Remarks ---
  sh.getRange("I2:I100").clearContent();
  sh.getRange("J2:J100").clearContent();

  // --- Default Shifts ---
  var shifts = [
    ["Morning"],
    ["Afternoon"],
    ["Afternoon"],
    ["Night"],
    ["Morning"],
    ["Morning"],
    ["Night"],
    ["Afternoon"],
    ["Morning"],
    ["OFF"],
    ["Morning"]
  ];
  sh.getRange(2, 6, shifts.length, 1).setValues(shifts);

  // --- Column widths ---
  var widths = [
    250, 100, 180, 140, 220,
    100, 140, 140, 120, 250
  ];
  widths.forEach(function(w, i) {
    sh.setColumnWidth(i + 1, w);
  });

  sh.setFrozenRows(1);

  // --- Sample agent names ---
  var agents = [
    ["Yimaumuaju Joshua"],
    ["Shehu Usman Daniya"],
    ["Nelson Okeah"],
    ["Mariam Ajikeola Abdulazeez"],
    ["Comfort Ijeoma Okereke"],
    ["Thomas Goodeness Dilah"],
    ["Subuhanallahi Abiodun Adeleye"],
    ["Emmanuel Okahena Owobu"],
    ["Isaac Idemeto"],
    ["Zainab Yusuf"],
    ["Cynthia Nkechi Nwadukwe"]
  ];
  sh.getRange(2, 1, agents.length, 1).setValues(agents);
}
