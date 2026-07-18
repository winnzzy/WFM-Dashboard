/**
 * ==========================================
 * AGENT ROSTER
 * ==========================================
 */

function buildRoster() {

  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName("Agent Roster");

  if (!sh) {
    sh = ss.insertSheet("Agent Roster");
  }

  sh.clear();

  const headers = [

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
sh.getRange(1,1,1,headers.length)
  .setValues([headers])
  .setBackground("#0F4C81")
  .setFontColor("white")
  .setFontWeight("bold")
  .setHorizontalAlignment("center");

// Employment Status
sh.getRange("G2:G100").setValue("Active");

const queueDefaults = [
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"],
  ["Auto"]
];

sh.getRange(2,8,queueDefaults.length,1).setValues(queueDefaults);

// Break Group (leave blank)
sh.getRange("I2:I100").clearContent();

// Remarks (leave blank)
sh.getRange("J2:J100").clearContent();

// Default Shifts
const shifts = [
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

sh.getRange(2,6,shifts.length,1).setValues(shifts);
  const widths = [
  250, // Agent
  100, // Center
  180, // Supervisor
  140, // Phone
  220, // Email
  100, // Shift
  140, // Employment Status
  140, // Queue Preference
  120, // Break Group
  250  // Remarks
];
  widths.forEach((w,i)=>{
    sh.setColumnWidth(i+1,w);
  });

  sh.setFrozenRows(1);

  sh.getRange("A2").setValue("Yimaumuaju Joshua");
  sh.getRange("A3").setValue("Shehu Usman Daniya");
  sh.getRange("A4").setValue("Muhammad Nazir");
  sh.getRange("A5").setValue("Mariam Ajikeola Abdulazeez");
  sh.getRange("A6").setValue("Comfort Ijeoma Okereke");
  sh.getRange("A7").setValue("Thomas Goodeness Dilah");
  sh.getRange("A8").setValue("Subuhanallahi Abiodun Adeleye");
  sh.getRange("A9").setValue("Emmanuel Okahena Owobu");
  sh.getRange("A10").setValue("Firdausi Suleiman Sanni");
  sh.getRange("A11").setValue("Zainab Yusuf");
  sh.getRange("A12").setValue("Cynthia Nkechi Nwadukwe");

  

}