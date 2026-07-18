/**
 * ==========================================================
 * DASHBOARD
 * ==========================================================
 */

function buildDashboardSheet() {

  const ss = SpreadsheetApp.getActive();

  let sh = ss.getSheetByName("Dashboard");

  if (!sh) {
    sh = ss.insertSheet("Dashboard");
  }

  sh.clear();
  sh.setHiddenGridlines(true);

  //==========================================================
  // TITLE
  //==========================================================

  sh.getRange("A1:P2")
    .merge()
    .setValue("CONTACT CENTER TEAM LEAD DASHBOARD")
    .setBackground("#0F4C81")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(20)
    .setHorizontalAlignment("center");

  sh.getRange("N3").setValue("Today");

  sh.getRange("O3")
    .setFormula("=TODAY()")
    .setNumberFormat("ddd dd-mmm-yyyy");

  //==========================================================
  // KPI CARDS
  //==========================================================

  createKPICard(
    sh,
    "A4",
    "Present",
    "=COUNTA('Daily Operations'!B6:B300)"
  );

  createKPICard(
    sh,
    "D4",
    "On Queue",
    "=COUNTIF('Daily Operations'!O6:O300,\"On Queue\")"
  );

  createKPICard(
    sh,
    "G4",
    "On Break",
    "=COUNTIF('Daily Operations'!O6:O300,\"On Break\")"
  );

  createKPICard(
    sh,
    "J4",
    "OFF",
    "=COUNTIFS('Agent Roster'!F2:F300,\"OFF\",'Agent Roster'!G2:G300,\"Active\")"
  );

  createKPICard(
    sh,
    "M4",
    "Overrides",
    "=COUNTIF('Daily Operations'!Q6:Q300,\"YES\")"
  );

  createKPICard(
    sh,
    "A8",
    "Late Returns",
    "=COUNTIF('Daily Operations'!N6:N300,\"Late*\")"
  );

  createKPICard(
    sh,
    "D8",
    "Early Returns",
    "=COUNTIF('Daily Operations'!N6:N300,\"Early*\")"
  );

  createKPICard(
    sh,
    "G8",
    "Morning",
    "=COUNTIF('Daily Operations'!E6:E300,\"Morning\")"
  );

  createKPICard(
    sh,
    "J8",
    "Afternoon",
    "=COUNTIF('Daily Operations'!E6:E300,\"Afternoon\")"
  );

  createKPICard(
    sh,
    "M8",
    "Night",
    "=COUNTIF('Daily Operations'!E6:E300,\"Night\")"
  );

  //==========================================================
  // QUEUE COVERAGE
  //==========================================================

  sh.getRange("A13:E13")
    .merge()
    .setValue("QUEUE COVERAGE")
    .setBackground("#0F4C81")
    .setFontColor("white")
    .setFontWeight("bold");

  sh.getRange("A14:B17").setValues([

    [
      "Call",
      "=COUNTIF('Daily Operations'!F6:F300,\"Call\")"
    ],

    [
      "Email",
      "=COUNTIF('Daily Operations'!F6:F300,\"Email\")"
    ],

    [
      "Clara",
      "=COUNTIF('Daily Operations'!F6:F300,\"Clara\")"
    ],

    [
      "Ebanqo",
      "=COUNTIF('Daily Operations'!F6:F300,\"Ebanqo\")"
    ]

  ]);

  //==========================================================
  // SHIFT COVERAGE
  //==========================================================

  sh.getRange("G13:K13")
    .merge()
    .setValue("SHIFT COVERAGE")
    .setBackground("#0F4C81")
    .setFontColor("white")
    .setFontWeight("bold");

  sh.getRange("G14:H17").setValues([

    [
      "Morning",
      "=COUNTIF('Daily Operations'!E6:E300,\"Morning\")"
    ],

    [
      "Afternoon",
      "=COUNTIF('Daily Operations'!E6:E300,\"Afternoon\")"
    ],

    [
      "Night",
      "=COUNTIF('Daily Operations'!E6:E300,\"Night\")"
    ],

    [
      "OFF",
      "=COUNTIFS('Agent Roster'!F2:F300,\"OFF\",'Agent Roster'!G2:G300,\"Active\")"
    ]

  ]);
    //==========================================================
  // LIVE BREAK MONITOR
  //==========================================================

  sh.getRange("A21:H21")
    .merge()
    .setValue("LIVE BREAK MONITOR")
    .setBackground("#0F4C81")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("A22:H22").setValues([[
    "Agent",
    "Queue",
    "Break Slot",
    "Actual Out",
    "Expected Back",
    "Minutes Left",
    "Status",
    "Variance"
  ]]);

  sh.getRange("A22:H22")
    .setBackground("#D9EAD3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("A23:H40")
    .setBorder(true, true, true, true, true, true);

  //==========================================================
  // LIVE QUEUE SHARE
  //==========================================================

  sh.getRange("J21:N21")
    .merge()
    .setValue("LIVE QUEUE SHARE")
    .setBackground("#0F4C81")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("J22:N22").setValues([[
    "Queue",
    "Assigned",
    "On Queue",
    "On Break",
    "Coverage"
  ]]);

  sh.getRange("J22:N22")
    .setBackground("#D9EAD3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("J23:N30")
    .setBorder(true, true, true, true, true, true);

  //==========================================================
  // LATE RETURNS
  //==========================================================

  sh.getRange("A43:H43")
    .merge()
    .setValue("LATE RETURNS")
    .setBackground("#C62828")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("A44:H44").setValues([[
    "Agent",
    "Queue",
    "Expected Back",
    "Actual Back",
    "Variance",
    "Supervisor",
    "Override",
    "Remarks"
  ]]);

  sh.getRange("A44:H44")
    .setBackground("#F4CCCC")
    .setFontWeight("bold");

  sh.getRange("A45:H60")
    .setBorder(true, true, true, true, true, true);

  //==========================================================
  // SUPERVISOR OVERRIDES
  //==========================================================

  sh.getRange("J43:P43")
    .merge()
    .setValue("SUPERVISOR OVERRIDES")
    .setBackground("#F57C00")
    .setFontColor("white")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sh.getRange("J44:P44").setValues([[
    "Agent",
    "Queue",
    "Time",
    "Supervisor",
    "Reason",
    "Status",
    "Remarks"
  ]]);

  sh.getRange("J44:P44")
    .setBackground("#FCE5CD")
    .setFontWeight("bold");

  sh.getRange("J45:P60")
    .setBorder(true, true, true, true, true, true);

  //==========================================================
  // FORMATTING
  //==========================================================

  sh.setFrozenRows(3);

  sh.autoResizeColumns(1,16);

  sh.setColumnWidths(1,16,120);

}