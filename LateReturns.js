function updateLateReturns() {

  const ss = SpreadsheetApp.getActive();

  const dash = ss.getSheetByName("Dashboard");
  const ops = ss.getSheetByName("Daily Operations");

  dash.getRange("A45:H60").clearContent();

  const data = ops.getRange("A6:O500").getValues();

  const output = [];

  data.forEach(r => {

    if (r[14] !== "Break Overdue") return;

    output.push([
      r[1],   // Agent
      r[5],   // Queue
      r[8],   // Scheduled Back
      "",
      "",
      r[3],   // Supervisor
      "",
      "Overdue"
    ]);

  });

  if (output.length) {

    dash.getRange(45,1,output.length,8)
        .setValues(output);

  }

}