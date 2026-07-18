/**
 * ==========================================================
 * BREAK AUTOMATION V2
 * ==========================================================
 */

function startBreak(){

  const sh = SpreadsheetApp.getActive()
    .getSheetByName("Daily Operations");

  const row = sh.getActiveCell().getRow();

  if(row < 6){
    SpreadsheetApp.getUi().alert("Select an agent.");
    return;
  }

  const status = sh.getRange(row,COL.STATUS).getValue();

  if(status=="On Break"){
    SpreadsheetApp.getUi().alert("Agent is already on break.");
    return;
  }
  const shift = sh.getRange(row,COL.SHIFT).getValue();

if (!isWithinBreakWindow(shift)) {

  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Supervisor Override",
    "This agent is outside the approved break window.\n\nContinue anyway?",
    ui.ButtonSet.YES_NO
  );

  if (response != ui.Button.YES) {
    return;
  }

  sh.getRange(row, COL.OVERRIDE)
    .setValue("YES");

  sh.getRange(row, COL.OVERRIDE_TIME)
    .setValue(new Date())
    .setNumberFormat("h:mm AM/PM");

  sh.getRange(row, COL.REMARKS)
    .setValue("Outside approved break window");

}

  const queue = sh.getRange(row,COL.QUEUE).getValue();

  const data = sh.getRange(6,1,300,16).getValues();

  let agentsOnBreak = 0;
  let sameQueue = false;

  data.forEach(r=>{

    if(r[COL.STATUS-1]!="On Break") return;

    agentsOnBreak++;

    if(r[COL.QUEUE-1]==queue){
      sameQueue=true;
    }

  });

  if(agentsOnBreak>=2){

    SpreadsheetApp.getUi().alert(
      "Maximum of 2 agents are already on break."
    );

    return;

  }

  if(sameQueue){

    SpreadsheetApp.getUi().alert(
      "Another agent on '"+queue+"' is already on break."
    );

    return;

  }

  const now = new Date();

  sh.getRange(row,COL.ACTUAL_OUT)
    .setValue(now)
    .setNumberFormat("h:mm AM/PM");

  sh.getRange(row,COL.STATUS)
    .setValue("On Break");

  refreshDashboard();


}

function returnFromBreak(){

  const sh=SpreadsheetApp.getActive()
  .getSheetByName("Daily Operations");

  const row=sh.getActiveCell().getRow();

  if(row<6){

    SpreadsheetApp.getUi()
    .alert("Select an agent.");

    return;

  }

  if(sh.getRange(row,COL.STATUS).getValue()!="On Break"){

    SpreadsheetApp.getUi()
    .alert("Agent is not on break.");

    return;

  }

  const actualBack=new Date();

  sh.getRange(row,COL.ACTUAL_BACK)
    .setValue(actualBack)
    .setNumberFormat("h:mm AM/PM");

  const actualOut=
    sh.getRange(row,COL.ACTUAL_OUT).getValue();

  const scheduledBack=
    sh.getRange(row,COL.SCHEDULED_BACK).getValue();

  const breakMinutes=
    Math.round(
      (actualBack-actualOut)/60000
    );

  sh.getRange(row,COL.BREAK_USED)
    .setValue(breakMinutes);

  const variance=
    Math.round(
      (actualBack-scheduledBack)/60000
    );

  if(variance<0){

    sh.getRange(row,COL.VARIANCE)
      .setValue(
        "Early by "+Math.abs(variance)+" mins"
      );

  }else if(variance==0){

    sh.getRange(row,COL.VARIANCE)
      .setValue("On Time");

  }else{

    sh.getRange(row,COL.VARIANCE)
      .setValue(
        "Late by "+variance+" mins"
      );

  }

  sh.getRange(row,COL.STATUS)
    .setValue("On Queue");

  refreshDashboard();


}