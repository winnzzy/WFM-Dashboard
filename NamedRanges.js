/**
 * =====================================
 * NAMED RANGES
 * =====================================
 */

function createNamedRanges() {

  const ss = SpreadsheetApp.getActive();

  const sh = ss.getSheetByName("Daily Operations");

  ss.setNamedRange("AgentColumn", sh.getRange("B6:B155"));
  ss.setNamedRange("CenterColumn", sh.getRange("C6:C155"));
  ss.setNamedRange("SupervisorColumn", sh.getRange("D6:D155"));
  ss.setNamedRange("ShiftColumn", sh.getRange("E6:E155"));
  ss.setNamedRange("QueueColumn", sh.getRange("F6:F155"));
  ss.setNamedRange("LoginColumn", sh.getRange("G6:G155"));
  ss.setNamedRange("BreakStartColumn", sh.getRange("H6:H155"));
  ss.setNamedRange("ExpectedBackColumn", sh.getRange("I6:I155"));
  ss.setNamedRange("ActualOutColumn", sh.getRange("J6:J155"));
  ss.setNamedRange("ActualBackColumn", sh.getRange("K6:K155"));
  ss.setNamedRange("BreakUsedColumn", sh.getRange("L6:L155"));
  ss.setNamedRange("VarianceColumn", sh.getRange("M6:M155"));
  ss.setNamedRange("StatusColumn", sh.getRange("N6:N155"));

}