/**
 * ==========================================================
 * FORMULAS (Optimized)
 * ==========================================================
 */

function applyFormulas(){

  const sh =
    SpreadsheetApp.getActive()
    .getSheetByName("Daily Operations");

  // Expected Back
  sh.getRange("I6").setFormula(`
=ARRAYFORMULA(
IF(H6:H="","",
H6:H+TIME(1,0,0))
)`);

  // Break Used
  sh.getRange("L6").setFormula(`
=ARRAYFORMULA(
IF((J6:J="")+(K6:K=""),
"",
ROUND((K6:K-J6:J)*1440,0))
)`);

  // Variance
  sh.getRange("M6").setFormula(`
=ARRAYFORMULA(
IF((I6:I="")+(K6:K=""),
"",
IF(K6:K=I6:I,
"On Time",
IF(K6:K>I6:I,
"Late "&ROUND((K6:K-I6:I)*1440,0)&" mins",
"Early "&ROUND((I6:I-K6:K)*1440,0)&" mins"))))
`);

  // Time Formatting

  sh.getRange("G:I")
    .setNumberFormat("h:mm AM/PM");

  sh.getRange("J:K")
    .setNumberFormat("h:mm AM/PM");

}