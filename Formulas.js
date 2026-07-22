/**
 * ==========================================================
 * FORMULAS
 * ==========================================================
 * Applies ARRAYFORMULA formulas to calculated columns
 * in the Daily Operations sheet.
 */

/**
 * Applies formulas for Expected Back, Break Used, and Variance.
 * Time formatting applied to relevant columns.
 */
function applyFormulas() {

  var sh = SpreadsheetApp.getActive()
    .getSheetByName(SHEETS.DAILY_OPS);

  // Scheduled Back (Column I) is populated by the scheduler as full DateTime.
  // If a legacy formula exists, replace it with static values.
  var scheduledBackRange = sh.getRange("I" + OPS_DATA_START_ROW + ":I" + MAX_OPS_ROWS);
  scheduledBackRange.setValues(scheduledBackRange.getValues());

  // Break Used (Column M) = Actual Back - Actual Out, in minutes.
  // Guard against blanks, non-datetime values, zero-date/time-only values,
  // and never allow negative duration.
  sh.getRange("M6").setFormula(
    '=ARRAYFORMULA(IF((COUNT(K6:K)+COUNT(L6:L))=0,"",IF((ISBLANK(K6:K))+(ISBLANK(L6:L)),"",IF(IFERROR((NOT(ISNUMBER(K6:K)))+(NOT(ISNUMBER(L6:L))),TRUE),"",IF((K6:K<=1)+(L6:L<=1),"",IF(ROUND((L6:L-K6:K)*1440,0)<0,0,ROUND((L6:L-K6:K)*1440,0)))))))'
  );

  // Variance (Column N) = Actual Back (L) vs Expected Back (I)
  // with strict validity checks and an upper bound to suppress
  // unrealistic values from mixed date/time types.
  sh.getRange("N6").setFormula(
    '=ARRAYFORMULA(IF((COUNT(I6:I)+COUNT(L6:L))=0,"",IF((ISBLANK(I6:I))+(ISBLANK(L6:L)),"",IF(IFERROR((NOT(ISNUMBER(I6:I)))+(NOT(ISNUMBER(L6:L))),TRUE),"",IF((I6:I<=1)+(L6:L<=1),"",IF(ABS(ROUND((L6:L-I6:I)*1440,0))>720,"",IF(ROUND((L6:L-I6:I)*1440,0)=0,"On Time",IF(ROUND((L6:L-I6:I)*1440,0)>0,"Late by "&ROUND((L6:L-I6:I)*1440,0)&" mins","Early by "&ABS(ROUND((L6:L-I6:I)*1440,0))&" mins"))))))))'
  );

  // Display-only formatting: keep full DateTime values internally,
  // but show Scheduled Out/Back (H:I) as time-only to users.
  sh.getRange("H:I").setNumberFormat("h:mm AM/PM");
}