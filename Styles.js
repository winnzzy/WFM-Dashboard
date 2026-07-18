/**
 * FORMATTING
 */

function applyFormatting() {

  const sh = SpreadsheetApp.getActive()
    .getSheetByName("Daily Operations");

  // Borders
  sh.getRange("A5:O155").setBorder(
    true,
    true,
    true,
    true,
    true,
    true,
    "#DADCE0",
    SpreadsheetApp.BorderStyle.SOLID
  );

  // Alternate row colours
  for (let row = 6; row <= 155; row++) {

    if (row % 2 === 0) {

      sh.getRange(row, 1, 1, 15)
        .setBackground("#F8F9FA");

    } else {

      sh.getRange(row, 1, 1, 15)
        .setBackground("#FFFFFF");

    }

  }

  // Font
  sh.getRange("A5:O155")
    .setFontFamily("Arial")
    .setFontSize(10);

  // Alignments
  sh.getRange("A:O")
    .setVerticalAlignment("middle");

}