/**
 * ==========================================================
 * FORMATTING
 * ==========================================================
 * Applies visual formatting to the Daily Operations sheet.
 */

/**
 * Applies borders, alternating row colors, fonts, and
 * alignment to the Daily Operations data area.
 *
 * Performance: uses batch operations instead of per-row loops.
 */
function applyFormatting() {

  var sh = SpreadsheetApp.getActive()
    .getSheetByName(SHEETS.DAILY_OPS);

  var dataRows = MAX_OPS_ROWS - OPS_DATA_START_ROW + 1; // 495 rows

  // --- Borders ---
  sh.getRange(
    OPS_DATA_START_ROW - 1,
    1,
    dataRows + 1,
    OPS_COL_COUNT
  ).setBorder(
    true, true, true, true, true, true,
    "#DADCE0",
    SpreadsheetApp.BorderStyle.SOLID
  );

  // --- Alternating row colors (batch write) ---
  var evenRows = [];
  var oddRows = [];

  for (var c = 0; c < OPS_COL_COUNT; c++) {
    evenRows.push(COLORS.ALT_ROW);
    oddRows.push(COLORS.BODY_BG);
  }

  // Even rows (6, 8, 10...)
  for (var row = OPS_DATA_START_ROW; row <= MAX_OPS_ROWS; row += 2) {
    sh.getRange(row, 1, 1, OPS_COL_COUNT)
      .setBackgrounds([evenRows]);
  }

  // Odd rows (7, 9, 11...)
  for (var row2 = OPS_DATA_START_ROW + 1; row2 <= MAX_OPS_ROWS; row2 += 2) {
    sh.getRange(row2, 1, 1, OPS_COL_COUNT)
      .setBackgrounds([oddRows]);
  }

  // --- Font ---
  sh.getRange(
    OPS_DATA_START_ROW - 1,
    1,
    dataRows + 1,
    OPS_COL_COUNT
  ).setFontFamily("Arial")
    .setFontSize(10);

  // --- Alignment ---
  sh.getRange(1, 1, sh.getMaxRows(), OPS_COL_COUNT)
    .setVerticalAlignment("middle");
}