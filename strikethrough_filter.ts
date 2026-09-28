function main(workbook: ExcelScript.Workbook) {
  const sheet = workbook.getActiveWorksheet();

  // ===== Grab whatever column/range you've selected before running =====
  const selectedRange = workbook.getSelectedRange();
  const targetColIndex = selectedRange.getColumnIndex(); // uses the first column of your selection
  // =======================================================================

  const usedRange = sheet.getUsedRange();
  const rowCount = usedRange.getRowCount();
  const colCount = usedRange.getColumnCount();
  const startRow = usedRange.getRowIndex();
  const startCol = usedRange.getColumnIndex();

  // Auto-detect header row: first row in usedRange where column A (Sr No) has a number
  let headerOffset = 0;
  for (let r = 0; r < rowCount; r++) {
    const val = usedRange.getCell(r, 0).getValue();
    if (typeof val === "string" && val.toLowerCase().includes("sr")) {
      headerOffset = r;
      break;
    }
  }

  const helperCol = startCol + colCount;
  sheet.getCell(startRow + headerOffset, helperCol).setValue("Strikethrough?");

  for (let r = headerOffset + 1; r < rowCount; r++) {
    const cell = sheet.getCell(startRow + r, targetColIndex);
    const font = cell.getFormat().getFont();
    const hasStrikethrough = font.getStrikethrough();
    sheet.getCell(startRow + r, helperCol).setValue(hasStrikethrough ? "Yes" : "No");
  }

  const fullRange = sheet.getRangeByIndexes(
    startRow + headerOffset,
    startCol,
    rowCount - headerOffset,
    colCount + 1
  );

  sheet.getAutoFilter().remove();
  sheet.getAutoFilter().apply(fullRange, undefined);
}
