/**
 * ================================================================
 * Script Name : Sheet_Index_Simple_Hide_Show
 * Description : Rebuilds the "Sheet Index" tab (Sr No, Sheet Name,
 *               Hyperlink, Visible) and hides/shows every tab based
 *               on the Visible column.
 *
 *               PRESERVES on every rebuild, matched by sheet name:
 *                 - Extra columns to the right of D
 *                 - Notes/footer rows below the table
 *                 - Cell formatting (fill, font, borders, number
 *                   format — via native copyFrom, not cell-by-cell)
 *                 - Column widths and row heights
 *
 *               PERFORMANCE NOTE: formatting is copied using
 *               Range.copyFrom() once per row, instead of reading/
 *               writing fill+font+bold+italic per cell. That's the
 *               difference between ~100 API calls and ~3000+ for a
 *               50-row sheet — this is what makes it fast.
 *
 * Platform    : Excel Office Scripts (Excel for Web / M365)
 * ================================================================
 */

function padRow(arr: string[], width: number): string[] {
  const out = arr.slice(0, width);
  while (out.length < width) out.push("");
  return out;
}

function main(workbook: ExcelScript.Workbook) {
  const INDEX_SHEET_NAME = "Sheet Index";
  const TEMP_PREFIX = "Sheet Index (rebuilding";
  // Unique per run — guarantees this run never needs to delete-then-immediately-
  // reuse a name, which was the source of the InvalidArgument naming collision.
  const TEMP_SHEET_NAME = `_IDX_${Date.now() % 100000}`;

  try {
    // ── STEP 0: Clean up ANY stray "rebuilding" sheets left by previous failed runs ──
    const strayTemps = workbook.getWorksheets()
      .map(s => s.getName())
      .filter(n => n.indexOf(TEMP_PREFIX) === 0);

    if (strayTemps.length > 0) {
      if (!workbook.getWorksheet(INDEX_SHEET_NAME)) {
        // No real index exists — recover data by renaming the stray one back.
        workbook.getWorksheet(strayTemps[0]).setName(INDEX_SHEET_NAME);
        strayTemps.shift();
      }
      // Anything left over at this point is genuinely orphaned — remove it.
      for (const strayName of strayTemps) {
        workbook.getWorksheet(strayName).delete();
      }
    }

    const oldSheet = workbook.getWorksheet(INDEX_SHEET_NAME);
    const isRebuild = oldSheet !== undefined;

    const existingVisibleFlag = new Map<string, string>();
    const extraColsBySheetName = new Map<string, string[]>();
    const oldRowIndexBySheetName = new Map<string, number>();
    const rowHeightBySheetName = new Map<string, number>();
    let extraHeaderRow: string[] = [];
    let headerRowHeight = -1;
    let footerRows: string[][] = [];
    let footerOldRowStart = -1;
    let columnWidths: number[] = [];
    let oldColCount = 0;

    // ── STEP 1: Read old sheet's content + light metadata in bulk (cheap calls) ──
    if (oldSheet) {
      const usedRange = oldSheet.getUsedRange();
      if (usedRange) {
        const oldGrid = usedRange.getFormulas(); // ONE call returns the whole grid
        const rowCount = oldGrid.length;
        oldColCount = rowCount > 0 ? oldGrid[0].length : 0;

        for (let c = 0; c < oldColCount; c++) {
          columnWidths.push(oldSheet.getRangeByIndexes(0, c, 1, 1).getFormat().getColumnWidth());
        }
        if (oldColCount > 4) extraHeaderRow = oldGrid[0].slice(4);
        if (rowCount > 0) {
          headerRowHeight = oldSheet.getRangeByIndexes(0, 0, 1, 1).getFormat().getRowHeight();
        }

        let i = 1;
        for (; i < rowCount; i++) {
          const name = oldGrid[i][1];
          const link = oldGrid[i][2];
          if (!name && !link) break; // footer boundary (both blank)

          existingVisibleFlag.set(name, oldGrid[i][3] || "Show");
          if (oldColCount > 4) extraColsBySheetName.set(name, oldGrid[i].slice(4));
          oldRowIndexBySheetName.set(name, i);
          rowHeightBySheetName.set(name, oldSheet.getRangeByIndexes(i, 0, 1, 1).getFormat().getRowHeight());
        }

        footerOldRowStart = i;
        for (; i < rowCount; i++) {
          footerRows.push(oldGrid[i]);
        }
      }
    }

    // ── STEP 2: Free up the real name by renaming the OLD sheet out of the way ──
    if (oldSheet) {
      oldSheet.setName(TEMP_SHEET_NAME);
    }
    const indexSheet = workbook.addWorksheet(INDEX_SHEET_NAME);
    indexSheet.setVisibility(ExcelScript.SheetVisibility.visible);
    indexSheet.setPosition(0);

    // ── STEP 3: Collect current sheet list ──
    const allSheets = workbook.getWorksheets();
    const sheetData: string[][] = [];
    let srNo = 1;

    for (const sheet of allSheets) {
      const sheetName = sheet.getName();
      if (sheetName === INDEX_SHEET_NAME || sheetName === TEMP_SHEET_NAME) continue;

      const escapedSheetName = sheetName.replace(/'/g, "''");
      const visibleFlag = existingVisibleFlag.get(sheetName) || "Show";

      sheetData.push([
        srNo.toString(),
        sheetName,
        `=HYPERLINK("#'${escapedSheetName}'!A1","Open")`,
        visibleFlag
      ]);
      srNo++;
    }

    // ── STEP 4: Write header + core data (A:D) ──
    indexSheet.getRange("A1:D1").setValues([["Sr No", "Sheet Name", "Link", "Visible"]]);

    let lastRowIdx = 1;
    if (sheetData.length > 0) {
      lastRowIdx = sheetData.length + 1;
      indexSheet.getRange(`A2:D${lastRowIdx}`).setFormulas(sheetData);

      const table = indexSheet.addTable(indexSheet.getRange(`A1:D${lastRowIdx}`), true);

      const headerRange = indexSheet.getRange("A1:D1");
      headerRange.getFormat().getFill().setColor("#2E4057");
      headerRange.getFormat().getFont().setColor("#FFFFFF");
      headerRange.getFormat().getFont().setBold(true);

      try {
        indexSheet.getRange(`D2:D${lastRowIdx}`).getDataValidation().setRule({
          list: { inCellDropDown: true, source: "Show,Hide" }
        });
      } catch {
        try {
          indexSheet.getRange(`D2:D${lastRowIdx}`).getDataValidation().setRule({
            list: { inCellDropDown: true, source: "Show;Hide" }
          });
        } catch { }
      }
      indexSheet.getRange(`A2:A${lastRowIdx}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
      indexSheet.getRange(`C2:D${lastRowIdx}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

      const linkRange = indexSheet.getRange(`C2:C${lastRowIdx}`);
      linkRange.getFormat().getFont().setUnderline(ExcelScript.RangeUnderlineStyle.single);
      linkRange.getFormat().getFont().setColor("#0563C1");
    }

    // ── STEP 5: Write extra columns (right of D), matched by sheet name ──
    const maxExtraCols = Math.max(
      extraHeaderRow.length,
      ...Array.from(extraColsBySheetName.values()).map(r => r.length),
      0
    );
    if (maxExtraCols > 0) {
      indexSheet.getRangeByIndexes(0, 4, 1, maxExtraCols).setFormulas([padRow(extraHeaderRow, maxExtraCols)]);
      sheetData.forEach((row, idx) => {
        const extra = extraColsBySheetName.get(row[1]);
        if (extra && extra.length > 0) {
          indexSheet.getRangeByIndexes(idx + 1, 4, 1, maxExtraCols).setFormulas([padRow(extra, maxExtraCols)]);
        }
      });
    }

    // ── STEP 6: Write footer/notes rows below the table ──
    if (footerRows.length > 0) {
      footerRows.forEach((rowArr, idx) => {
        indexSheet.getRangeByIndexes(lastRowIdx + idx, 0, 1, rowArr.length).setFormulas([rowArr]);
      });
    }

    // ── STEP 7: FAST formatting restore — one copyFrom call per row, not per cell ──
    if (isRebuild && oldSheet && oldColCount > 0) {
      indexSheet.getRangeByIndexes(0, 0, 1, oldColCount)
        .copyFrom(oldSheet.getRangeByIndexes(0, 0, 1, oldColCount), ExcelScript.RangeCopyType.formats);

      sheetData.forEach((row, idx) => {
        const oldIdx = oldRowIndexBySheetName.get(row[1]);
        if (oldIdx !== undefined) {
          indexSheet.getRangeByIndexes(idx + 1, 0, 1, oldColCount)
            .copyFrom(oldSheet.getRangeByIndexes(oldIdx, 0, 1, oldColCount), ExcelScript.RangeCopyType.formats);
        }
      });

      footerRows.forEach((rowArr, idx) => {
        const oldIdx = footerOldRowStart + idx;
        indexSheet.getRangeByIndexes(lastRowIdx + idx, 0, 1, oldColCount)
          .copyFrom(oldSheet.getRangeByIndexes(oldIdx, 0, 1, oldColCount), ExcelScript.RangeCopyType.formats);
      });
    }

    // ── STEP 8: Column widths / row heights (already cheap — single property per column/row) ──
    indexSheet.getUsedRange().getFormat().autofitColumns();
    if (isRebuild) {
      const finalColCount = indexSheet.getUsedRange().getFormulas()[0].length;
      for (let c = 0; c < finalColCount; c++) {
        if (columnWidths[c] !== undefined && columnWidths[c] >= 0) {
          indexSheet.getRangeByIndexes(0, c, 1, 1).getFormat().setColumnWidth(columnWidths[c]);
        }
      }
      if (headerRowHeight >= 0) {
        indexSheet.getRangeByIndexes(0, 0, 1, 1).getFormat().setRowHeight(headerRowHeight);
      }
      sheetData.forEach((row, idx) => {
        const h = rowHeightBySheetName.get(row[1]);
        if (h !== undefined && h >= 0) {
          indexSheet.getRangeByIndexes(idx + 1, 0, 1, 1).getFormat().setRowHeight(h);
        }
      });
    }

    indexSheet.getFreezePanes().freezeRows(1);

    // ── STEP 9: Remove the old (renamed) sheet now that the new one is fully built ──
    if (isRebuild && oldSheet) {
      oldSheet.delete();
    }

    // ── STEP 10: Apply Show/Hide ──
    let shown = 0;
    let hidden = 0;
    for (const row of sheetData) {
      const targetSheet = workbook.getWorksheet(row[1]);
      if (!targetSheet) continue;
      const shouldShow = row[3] !== "Hide";
      targetSheet.setVisibility(
        shouldShow ? ExcelScript.SheetVisibility.visible : ExcelScript.SheetVisibility.hidden
      );
      shouldShow ? shown++ : hidden++;
    }

    console.log(`✅ Sheet Index rebuilt (${sheetData.length} tabs). Shown: ${shown}, Hidden: ${hidden}.`);

  } catch (error) {
    console.log(`❌ Script failed: ${String(error)}`);
    throw error;
  }
}
