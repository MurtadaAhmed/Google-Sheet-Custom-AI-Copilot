/* =========================
   WORKBOOK / CONTEXT HELPERS
========================= */

function getActualDataBounds(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow === 0 || lastColumn === 0) {
    return { lastRow: 0, lastColumn: 0, isEmpty: true };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();

  let realLastRow = 0;
  let realLastColumn = 0;

  for (let r = 0; r < values.length; r++) {
    let rowHasData = false;
    for (let c = 0; c < values[r].length; c++) {
      if (values[r][c] !== "") {
        rowHasData = true;
        realLastColumn = Math.max(realLastColumn, c + 1);
      }
    }
    if (rowHasData) realLastRow = r + 1;
  }

  if (realLastRow === 0 || realLastColumn === 0) {
    return { lastRow: 0, lastColumn: 0, isEmpty: true };
  }

  return { lastRow: realLastRow, lastColumn: realLastColumn, isEmpty: false };
}

function getSheetMetadata(sheet, bounds, allNamedRanges) {
  const maxHeaderCols = Math.min(bounds.lastColumn || 0, 20);
  let headerRow = [];
  if (!bounds.isEmpty && bounds.lastRow >= 1 && maxHeaderCols > 0) {
    headerRow = sheet.getRange(1, 1, 1, maxHeaderCols).getDisplayValues()[0];
  }

  return {
    name: sheet.getName(),
    isEmpty: bounds.isEmpty,
    lastRow: bounds.lastRow,
    lastColumn: bounds.lastColumn,
    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns(),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns(),
    chartCount: safeGetCharts(sheet).length,
    headerPreview: headerRow,
    namedRanges: (allNamedRanges || [])
      .filter(nr => nr.getRange().getSheet().getName() === sheet.getName())
      .slice(0, 20)
      .map(nr => ({
        name: nr.getName(),
        range: nr.getRange().getA1Notation()
      }))
  };
}

function buildSheetContextString(sheet, allNamedRanges) {
  try {
    const bounds = getActualDataBounds(sheet);
    const meta = getSheetMetadata(sheet, bounds, allNamedRanges);

    let out = [];
    out.push(`=== SHEET: "${sheet.getName()}" ===`);
    out.push(`META: rows=${meta.lastRow}, cols=${meta.lastColumn}, frozenRows=${meta.frozenRows}, frozenColumns=${meta.frozenColumns}, charts=${meta.chartCount}, empty=${meta.isEmpty}`);
    if (meta.headerPreview && meta.headerPreview.length > 0) {
      out.push(`HEADER_PREVIEW: ${meta.headerPreview.map(v => v === "" ? "[EMPTY]" : v).join(" | ")}`);
    }
    if (meta.namedRanges.length > 0) {
      out.push(`NAMED_RANGES: ${meta.namedRanges.map(n => `${n.name}=${n.range}`).join(", ")}`);
    }

    if (bounds.isEmpty) {
      out.push("[Empty]");
      return out.join("\n");
    }

    const maxCols = Math.min(bounds.lastColumn, 20);
    const headRows = Math.min(bounds.lastRow, 40);
    const tailRows = bounds.lastRow > 60 ? Math.min(10, bounds.lastRow - headRows) : Math.max(0, bounds.lastRow - headRows);
    const showTail = bounds.lastRow > headRows + tailRows;

    const headValues = sheet.getRange(1, 1, headRows, maxCols).getValues();
    const headFormulas = sheet.getRange(1, 1, headRows, maxCols).getFormulas();

    for (let r = 0; r < headRows; r++) {
      let row = [];
      for (let c = 0; c < maxCols; c++) {
        let cell = headFormulas[r][c] ? headFormulas[r][c] : headValues[r][c];
        row.push(cell === "" ? "[EMPTY]" : cell);
      }
      out.push(row.join(" | "));
    }

    if (showTail) {
      out.push(`[... ${bounds.lastRow - headRows - tailRows} middle rows hidden to save memory ...]`);
      const startTailRow = bounds.lastRow - tailRows + 1;
      const tailValues = sheet.getRange(startTailRow, 1, tailRows, maxCols).getValues();
      const tailFormulas = sheet.getRange(startTailRow, 1, tailRows, maxCols).getFormulas();

      for (let r = 0; r < tailRows; r++) {
        let row = [];
        for (let c = 0; c < maxCols; c++) {
          let cell = tailFormulas[r][c] ? tailFormulas[r][c] : tailValues[r][c];
          row.push(cell === "" ? "[EMPTY]" : cell);
        }
        out.push(row.join(" | "));
      }
    }

    return out.join("\n");
  } catch (e) {
    return `=== SHEET: "${sheet.getName()}" ===\n[Context unavailable: ${e.message}]`;
  }
}

function buildContextForSheets(sheetNames) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const realSheets = spreadsheet.getSheets();
  const allNamedRanges = spreadsheet.getNamedRanges();
  const selected = realSheets.filter(s => sheetNames.includes(s.getName()));
  return selected.map(sheet => buildSheetContextString(sheet, allNamedRanges)).join("\n\n");
}

function safeGetCharts(sheet) {
  try {
    return sheet.getCharts();
  } catch (e) {
    return [];
  }
}
