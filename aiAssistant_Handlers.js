/* =========================
   ACTION HANDLERS
========================= */

function handleSheetStructureAction(spreadsheet, edit, state) {
  if (edit.action === "addSheet" && edit.sheetName) {
    let existing = spreadsheet.getSheetByName(edit.sheetName);
    if (!existing) {
      existing = spreadsheet.insertSheet(edit.sheetName);
      state.newlyAddedSheets.push(edit.sheetName);
      state.scopeChanges.push(`Added sheet "${edit.sheetName}" to scope.`);
    }

    if (!state.activeWritableSheets.includes(edit.sheetName)) {
      state.activeWritableSheets.push(edit.sheetName);
    }
    return true;
  }

  if (edit.action === "deleteSheet" && edit.sheetName) {
    const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
    if (targetSheet) {
      spreadsheet.deleteSheet(targetSheet);
      state.deletedSheets.push(edit.sheetName);
      state.activeWritableSheets = state.activeWritableSheets.filter(n => n !== edit.sheetName);
      state.scopeChanges.push(`Deleted sheet "${edit.sheetName}" from workbook/scope.`);
    }
    return true;
  }

  if (edit.action === "renameSheet" && edit.sheetName && edit.newSheetName) {
    const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
    if (targetSheet && !spreadsheet.getSheetByName(edit.newSheetName)) {
      targetSheet.setName(edit.newSheetName);

      state.activeWritableSheets = state.activeWritableSheets.map(n => n === edit.sheetName ? edit.newSheetName : n);
      state.allowedReadSheets = state.allowedReadSheets.map(n => n === edit.sheetName ? edit.newSheetName : n);

      if (state.newlyAddedSheets.includes(edit.sheetName)) {
        state.newlyAddedSheets = state.newlyAddedSheets.map(n => n === edit.sheetName ? edit.newSheetName : n);
      }

      state.scopeChanges.push(`Renamed sheet "${edit.sheetName}" to "${edit.newSheetName}".`);
    }
    return true;
  }

  if (edit.action === "duplicateSheet" && edit.sheetName && edit.newSheetName) {
    const sourceSheet = spreadsheet.getSheetByName(edit.sheetName);
    if (sourceSheet && !spreadsheet.getSheetByName(edit.newSheetName)) {
      const newSheet = sourceSheet.copyTo(spreadsheet).setName(edit.newSheetName);
      spreadsheet.setActiveSheet(newSheet);
      spreadsheet.moveActiveSheet(spreadsheet.getNumSheets());

      if (!state.activeWritableSheets.includes(edit.newSheetName)) {
        state.activeWritableSheets.push(edit.newSheetName);
      }
      state.newlyAddedSheets.push(edit.newSheetName);
      state.scopeChanges.push(`Duplicated sheet "${edit.sheetName}" as "${edit.newSheetName}" and added it to scope.`);
    }
    return true;
  }

  if (edit.action === "moveSheet" && edit.sheetName && edit.newIndex !== undefined) {
    const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
    if (targetSheet) {
      const bounded = clamp(parseInt(edit.newIndex, 10), 1, spreadsheet.getNumSheets());
      spreadsheet.setActiveSheet(targetSheet);
      spreadsheet.moveActiveSheet(bounded);
    }
    return true;
  }

  return false;
}

function handleFreezeClearAppendProtectActions(spreadsheet, edit, state) {
  if (!edit.sheetName || !state.activeWritableSheets.includes(edit.sheetName)) return false;

  const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
  if (!targetSheet) return false;

  if (edit.action === "freeze") {
    if (edit.frozenRows !== undefined) targetSheet.setFrozenRows(edit.frozenRows);
    if (edit.frozenColumns !== undefined) targetSheet.setFrozenColumns(edit.frozenColumns);
    return true;
  }

  if (edit.action === "clear" && edit.range) {
    targetSheet.getRange(edit.range).clear();
    return true;
  }

  if (edit.action === "clearContent" && edit.range) {
    targetSheet.getRange(edit.range).clearContent();
    return true;
  }

  if (edit.action === "clearConditionalFormat" && !edit.range) {
    targetSheet.setConditionalFormatRules([]);
    return true;
  }

  if (edit.action === "appendRow" && Array.isArray(edit.values)) {
    targetSheet.appendRow(edit.values);
    return true;
  }

  if (edit.action === "appendRows" && Array.isArray(edit.values2D) && edit.values2D.length > 0) {
    const width = edit.values2D[0].length;
    const startRow = Math.max(targetSheet.getLastRow(), 0) + 1;
    const appendRange = targetSheet.getRange(startRow, 1, edit.values2D.length, width);
    appendRange.setValues(edit.values2D);
    return true;
  }

  if (edit.action === "protectRange" && edit.range) {
    const targetRange = targetSheet.getRange(edit.range);
    const protection = targetRange.protect();

    if (edit.description) protection.setDescription(edit.description);

    if (edit.warningOnly === true) {
      protection.setWarningOnly(true);
      return true;
    }

    const preserveExistingEditors = edit.preserveExistingEditors !== false;
    const me = Session.getEffectiveUser();

    if (!preserveExistingEditors) {
      const currentEditors = protection.getEditors();
      if (currentEditors && currentEditors.length > 0) protection.removeEditors(currentEditors);
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
    }

    try { protection.addEditor(me); } catch (e) {}

    if (Array.isArray(edit.addEditors) && edit.addEditors.length > 0) {
      edit.addEditors.forEach(email => {
        try { protection.addEditor(email); } catch (e) {}
      });
    }

    if (edit.domainEdit === false && protection.canDomainEdit()) protection.setDomainEdit(false);

    return true;
  }

  return false;
}

function handleMergeResizeNamedRangeActions(spreadsheet, edit, state) {
  if (!edit.sheetName || !state.activeWritableSheets.includes(edit.sheetName)) return false;
  const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
  if (!targetSheet) return false;

  if (edit.action === "merge" && edit.range) {
    const r = targetSheet.getRange(edit.range);
    cleanupOverlappingMergedRanges(targetSheet, r);

    if (edit.mergeType === "across") r.mergeAcross();
    else if (edit.mergeType === "vertically") r.mergeVertically();
    else r.merge();

    return true;
  }

  if (edit.action === "unmerge" && edit.range) {
    const r = targetSheet.getRange(edit.range);
    cleanupOverlappingMergedRanges(targetSheet, r);
    r.breakApart();
    return true;
  }

  if (edit.action === "resizeColumn" && edit.column && edit.width) {
    targetSheet.setColumnWidth(edit.column, edit.width);
    return true;
  }

  if (edit.action === "resizeColumns" && edit.startColumn && edit.numColumns && edit.width) {
    for (let c = edit.startColumn; c < edit.startColumn + edit.numColumns; c++) {
      targetSheet.setColumnWidth(c, edit.width);
    }
    return true;
  }

  if (edit.action === "resizeRow" && edit.row && edit.height) {
    targetSheet.setRowHeight(edit.row, edit.height);
    return true;
  }

  if (edit.action === "resizeRows" && edit.startRow && edit.numRows && edit.height) {
    for (let r = edit.startRow; r < edit.startRow + edit.numRows; r++) {
      targetSheet.setRowHeight(r, edit.height);
    }
    return true;
  }

  if (edit.action === "autoResizeColumns") {
    if (edit.startColumn && edit.numColumns) {
      targetSheet.autoResizeColumns(edit.startColumn, edit.numColumns);
      return true;
    }
    if (edit.range) {
      const r = targetSheet.getRange(edit.range);
      targetSheet.autoResizeColumns(r.getColumn(), r.getNumColumns());
      return true;
    }
  }

  if (edit.action === "autoResizeRows") {
    if (edit.startRow && edit.numRows) {
      targetSheet.autoResizeRows(edit.startRow, edit.numRows);
      return true;
    }
    if (edit.range) {
      const r = targetSheet.getRange(edit.range);
      targetSheet.autoResizeRows(r.getRow(), r.getNumRows());
      return true;
    }
  }

  if (edit.action === "addNamedRange" && edit.range && edit.namedRangeName) {
    spreadsheet.setNamedRange(edit.namedRangeName, targetSheet.getRange(edit.range));
    return true;
  }

  if (edit.action === "removeNamedRange" && edit.namedRangeName) {
    spreadsheet.getNamedRanges().forEach(nr => {
      if (nr.getName() === edit.namedRangeName) nr.remove();
    });
    return true;
  }

  return false;
}

function handleChartActions(spreadsheet, edit, state) {
  // 1. Basic Validation
  if (!edit || typeof edit !== 'object') return false;
  if (!edit.sheetName || !state.activeWritableSheets.includes(edit.sheetName)) return false;

  const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
  if (!targetSheet) return false;

  // 2. Handle Clear Charts
  if (edit.action === "clearCharts") {
    safeGetCharts(targetSheet).forEach(chart => targetSheet.removeChart(chart));
    return true;
  }

  // 3. Handle Add Chart
  if (edit.action === "addChart") {
    // Validate that we have enough data to make a chart
    if (!edit.chartData) {
      state.executionWarnings.push(`Chart creation on ${edit.sheetName} failed: Missing chartData.`);
      return true; // Return true to avoid throwing, but mark as warning
    }

    // SMART FALLBACK FOR ANCHOR RANGE
    // If AI forgets 'range', place chart below the data
    let anchorRange = edit.range;
    if (!anchorRange) {
      const lastRow = targetSheet.getLastRow();
      if (lastRow > 0) {
        // Place it 5 rows below the last data row, in column A
        anchorRange = `A${lastRow + 5}`;
        state.scopeChanges.push(`Note: Placed chart at ${anchorRange} as anchor range was not specified.`);
      } else {
        state.executionWarnings.push(`Chart on ${edit.sheetName} failed: Cannot determine anchor range for empty sheet.`);
        return true;
      }
    }

    // Parse chartData robustly
    let rangesToChart = [];

    if (Array.isArray(edit.chartData)) {
      rangesToChart = edit.chartData.filter(r => r && String(r).trim() !== "");
    } else if (typeof edit.chartData === 'string') {
      // Split by comma, handle potential spaces
      rangesToChart = edit.chartData.split(',').map(s => s.trim()).filter(s => s !== "");
    } else if (typeof edit.chartData === 'object') {
      // If it's an object, try to stringify and parse, or take first value
      try {
        const str = JSON.stringify(edit.chartData);
        rangesToChart = str.split(',').map(s => s.trim().replace(/"/g, '')).filter(s => s !== "");
      } catch (e) {
        // If stringify fails, try to get keys or values
        rangesToChart = [String(edit.chartData)];
      }
    } else {
      rangesToChart = [String(edit.chartData)];
    }

    if (rangesToChart.length === 0) {
      state.executionWarnings.push(`Chart creation on ${edit.sheetName} failed: Could not parse chartData.`);
      return true;
    }

    // Replace charts if requested
    if (edit.replaceCharts === true) {
      safeGetCharts(targetSheet).forEach(chart => targetSheet.removeChart(chart));
    }

    // Get Anchor Cell
    const anchorCell = targetSheet.getRange(anchorRange);
    let chartBuilder = targetSheet.newChart();

    // Add Ranges
    rangesToChart.forEach(rangeRef => {
      try {
        const rng = getRangeByMaybeQualifiedA1(targetSheet, rangeRef, state.allowedReadSheets, state.activeWritableSheets);
        chartBuilder.addRange(rng);
      } catch (e) {
        state.executionWarnings.push(`Chart range error: ${e.message} for ref "${rangeRef}"`);
      }
    });

    // Set Position
    chartBuilder.setPosition(anchorCell.getRow(), anchorCell.getColumn(), 0, 0);

    // Set Options
    if (edit.chartTitle) chartBuilder.setOption('title', edit.chartTitle);
    if (edit.xAxisTitle) chartBuilder.setOption('hAxis.title', edit.xAxisTitle);
    if (edit.yAxisTitle) chartBuilder.setOption('vAxis.title', edit.yAxisTitle);
    if (edit.legendPosition) chartBuilder.setOption('legend.position', edit.legendPosition);
    if (edit.chartWidth) chartBuilder.setOption('width', edit.chartWidth);
    if (edit.chartHeight) chartBuilder.setOption('height', edit.chartHeight);

    // Determine Chart Type
    const type = String(edit.chartType || 'COLUMN').toUpperCase();

    if (type === "PIE") chartBuilder = chartBuilder.asPieChart();
    else if (type === "BAR") chartBuilder = chartBuilder.asBarChart();
    else if (type === "LINE") chartBuilder = chartBuilder.asLineChart();
    else if (type === "AREA") chartBuilder = chartBuilder.asAreaChart();
    else if (type === "SCATTER") chartBuilder = chartBuilder.asScatterChart();
    else if (type === "COMBO") chartBuilder = chartBuilder.asComboChart();
    else if (type === "TABLE") chartBuilder = chartBuilder.asTableChart();
    else chartBuilder = chartBuilder.asColumnChart();

    // Build and Insert
    const chart = chartBuilder.build();
    targetSheet.insertChart(chart);

    return true;
  }

  return false;
}

function handleCellRangeActions(spreadsheet, edit, state) {
  if (!edit.sheetName || !edit.range) return false;
  if (!state.activeWritableSheets.includes(edit.sheetName)) return false;

  const targetSheet = spreadsheet.getSheetByName(edit.sheetName);
  if (!targetSheet) return false;

  let rawRange;
  try {
    rawRange = targetSheet.getRange(edit.range);
  } catch (e) {
    state.executionWarnings.push(`Invalid range '${edit.range}' on '${edit.sheetName}'. Skipping edit.`);
    return true; // Skip gracefully without crashing the whole script
  }

  const is2DArrayInsert = (edit.values2D !== undefined && Array.isArray(edit.values2D) && edit.values2D.length > 0);

  // --- THE ULTIMATE AUTO-RESIZE & MERGE CONFLICT FIX ---
  if (is2DArrayInsert) {
    const numRows = edit.values2D.length;
    const numCols = edit.values2D[0].length || 1;
    rawRange = targetSheet.getRange(rawRange.getRow(), rawRange.getColumn(), numRows, numCols);

    // CRITICAL: Shatter any merged cells in the exact path of this 2D grid BEFORE writing.
    // This prevents the "Cannot overwrite part of a merged cell" fatal crash.
    cleanupOverlappingMergedRanges(targetSheet, rawRange);
  }

  // Snap to merged boundaries ONLY if we are applying formatting or a single value.
  // If we are injecting a 2D grid, strictly use the mathematical grid we just built.
  const targetRange = is2DArrayInsert ? rawRange : getSafeTargetRange(rawRange);

  if (edit.action === "clearConditionalFormat") {
    clearConditionalFormatRulesForRange(targetSheet, targetRange);
    return true;
  }

  if (edit.action === "insertImage" && edit.imageUrl && edit.row && edit.column) {
    try {
      targetSheet.insertImage(edit.imageUrl, edit.column, edit.row);
    } catch (e) {
      state.executionWarnings.push(`Image insertion failed: ${e.toString()}`);
    }
    return true;
  }

  if (edit.action === "comment" && edit.note !== undefined) {
    targetRange.setNote(String(edit.note));
    return true;
  }

  // --- BULLETPROOF DATA INJECTION ---
  if (is2DArrayInsert) {
    // Note: We no longer need validate2DRangeFit() because we mathematically forced the range to fit perfectly.

    const containsFormula = edit.values2D.some(row => (row || []).some(cell => isFormulaString(cell)));

    const sanitized2D = edit.values2D.map(row =>
      (row || []).map(cell => {
        if (cell === null || cell === undefined) return "";
        // Google Sheets setFormulas() strictly requires strings. If there is a mix of numbers and formulas, force strings.
        if (containsFormula) return String(cell);
        // Otherwise, preserve true numbers/booleans for setValues()
        if (typeof cell === 'object') return JSON.stringify(cell);
        return cell;
      })
    );

    try {
      if (containsFormula) {
        // Scope permissions check
        sanitized2D.forEach(row => {
          row.forEach(cell => {
            if (isFormulaString(cell)) {
              const refCheck = sanitizeFormulaReferences(cell, state.allowedReadSheets, state.activeWritableSheets);
              if (!refCheck.ok) throw new Error(refCheck.reason);
            }
          });
        });
        // Write plain values first so non-formula cells are safe text/numbers,
        // then overwrite only formula cells individually to avoid setFormulas()
        // corrupting plain-text cells in a mixed array.
        const plainValues = sanitized2D.map(row =>
          row.map(cell => isFormulaString(cell) ? '' : cell)
        );
        targetRange.setValues(plainValues);
        for (let r = 0; r < sanitized2D.length; r++) {
          for (let c = 0; c < sanitized2D[r].length; c++) {
            if (isFormulaString(sanitized2D[r][c])) {
              targetRange.getCell(r + 1, c + 1).setFormula(sanitized2D[r][c]);
            }
          }
        }
      } else {
        targetRange.setValues(sanitized2D);
      }
    } catch (e) {
      state.executionWarnings.push(`Failed to write 2D data to ${edit.range}: ${e.message}`);
    }

  } else if (edit.value !== undefined) {
    try {
      if (isFormulaString(edit.value)) {
        const refCheck = sanitizeFormulaReferences(edit.value, state.allowedReadSheets, state.activeWritableSheets);
        if (!refCheck.ok) throw new Error(refCheck.reason);
        targetRange.setFormula(String(edit.value));
      } else {
        targetRange.setValue(edit.value);
      }
    } catch (e) {
      state.executionWarnings.push(`Failed to write value to ${edit.range}: ${e.message}`);
    }
  }

  // --- SAFE FORMATTING ISOLATION ---
  try {
    setRichFormatting(targetRange, edit);
    applyDataValidation(targetRange, edit, state);

    if (edit.conditionalFormat) {
      replaceOrAppendConditionalFormatRule(
        targetSheet,
        targetRange,
        edit.conditionalFormat,
        edit.replaceConditionalFormat === true
      );
    }
  } catch (e) {
    state.executionWarnings.push(`Formatting failed on ${edit.range}: ${e.message}`);
  }

  // --- QUEUE FORMULAS FOR AUTO-HEALING INSPECTION ---
  if (edit.value !== undefined && isFormulaString(edit.value)) {
    state.formulaRangesToInspect.push({ sheetName: edit.sheetName, range: targetRange.getA1Notation() });
  }
  if (is2DArrayInsert) {
    const hasFormula = edit.values2D.some(row => (row || []).some(cell => isFormulaString(cell)));
    if (hasFormula) {
      state.formulaRangesToInspect.push({ sheetName: edit.sheetName, range: targetRange.getA1Notation() });
    }
  }

  return true;
}
