/* =========================
   EXECUTION ENGINE
========================= */

function applyEditsToSheet(spreadsheet, edits, allowedReadSheets) {
  if (!edits || edits.length === 0) {
    return {
      activeWritableSheets: allowedReadSheets ? allowedReadSheets.slice() : [],
      newlyAddedSheets: [],
      deletedSheets: [],
      scopeChanges: [],
      formulaRangesToInspect: [],
      executionWarnings: []
    };
  }

  const state = {
    allowedReadSheets: allowedReadSheets.slice(),
    activeWritableSheets: allowedReadSheets.slice(),
    newlyAddedSheets: [],
    deletedSheets: [],
    scopeChanges: [],
    formulaRangesToInspect: [],
    executionWarnings: []
  };

  for (let i = 0; i < edits.length; i++) {
    const edit = edits[i];
    if (!edit || typeof edit !== 'object') continue;

    const handled =
      handleSheetStructureAction(spreadsheet, edit, state) ||
      handleFreezeClearAppendProtectActions(spreadsheet, edit, state) ||
      handleMergeResizeNamedRangeActions(spreadsheet, edit, state) ||
      handleChartActions(spreadsheet, edit, state) ||
      handleCellRangeActions(spreadsheet, edit, state);

    if (!handled) {
      const target = edit.sheetName || '(no sheet)';
      const action = edit.action || (edit.range ? 'range edit on ' + edit.range : 'unknown');
      state.executionWarnings.push(`Edit ${i + 1} was not applied — sheet "${target}" may be out of scope or not found (action: ${action}).`);
    }
  }

  state.newlyAddedSheets = dedupeArray(state.newlyAddedSheets);
  state.deletedSheets = dedupeArray(state.deletedSheets);
  state.scopeChanges = dedupeArray(state.scopeChanges);

  return state;
}

/* =========================
   FORMULA CORRECTION SUPPORT
========================= */

function rebuildAllowedContextAfterEdits(originalSelectedSheets, executionState) {
  // Use activeWritableSheets so renames are reflected (it is updated in-place by renameSheet).
  // Fall back to original list if activeWritableSheets is somehow absent.
  const base = (executionState.activeWritableSheets && executionState.activeWritableSheets.length > 0)
    ? executionState.activeWritableSheets
    : originalSelectedSheets;
  const effectiveReadable = dedupeArray(base.filter(name => !executionState.deletedSheets.includes(name)));
  return {
    effectiveReadableSheets: effectiveReadable,
    contextString: buildContextForSheets(effectiveReadable)
  };
}

function collectFormulaIssues(formulaRangesToInspect) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const errors = [];

  (formulaRangesToInspect || []).forEach(item => {
    const sheet = spreadsheet.getSheetByName(item.sheetName);
    if (!sheet) return;

    const range = sheet.getRange(item.range);
    const issues = detectRangeFormulaIssues(range);

    issues.forEach(issue => {
      errors.push(
        `In sheet '${item.sheetName}' at range '${item.range}', formula '${issue.formula}' produced '${issue.displayValue}'.`
      );
    });
  });

  return errors;
}
