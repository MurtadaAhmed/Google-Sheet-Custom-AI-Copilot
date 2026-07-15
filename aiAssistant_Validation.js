/* =========================
   FORMULA / VALIDATION HELPERS
========================= */

function getAllowedSheetNamesForReferences(existingAllowedSheets, newlyCreatedSheets) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allCurrent = spreadsheet.getSheets().map(s => s.getName());

  const out = [];
  existingAllowedSheets.concat(newlyCreatedSheets).forEach(name => {
    if (allCurrent.includes(name) && !out.includes(name)) out.push(name);
  });

  return out;
}

function extractFormulaSheetReferences(formula) {
  if (!formula || typeof formula !== 'string' || formula.trim().charAt(0) !== '=') return [];

  const refs = [];
  const quotedRegex = /'([^']+)'!/g;
  const simpleRegex = /\b([A-Za-z0-9_]+)!/g;

  let match;
  while ((match = quotedRegex.exec(formula)) !== null) refs.push(match[1]);
  while ((match = simpleRegex.exec(formula)) !== null) refs.push(match[1]);

  return refs.filter((name, idx, arr) => arr.indexOf(name) === idx);
}

function sanitizeFormulaReferences(formula, allowedReadSheets, writableNewSheets) {
  if (!formula || typeof formula !== 'string' || formula.charAt(0) !== '=') return { ok: true };

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allExistingSheets = spreadsheet.getSheets().map(s => s.getName());
  const allowedRefs = getAllowedSheetNamesForReferences(allowedReadSheets, writableNewSheets);

  const refs = extractFormulaSheetReferences(formula);
  const disallowed = refs.filter(ref => allExistingSheets.includes(ref) && !allowedRefs.includes(ref));

  if (disallowed.length > 0) {
    return {
      ok: false,
      reason: "Formula references sheet(s) outside allowed scope: " + disallowed.join(", ")
    };
  }

  return { ok: true };
}

function validateEdits(edits, allowedReadSheets, allowedWritableSheets) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const existingSheetNames = spreadsheet.getSheets().map(s => s.getName());
  const errors = [];
  const tempWritableSheets = allowedWritableSheets.slice();
  const tempCreatedSheets = [];

  (edits || []).forEach((edit, index) => {
    const label = `Edit ${index + 1}`;

    if (!edit || typeof edit !== 'object') {
      errors.push(`${label}: edit is not an object.`);
      return;
    }

    if (edit.action === "addSheet" && edit.sheetName) {
      if (!tempWritableSheets.includes(edit.sheetName)) tempWritableSheets.push(edit.sheetName);
      if (!tempCreatedSheets.includes(edit.sheetName)) tempCreatedSheets.push(edit.sheetName);
      return;
    }

    if (edit.action === "renameSheet") {
      if (!edit.sheetName || !edit.newSheetName) {
        errors.push(`${label}: renameSheet requires sheetName and newSheetName.`);
        return;
      }

      if (!tempWritableSheets.includes(edit.sheetName) && !existingSheetNames.includes(edit.sheetName)) {
        errors.push(`${label}: source sheet "${edit.sheetName}" does not exist.`);
        return;
      }

      const idx = tempWritableSheets.indexOf(edit.sheetName);
      if (idx !== -1) tempWritableSheets[idx] = edit.newSheetName;
      else if (!tempWritableSheets.includes(edit.newSheetName)) tempWritableSheets.push(edit.newSheetName);

      return;
    }

    if (edit.action === "duplicateSheet") {
      if (!edit.sheetName || !edit.newSheetName) {
        errors.push(`${label}: duplicateSheet requires sheetName and newSheetName.`);
        return;
      }

      if (!tempWritableSheets.includes(edit.sheetName) && !existingSheetNames.includes(edit.sheetName)) {
        errors.push(`${label}: source sheet "${edit.sheetName}" does not exist.`);
        return;
      }

      if (!tempWritableSheets.includes(edit.newSheetName)) tempWritableSheets.push(edit.newSheetName);
      return;
    }

    if (edit.action === "moveSheet") {
      if (!edit.sheetName || edit.newIndex === undefined) {
        errors.push(`${label}: moveSheet requires sheetName and newIndex.`);
      }
      return;
    }

    if (edit.action === "deleteSheet") {
      if (!edit.sheetName) errors.push(`${label}: deleteSheet requires sheetName.`);
      return;
    }

    if (edit.sheetName && !tempWritableSheets.includes(edit.sheetName) && !existingSheetNames.includes(edit.sheetName)) {
      errors.push(`${label}: target sheet "${edit.sheetName}" is not writable and does not exist.`);
      return;
    }

    if (edit.value && typeof edit.value === 'string' && edit.value.charAt(0) === '=') {
      const formulaScopeCheck = sanitizeFormulaReferences(edit.value, allowedReadSheets, tempWritableSheets);
      if (!formulaScopeCheck.ok) errors.push(`${label}: ${formulaScopeCheck.reason}`);
    }

    if (edit.values2D && edit.values2D.length > 0) {
      for (let r = 0; r < edit.values2D.length; r++) {
        for (let c = 0; c < edit.values2D[r].length; c++) {
          const cell = edit.values2D[r][c];
          if (typeof cell === 'string' && cell.charAt(0) === '=') {
            const formulaScopeCheck = sanitizeFormulaReferences(cell, allowedReadSheets, tempWritableSheets);
            if (!formulaScopeCheck.ok) errors.push(`${label}: ${formulaScopeCheck.reason}`);
          }
        }
      }
    }
  });

  return errors;
}

function normalizeA1RangeString(a1Notation) {
  return String(a1Notation || '').trim();
}

function getRangeByMaybeQualifiedA1(defaultSheet, ref, allowedReadSheets, writableSheets) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const combinedAllowed = getAllowedSheetNamesForReferences(allowedReadSheets, writableSheets);

  const text = normalizeA1RangeString(ref);
  if (!text) throw new Error("Empty range reference.");

  const quotedMatch = text.match(/^'([^']+)'!(.+)$/);
  if (quotedMatch) {
    const sheetName = quotedMatch[1];
    const rangeA1 = quotedMatch[2];
    if (!combinedAllowed.includes(sheetName)) throw new Error(`Chart/data range references disallowed sheet "${sheetName}".`);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    return sheet.getRange(rangeA1);
  }

  const simpleMatch = text.match(/^([A-Za-z0-9_]+)!(.+)$/);
  if (simpleMatch) {
    const sheetName = simpleMatch[1];
    const rangeA1 = simpleMatch[2];
    if (!combinedAllowed.includes(sheetName)) throw new Error(`Chart/data range references disallowed sheet "${sheetName}".`);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" not found.`);
    return sheet.getRange(rangeA1);
  }

  return defaultSheet.getRange(text);
}

function isFormulaString(value) {
  return typeof value === 'string' && value.trim().charAt(0) === '=';
}

function detectRangeFormulaIssues(targetRange) {
  const issues = [];
  try {
    const displayValues = targetRange.getDisplayValues();
    const formulas = targetRange.getFormulas();

    for (let r = 0; r < displayValues.length; r++) {
      for (let c = 0; c < displayValues[r].length; c++) {
        if (formulas[r][c]) {
          const dv = String(displayValues[r][c] || '');
          if (/^#(ERROR|REF|N\/A|VALUE|NAME\?|DIV\/0!|NUM!|NULL!|SPILL!)/i.test(dv)) {
            issues.push({
              rowOffset: r,
              colOffset: c,
              formula: formulas[r][c],
              displayValue: dv
            });
          }
        }
      }
    }
  } catch (e) {
    issues.push({
      rowOffset: 0,
      colOffset: 0,
      formula: '',
      displayValue: 'Formula inspection failed: ' + e.toString()
    });
  }
  return issues;
}
