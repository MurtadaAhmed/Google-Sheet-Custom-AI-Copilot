function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🤖 AI Assistant')
    .addItem('Open Chat', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Sheet AI Assistant')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* =========================
   SIDEBAR STATE / SETTINGS
========================= */

function getSidebarState() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const userProps = PropertiesService.getUserProperties();
  let selectedSheets = [];
  let sessionHistory = [];

  try {
    selectedSheets = JSON.parse(userProps.getProperty('AI_SELECTED_SHEETS') || '[]');
  } catch (e) {
    selectedSheets = [];
  }

  try {
    sessionHistory = JSON.parse(userProps.getProperty('AI_SESSION_HISTORY') || '[]');
  } catch (e) {
    sessionHistory = [];
  }

  selectedSheets = selectedSheets.filter(name => allSheetNames.includes(name));

  return {
    allSheetNames: allSheetNames,
    selectedSheets: selectedSheets,
    sessionHistory: sessionHistory
  };
}

function saveSidebarState(selectedSheets, sessionHistory) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const realSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const safeSelectedSheets = Array.isArray(selectedSheets)
    ? selectedSheets.filter(name => realSheetNames.includes(name))
    : [];

  const safeSessionHistory = Array.isArray(sessionHistory) ? sessionHistory : [];

  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('AI_SELECTED_SHEETS', JSON.stringify(safeSelectedSheets));
  userProps.setProperty('AI_SESSION_HISTORY', JSON.stringify(safeSessionHistory));
  return true;
}

function clearSidebarState() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty('AI_SELECTED_SHEETS');
  userProps.deleteProperty('AI_SESSION_HISTORY');
  return true;
}

function getSheetNames() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getSheets().map(sheet => sheet.getName());
}

function updateSelectedSheets(selectedSheets) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const realSheetNames = spreadsheet.getSheets().map(s => s.getName());

  const safeSelectedSheets = Array.isArray(selectedSheets)
    ? selectedSheets.filter(name => realSheetNames.includes(name))
    : [];

  const userProps = PropertiesService.getUserProperties();
  userProps.setProperty('AI_SELECTED_SHEETS', JSON.stringify(safeSelectedSheets));
  return safeSelectedSheets;
}

/* =========================
   API / PARSING
========================= */

function fetchAiResponse(messagesArray) {
  const scriptProps = PropertiesService.getScriptProperties();

  const apiKey = scriptProps.getProperty('AI_API_KEY');
  if (!apiKey) throw new Error("AI_API_KEY not found in Script Properties.");

  const url = scriptProps.getProperty('AI_API_ENDPOINT');
  if (!url) throw new Error("AI_API_ENDPOINT not found in Script Properties.");

  const model = scriptProps.getProperty('AI_MODEL');
  if (!model) throw new Error("AI_MODEL not found in Script Properties.");

  const payload = {
    model: model,
    messages: messagesArray,
    temperature: 0.1
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const httpCode = response.getResponseCode();
  const raw = response.getContentText();

  let parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (e) {
    throw new Error("API returned non-JSON response. HTTP " + httpCode + ". Raw: " + raw);
  }

  parsed._meta = {
    httpCode: httpCode,
    requestBytes: JSON.stringify(payload).length,
    responseBytes: raw.length
  };

  return parsed;
}

function extractMessageContent(content) {
  if (typeof content === 'string') return content;

  if (Array.isArray(content)) {
    return content.map(part => {
      if (typeof part === 'string') return part;
      if (part && typeof part.text === 'string') return part.text;
      if (part && part.type === 'text' && typeof part.content === 'string') return part.content;
      return '';
    }).join('\n').trim();
  }

  if (content && typeof content.text === 'string') return content.text;

  return String(content || '');
}

function parseAiResponse(json) {
  if (json.error) {
    return {
      message: "API Error: " + (json.error.message || JSON.stringify(json.error)),
      edits: [],
      usage: json.usage || json._meta || {}
    };
  }

  if (!json.choices || !json.choices[0] || !json.choices[0].message) {
    return {
      message: "Unexpected API format.",
      edits: [],
      usage: json.usage || json._meta || {}
    };
  }

  let aiRawReply = extractMessageContent(json.choices[0].message.content).trim();
  let aiResponseObj = { message: "", edits: [] };

  const jsonStartIndex = aiRawReply.indexOf('{');
  const jsonEndIndex = aiRawReply.lastIndexOf('}');

  if (jsonStartIndex !== -1 && jsonEndIndex !== -1 && jsonEndIndex >= jsonStartIndex) {
    try {
      aiResponseObj = JSON.parse(aiRawReply.substring(jsonStartIndex, jsonEndIndex + 1));
    } catch (e) {
      aiResponseObj = {
        message: "JSON formatting error. Raw text: " + aiRawReply,
        edits: []
      };
    }
  } else {
    aiResponseObj = {
      message: aiRawReply,
      edits: []
    };
  }

  if (!Array.isArray(aiResponseObj.edits)) aiResponseObj.edits = [];
  if (typeof aiResponseObj.message !== 'string') aiResponseObj.message = String(aiResponseObj.message || '');
  aiResponseObj.usage = json.usage || json._meta || {};

  return aiResponseObj;
}

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

/* =========================
   UTILITIES
========================= */

function dedupeArray(arr) {
  return (arr || []).filter((v, i, a) => a.indexOf(v) === i);
}

function clamp(n, min, max) {
  return Math.max(min, Math.min(max, n));
}

function safeBoolean(value, defaultValue) {
  if (typeof value === 'boolean') return value;
  return defaultValue;
}

function overlapsRange(a, b) {
  const aRow1 = a.getRow();
  const aCol1 = a.getColumn();
  const aRow2 = aRow1 + a.getNumRows() - 1;
  const aCol2 = aCol1 + a.getNumColumns() - 1;

  const bRow1 = b.getRow();
  const bCol1 = b.getColumn();
  const bRow2 = bRow1 + b.getNumRows() - 1;
  const bCol2 = bCol1 + b.getNumColumns() - 1;

  const rowsOverlap = aRow1 <= bRow2 && aRow2 >= bRow1;
  const colsOverlap = aCol1 <= bCol2 && aCol2 >= bCol1;

  return rowsOverlap && colsOverlap;
}

function cleanupOverlappingMergedRanges(targetSheet, targetRange) {
  const mergedRanges = targetSheet.getDataRange().getMergedRanges();
  mergedRanges.forEach(mr => {
    if (overlapsRange(mr, targetRange)) {
      try {
        mr.breakApart();
      } catch (e) {}
    }
  });
}

function getSafeTargetRange(targetRange) {
  try {
    const mergedRanges = targetRange.getMergedRanges();
    if (mergedRanges && mergedRanges.length > 0) {
      return mergedRanges[0];
    }
  } catch (e) {}
  return targetRange;
}

/* =========================
   CONDITIONAL FORMATTING
========================= */

function buildConditionalFormatRule(targetRange, conditionalFormat) {
  let ruleBuilder = SpreadsheetApp.newConditionalFormatRule().setRanges([targetRange]);

  if (conditionalFormat.color) ruleBuilder.setBackground(conditionalFormat.color);
  if (conditionalFormat.fontColor) ruleBuilder.setFontColor(conditionalFormat.fontColor);
  if (conditionalFormat.bold === true) ruleBuilder.setBold(true);
  if (conditionalFormat.italic === true) ruleBuilder.setItalic(true);
  if (conditionalFormat.underline === true) ruleBuilder.setUnderline(true);
  if (conditionalFormat.strikethrough === true) ruleBuilder.setStrikethrough(true);

  const type = conditionalFormat.type;

  if (type === "greaterThan") {
    ruleBuilder.whenNumberGreaterThan(parseFloat(conditionalFormat.value));
  } else if (type === "greaterThanOrEqualTo") {
    ruleBuilder.whenNumberGreaterThanOrEqualTo(parseFloat(conditionalFormat.value));
  } else if (type === "lessThan") {
    ruleBuilder.whenNumberLessThan(parseFloat(conditionalFormat.value));
  } else if (type === "lessThanOrEqualTo") {
    ruleBuilder.whenNumberLessThanOrEqualTo(parseFloat(conditionalFormat.value));
  } else if (type === "equalTo") {
    ruleBuilder.whenNumberEqualTo(parseFloat(conditionalFormat.value));
  } else if (type === "notEqualTo") {
    ruleBuilder.whenNumberNotEqualTo(parseFloat(conditionalFormat.value));
  } else if (type === "numberBetween") {
    ruleBuilder.whenNumberBetween(parseFloat(conditionalFormat.min), parseFloat(conditionalFormat.max));
  } else if (type === "numberNotBetween") {
    ruleBuilder.whenNumberNotBetween(parseFloat(conditionalFormat.min), parseFloat(conditionalFormat.max));
  } else if (type === "textEqualTo") {
    ruleBuilder.whenTextEqualTo(String(conditionalFormat.value));
  } else if (type === "textContains") {
    ruleBuilder.whenTextContains(String(conditionalFormat.value));
  } else if (type === "textStartsWith") {
    ruleBuilder.whenTextStartsWith(String(conditionalFormat.value));
  } else if (type === "textEndsWith") {
    ruleBuilder.whenTextEndsWith(String(conditionalFormat.value));
  } else if (type === "empty") {
    ruleBuilder.whenCellEmpty();
  } else if (type === "notEmpty") {
    ruleBuilder.whenCellNotEmpty();
  } else if (type === "dateBefore") {
    ruleBuilder.whenDateBefore(new Date(conditionalFormat.value));
  } else if (type === "dateAfter") {
    ruleBuilder.whenDateAfter(new Date(conditionalFormat.value));
  } else if (type === "customFormula" && conditionalFormat.formula) {
    ruleBuilder.whenFormulaSatisfied(conditionalFormat.formula);
  } else {
    throw new Error("Unsupported conditionalFormat type: " + type);
  }

  return ruleBuilder.build();
}

function replaceOrAppendConditionalFormatRule(targetSheet, targetRange, conditionalFormat, replaceExistingForRange) {
  let rules = targetSheet.getConditionalFormatRules();

  if (replaceExistingForRange) {
    const a1 = targetRange.getA1Notation();
    rules = rules.filter(rule => {
      const ranges = rule.getRanges() || [];
      return !ranges.some(r => r.getA1Notation() === a1 && r.getSheet().getName() === targetSheet.getName());
    });
  }

  rules.push(buildConditionalFormatRule(targetRange, conditionalFormat));
  targetSheet.setConditionalFormatRules(rules);
}

function clearConditionalFormatRulesForRange(targetSheet, targetRange) {
  const a1 = targetRange.getA1Notation();
  const rules = targetSheet.getConditionalFormatRules().filter(rule => {
    const ranges = rule.getRanges() || [];
    return !ranges.some(r => r.getA1Notation() === a1 && r.getSheet().getName() === targetSheet.getName());
  });
  targetSheet.setConditionalFormatRules(rules);
}

/* =========================
   RANGE FORMATTING / VALIDATION
========================= */

function setRichFormatting(targetRange, edit) {
  if (edit.backgroundColor) targetRange.setBackground(edit.backgroundColor);
  if (edit.fontColor) targetRange.setFontColor(edit.fontColor);
  if (edit.fontWeight) targetRange.setFontWeight(edit.fontWeight);
  if (edit.horizontalAlignment) targetRange.setHorizontalAlignment(edit.horizontalAlignment);
  if (edit.numberFormat) targetRange.setNumberFormat(edit.numberFormat);

  if (edit.fontSize !== undefined) targetRange.setFontSize(edit.fontSize);
  if (edit.fontFamily) targetRange.setFontFamily(edit.fontFamily);
  if (edit.fontStyle) targetRange.setFontStyle(edit.fontStyle);
  if (edit.underline !== undefined || edit.strikethrough !== undefined) {
    if (edit.strikethrough === true) targetRange.setFontLine("line-through");
    else if (edit.underline === true) targetRange.setFontLine("underline");
    else targetRange.setFontLine("none");
  }
  if (edit.verticalAlignment) targetRange.setVerticalAlignment(edit.verticalAlignment);
  if (edit.wrap !== undefined) targetRange.setWrap(safeBoolean(edit.wrap, true));
  if (edit.wrapStrategy && SpreadsheetApp.WrapStrategy && SpreadsheetApp.WrapStrategy[edit.wrapStrategy]) {
    targetRange.setWrapStrategy(SpreadsheetApp.WrapStrategy[edit.wrapStrategy]);
  }
  if (edit.textRotation !== undefined) targetRange.setTextRotation(edit.textRotation);

  if (edit.note !== undefined) targetRange.setNote(String(edit.note));

  if (edit.border) {
    targetRange.setBorder(
      !!edit.border.top,
      !!edit.border.left,
      !!edit.border.bottom,
      !!edit.border.right,
      !!edit.border.vertical,
      !!edit.border.horizontal,
      edit.border.color || null,
      edit.border.style && SpreadsheetApp.BorderStyle && SpreadsheetApp.BorderStyle[edit.border.style]
        ? SpreadsheetApp.BorderStyle[edit.border.style]
        : null
    );
  }
}

function applyDataValidation(targetRange, edit) {
  if (!edit.dataValidation) return;

  let rule;
  if (edit.dataValidation.type === "dropdown" && Array.isArray(edit.dataValidation.values)) {
    rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(edit.dataValidation.values, true);
    if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
    targetRange.setDataValidation(rule.build());
  } else if (edit.dataValidation.type === "dropdownFromRange" && edit.dataValidation.sourceSheet && edit.dataValidation.sourceRange) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(edit.dataValidation.sourceSheet);
    if (sourceSheet) {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceSheet.getRange(edit.dataValidation.sourceRange), true);
      if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
      targetRange.setDataValidation(rule.build());
    }
  } else if (edit.dataValidation.type === "checkbox") {
    rule = SpreadsheetApp.newDataValidation().requireCheckbox();
    if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
    targetRange.setDataValidation(rule.build());
  }
}

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
    applyDataValidation(targetRange, edit);

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

/* =========================
   EXECUTION ENGINE
========================= */

function applyEditsToSheet(spreadsheet, edits, allowedReadSheets) {
  if (!edits || edits.length === 0) {
    return {
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

    if (handleSheetStructureAction(spreadsheet, edit, state)) continue;
    if (handleFreezeClearAppendProtectActions(spreadsheet, edit, state)) continue;
    if (handleMergeResizeNamedRangeActions(spreadsheet, edit, state)) continue;
    if (handleChartActions(spreadsheet, edit, state)) continue;
    if (handleCellRangeActions(spreadsheet, edit, state)) continue;
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
  const survivingBaseSheets = originalSelectedSheets.filter(name => !executionState.deletedSheets.includes(name));
  const effectiveReadable = dedupeArray(survivingBaseSheets.concat(executionState.newlyAddedSheets));
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

/* =========================
   MAIN CHAT PROCESSOR
========================= */

function processChat(sessionHistory, selectedSheets) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const existingSheets = spreadsheet.getSheets().map(s => s.getName());

    if (!selectedSheets || !Array.isArray(selectedSheets) || selectedSheets.length === 0) {
      return JSON.stringify({
        message: "Please select at least one sheet before starting the chat.",
        edits: [],
        newlyAddedSheets: [],
        scopeChanges: [],
        usage: {}
      });
    }

    selectedSheets = selectedSheets.filter(name => existingSheets.includes(name));

    if (selectedSheets.length === 0) {
      return JSON.stringify({
        message: "None of the selected sheets were found.",
        edits: [],
        newlyAddedSheets: [],
        scopeChanges: [],
        usage: {}
      });
    }

    const contextString = buildContextForSheets(selectedSheets);

    const systemPrompt = `You are an expert AI Google Sheets orchestrator.

READ ACCESS:
- You may read and reference ONLY these originally selected input sheets:
${selectedSheets.join(", ")}

WRITE ACCESS:
- You may modify these selected sheets.
- You may create NEW sheets for outputs like dashboards, summaries, reports, or transformed tables.
- Newly created sheets become writable immediately.
- If a new sheet is created, later edits in the SAME response may write to it.

STRICT SCOPE RULES:
1. Never read from, reference, or use any pre-existing sheet outside the allowed input sheet list above.
2. Never produce formulas that reference any pre-existing disallowed sheet.
3. If the user asks about a non-selected pre-existing sheet, explain that it is outside scope.
4. Return ONLY one raw JSON object. No markdown fences. No extra commentary outside JSON.

WORKBOOK DATA IN SCOPE:
${contextString}

CONTEXT GUIDANCE:
- Prefer batched edits over many small edits.
- Use multi-cell tables when helpful with "values2D".
- If you create a sheet, you may immediately populate it.
- If useful, include formatting, sizing, merges, named ranges, notes, charts, and protections.
- When working with merged cells, target the full merged range for later formatting/value edits.
- For conditional formatting, use explicit exact ranges and explicit rule types.

SUPPORTED EDITS:
- Structural: addSheet, deleteSheet, renameSheet, duplicateSheet, moveSheet
- Sheet: freeze, clear, clearContent, clearCharts
- Data: appendRow, appendRows
- Protection: protectRange
- Merge: merge, unmerge
- Resize: resizeColumn, resizeColumns, resizeRow, resizeRows, autoResizeColumns, autoResizeRows
- For autoResizeColumns and autoResizeRows, you may use either:
  - startColumn + numColumns / startRow + numRows
  - or a range like "A:H" or "2:10"
- Named ranges: addNamedRange, removeNamedRange
- Charts: addChart
  REQUIRED: You MUST provide "sheetName", "range" (the anchor cell like "A15"), "chartData" (e.g., "A1:B10"), and "chartType" (e.g., PIE, BAR).
  Example:
  {"action": "addChart", "sheetName": "Sheet9", "range": "D2", "chartData": "A1:B10", "chartType": "PIE", "chartTitle": "Salaries"}
  Note: "range" is the top-left cell where the chart image will be placed.
- Range edits: use sheetName + range with value OR values2D
- Formatting keys:
  backgroundColor, fontColor, fontWeight, fontSize, fontFamily, fontStyle,
  underline, strikethrough, horizontalAlignment, verticalAlignment,
  numberFormat, wrap, wrapStrategy, textRotation, border, note
- Data validation:
  {"type":"dropdown","values":["A","B"]}
  {"type":"dropdownFromRange","sourceSheet":"Sheet1","sourceRange":"A1:A10"}
  {"type":"checkbox"}
- Conditional formatting:
  Put it on a specific exact range and use one of these types:
  greaterThan, greaterThanOrEqualTo, lessThan, lessThanOrEqualTo, equalTo, notEqualTo,
  numberBetween, numberNotBetween, textEqualTo, textContains, textStartsWith, textEndsWith,
  empty, notEmpty, dateBefore, dateAfter, customFormula
  Example:
  {"sheetName":"Sheet1","range":"C2:C20","conditionalFormat":{"type":"greaterThan","value":"100","color":"#c6efce"},"replaceConditionalFormat":true}
- Images:
  {"action":"insertImage","sheetName":"Dashboard","imageUrl":"https://...","row":1,"column":1}
- Notes/comments:
  Use "note" on a range edit, or action "comment" with note

IMPORTANT DATA RULES:
- BATCHING IS MANDATORY: Never create individual edit objects for single cells (e.g., A8, A9, A10...). Always use "values2D" to pass the entire data grid in one single edit object.
- FORMULA ARRAYS: If writing formulas into a table, use one single cell edit with a spilling array formula (e.g., =ARRAYFORMULA(...)) in the top-left cell of the range.
- DO NOT list individual cells for data or formulas. If you exceed 5 individual cell edits for the same table, you are failing the batching requirement.
- FAILURE TO BATCH will result in JSON truncation and script failure.
- NEVER use per-cell cross-sheet reference formulas (e.g., =Sheet10!B2, =Sheet10!C3) inside a values2D array. This causes #ERROR! on non-formula cells. Instead, embed literal values directly — they are already visible in the workbook context above. For computed columns, use a single ARRAYFORMULA in the top-left cell only.

FORMAT EXAMPLE:
{
  "message": "I built a dashboard and added formatting.",
  "edits": [
    { "action": "addSheet", "sheetName": "Dashboard" },
    { "action": "merge", "sheetName": "Dashboard", "range": "A1:D1" },
    { "sheetName": "Dashboard", "range": "A1:D1", "value": "Sales Dashboard", "fontWeight": "bold", "fontSize": 14, "horizontalAlignment": "center" },
    { "sheetName": "Sheet1", "range": "C2:C20", "conditionalFormat": { "type": "greaterThan", "value": "100", "color": "#c6efce" }, "replaceConditionalFormat": true }
  ]
}`;

    let messagesPayload = [{ role: "system", content: systemPrompt }];
    if (Array.isArray(sessionHistory) && sessionHistory.length > 0) {
      messagesPayload = messagesPayload.concat(sessionHistory);
    }

    const initialResponse = fetchAiResponse(messagesPayload);
    const aiResponseObj = parseAiResponse(initialResponse);

    const validationErrors = validateEdits(aiResponseObj.edits, selectedSheets, selectedSheets.slice());
    if (validationErrors.length > 0) {
      return JSON.stringify({
        message: "AI response failed validation:\n- " + validationErrors.join("\n- "),
        edits: [],
        newlyAddedSheets: [],
        scopeChanges: [],
        usage: aiResponseObj.usage || {}
      });
    }

    const executionState = applyEditsToSheet(spreadsheet, aiResponseObj.edits, selectedSheets);

    SpreadsheetApp.flush();

    const rebuilt = rebuildAllowedContextAfterEdits(selectedSheets, executionState);
    const formulaErrors = collectFormulaIssues(executionState.formulaRangesToInspect);

    if (formulaErrors.length > 0) {
      const correctionPrompt = `Your last edit caused formula/data issues.

Updated in-scope workbook context after your last edits:
${rebuilt.contextString}

Detected issues:
${formulaErrors.join("\n")}

Fix these issues and return a NEW JSON object with corrected edits only.

Rules:
- You may read only from these original input sheets: ${selectedSheets.join(", ")}.
- You may also write to these newly created sheets: ${executionState.newlyAddedSheets.join(", ") || "[none]"}.
- Do not reference any disallowed existing sheet.
- If correcting an array/spill issue, ensure the range size or formula placement is valid.`;

      const correctionMessages = [
        ...messagesPayload,
        { role: "assistant", content: extractMessageContent(initialResponse.choices[0].message.content) },
        { role: "user", content: correctionPrompt }
      ];

      const correctionJson = fetchAiResponse(correctionMessages);
      const correctedResponseObj = parseAiResponse(correctionJson);

      const secondValidationErrors = validateEdits(
        correctedResponseObj.edits,
        selectedSheets,
        rebuilt.effectiveReadableSheets.slice()
      );

      if (secondValidationErrors.length > 0) {
        return JSON.stringify({
          message: "Initial edits were applied, but correction attempt failed validation:\n- " + secondValidationErrors.join("\n- "),
          edits: aiResponseObj.edits,
          newlyAddedSheets: executionState.newlyAddedSheets,
          scopeChanges: executionState.scopeChanges,
          machineSummary: {
            warnings: executionState.executionWarnings,
            formulaErrors: formulaErrors
          },
          usage: correctedResponseObj.usage || aiResponseObj.usage || {}
        });
      }

      const correctionState = applyEditsToSheet(spreadsheet, correctedResponseObj.edits, rebuilt.effectiveReadableSheets);
      SpreadsheetApp.flush();

      const finalScopeChanges = dedupeArray(executionState.scopeChanges.concat(correctionState.scopeChanges));
      const finalNewSheets = dedupeArray(executionState.newlyAddedSheets.concat(correctionState.newlyAddedSheets));
      const finalWarnings = dedupeArray(executionState.executionWarnings.concat(correctionState.executionWarnings));

      return JSON.stringify({
        message: (correctedResponseObj.message || "I corrected formula issues.") + "\n\n(Self-corrected a formula/data issue during execution.)",
        edits: correctedResponseObj.edits,
        newlyAddedSheets: finalNewSheets,
        deletedSheets: dedupeArray(executionState.deletedSheets.concat(correctionState.deletedSheets)),
        scopeChanges: finalScopeChanges,
        updatedScope: dedupeArray(
          selectedSheets
            .filter(name => !executionState.deletedSheets.includes(name))
            .filter(name => !correctionState.deletedSheets.includes(name))
            .concat(finalNewSheets)
        ),
        machineSummary: {
          warnings: finalWarnings,
          formulaErrors: formulaErrors
        },
        usage: correctedResponseObj.usage || aiResponseObj.usage || {}
      });
    }

    return JSON.stringify({
      message: aiResponseObj.message,
      edits: aiResponseObj.edits,
      newlyAddedSheets: executionState.newlyAddedSheets,
      deletedSheets: executionState.deletedSheets,
      scopeChanges: executionState.scopeChanges,
      updatedScope: dedupeArray(
        selectedSheets
          .filter(name => !executionState.deletedSheets.includes(name))
          .concat(executionState.newlyAddedSheets)
      ),
      machineSummary: {
        warnings: executionState.executionWarnings
      },
      usage: aiResponseObj.usage || {}
    });

  } catch (error) {
    return JSON.stringify({
      message: "System Error: " + error.toString(),
      edits: [],
      newlyAddedSheets: [],
      deletedSheets: [],
      scopeChanges: [],
      updatedScope: [],
      machineSummary: { warnings: [] },
      usage: {}
    });
  }
}
