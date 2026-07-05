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
  } else if (type === "customFormula") {
    const formula = conditionalFormat.formula || conditionalFormat.formulaString || conditionalFormat.condition || conditionalFormat.expression;
    if (!formula) throw new Error("customFormula type requires a 'formula' field (e.g. =$C2>320)");
    ruleBuilder.whenFormulaSatisfied(formula);
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
  const tr1 = targetRange.getRow();
  const tc1 = targetRange.getColumn();
  const tr2 = tr1 + targetRange.getNumRows() - 1;
  const tc2 = tc1 + targetRange.getNumColumns() - 1;
  const sheetName = targetSheet.getName();

  const rules = targetSheet.getConditionalFormatRules().filter(rule => {
    const ranges = rule.getRanges() || [];
    return !ranges.some(r => {
      if (r.getSheet().getName() !== sheetName) return false;
      const rr1 = r.getRow();
      const rc1 = r.getColumn();
      const rr2 = rr1 + r.getNumRows() - 1;
      const rc2 = rc1 + r.getNumColumns() - 1;
      // Remove rule if its range intersects the target range
      return rr1 <= tr2 && rr2 >= tr1 && rc1 <= tc2 && rc2 >= tc1;
    });
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

function applyDataValidation(targetRange, edit, state) {
  if (!edit.dataValidation) return;

  const allowedSheets = state
    ? [].concat(state.allowedReadSheets || [], state.activeWritableSheets || [])
    : null;

  let rule;
  if (edit.dataValidation.type === "dropdown" && Array.isArray(edit.dataValidation.values)) {
    rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(edit.dataValidation.values, true);
    if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
    targetRange.setDataValidation(rule.build());
  } else if (edit.dataValidation.type === "dropdownFromRange" && edit.dataValidation.sourceSheet && edit.dataValidation.sourceRange) {
    if (allowedSheets && !allowedSheets.includes(edit.dataValidation.sourceSheet)) {
      if (state) state.executionWarnings.push(`Data validation rejected: sourceSheet "${edit.dataValidation.sourceSheet}" is not in scope.`);
      return;
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(edit.dataValidation.sourceSheet);
    if (sourceSheet) {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(sourceSheet.getRange(edit.dataValidation.sourceRange), true);
      if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
      targetRange.setDataValidation(rule.build());
    } else if (state) {
      state.executionWarnings.push(`Data validation skipped: sourceSheet "${edit.dataValidation.sourceSheet}" not found.`);
    }
  } else if (edit.dataValidation.type === "dropdownFromNamedRange" && edit.dataValidation.namedRangeName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedRange = ss.getRangeByName(edit.dataValidation.namedRangeName);
    if (namedRange) {
      const backingSheet = namedRange.getSheet().getName();
      if (allowedSheets && !allowedSheets.includes(backingSheet)) {
        if (state) state.executionWarnings.push(`Data validation rejected: named range "${edit.dataValidation.namedRangeName}" references out-of-scope sheet "${backingSheet}".`);
        return;
      }
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(namedRange, true);
      if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
      targetRange.setDataValidation(rule.build());
    } else if (state) {
      state.executionWarnings.push(`Data validation skipped: named range "${edit.dataValidation.namedRangeName}" not found. If created via addNamedRange in the same response, ensure that edit appears first.`);
    }
  } else if (edit.dataValidation.type === "checkbox") {
    rule = SpreadsheetApp.newDataValidation().requireCheckbox();
    if (edit.dataValidation.helpText) rule.setHelpText(edit.dataValidation.helpText);
    targetRange.setDataValidation(rule.build());
  }
}
