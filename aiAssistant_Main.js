/* =========================
   ENTRY POINTS
========================= */

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
- DATA SCALE CONSISTENCY: When generating dummy or sample data, store duration/time columns (e.g., Average Handle Time) as plain numbers in a consistent unit (e.g., seconds: 265, 312, 198) — NOT as TIME() values or formatted strings. This avoids h:mm vs m:ss ambiguity and makes formula thresholds straightforward (e.g., C2<300 instead of C2<TIME(0,5,0)).
- THRESHOLD CALIBRATION: Before writing any ARRAYFORMULA or conditional format threshold that references existing data, look at the actual values in WORKBOOK DATA IN SCOPE above. If a user specifies a threshold on a different scale than the data (e.g., user says 'CSAT > 90' but data is on a 1–5 scale, or 'AHT < 300' but AHT is stored as h:mm TIME values), adapt both consistently. Do not blindly apply the user's raw number to a mismatched scale.

SUPPORTED EDITS:
- Structural: addSheet, deleteSheet, renameSheet, duplicateSheet, moveSheet
  moveSheet REQUIRED fields: action, sheetName, newIndex (1-based integer; 1 = first tab, 2 = second, etc.)
  Example: {"action": "moveSheet", "sheetName": "Executive_Summary", "newIndex": 1}
- Sheet: freeze, clear, clearContent, clearCharts, clearConditionalFormat
  freeze REQUIRED fields: action, sheetName, and at least one of frozenRows or frozenColumns (integer, 0 to unfreeze).
  Example: {"action": "freeze", "sheetName": "Sheet1", "frozenRows": 1, "frozenColumns": 0}
  clearConditionalFormat (sheet-level, no range): clears ALL conditional formatting rules on the sheet.
  Example: {"action": "clearConditionalFormat", "sheetName": "Sheet1"}
  clearConditionalFormat (range-level): clears rules whose ranges intersect the given range. Include "range".
  Example: {"action": "clearConditionalFormat", "sheetName": "Sheet1", "range": "A2:E11"}
- Data: appendRow, appendRows
- Protection: protectRange
- Merge: merge, unmerge
- Resize: resizeColumn, resizeColumns, resizeRow, resizeRows, autoResizeColumns, autoResizeRows
- For autoResizeColumns and autoResizeRows, you may use either:
  - startColumn + numColumns / startRow + numRows
  - or a range like "A:H" or "2:10"
- Named ranges: addNamedRange, removeNamedRange
  addNamedRange REQUIRED fields: action, sheetName, range (A1 notation), namedRangeName
  Example: {"action": "addNamedRange", "sheetName": "Config", "range": "A1:A3", "namedRangeName": "StatusList"}
  removeNamedRange REQUIRED fields: action, sheetName, namedRangeName
  Example: {"action": "removeNamedRange", "sheetName": "Config", "namedRangeName": "StatusList"}
  Note: sheetName is required for routing only (must be a writable sheet in scope). Named ranges are matched by name across the entire workbook — the removal is not filtered to the given sheet.
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
  {"type":"dropdownFromNamedRange","namedRangeName":"StatusList"}
  {"type":"checkbox"}
  Note: Use "dropdownFromNamedRange" when a named range already exists in the workbook (e.g. created via addNamedRange). Use "dropdownFromRange" when referencing a raw sheet range directly.
  ORDERING: If you create a named range with addNamedRange AND reference it with dropdownFromNamedRange in the same response, the addNamedRange edit MUST appear first in the edits array — named ranges are registered in array order and are not available until their addNamedRange edit has been processed.
- Conditional formatting:
  Put it on a specific exact range and use one of these types:
  greaterThan, greaterThanOrEqualTo, lessThan, lessThanOrEqualTo, equalTo, notEqualTo,
  numberBetween, numberNotBetween, textEqualTo, textContains, textStartsWith, textEndsWith,
  empty, notEmpty, dateBefore, dateAfter, customFormula
  Example (value-based):
  {"sheetName":"Sheet1","range":"C2:C20","conditionalFormat":{"type":"greaterThan","value":"100","color":"#c6efce"},"replaceConditionalFormat":true}
  Example (customFormula — highlight entire row when column C > 320):
  {"sheetName":"Sheet1","range":"A2:E11","conditionalFormat":{"type":"customFormula","formula":"=$C2>320","color":"#FFFF99"},"replaceConditionalFormat":true}
  Note: For customFormula, use "formula" as the field name. Use absolute column + relative row (e.g. =$C2>320) so the rule evaluates correctly across the entire range.
- Images:
  {"action":"insertImage","sheetName":"Dashboard","imageUrl":"https://...","row":1,"column":1}
  IMPORTANT: imageUrl must be a direct, publicly accessible image URL (e.g. a raw GitHub image, Google-hosted image, or similar). Redirect-based services like via.placeholder.com or URL shorteners may fail. Use a reliable direct image URL such as https://dummyimage.com/150x150/000/fff.png or https://picsum.photos/150 instead.
- Notes/comments:
  Use "note" on a range edit, or action "comment" with note
  IMPORTANT: When the user asks you to add a note to a specific cell (e.g., "the Total Top Performers cell"), you MUST look up the exact A1 address of that cell from the WORKBOOK DATA IN SCOPE above before writing the edit. Do not guess row numbers.

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

      const finalDeletedSheets = dedupeArray(executionState.deletedSheets.concat(correctionState.deletedSheets));
      return JSON.stringify({
        message: (correctedResponseObj.message || "I corrected formula issues.") + "\n\n(Self-corrected a formula/data issue during execution.)",
        edits: correctedResponseObj.edits,
        newlyAddedSheets: finalNewSheets,
        deletedSheets: finalDeletedSheets,
        scopeChanges: finalScopeChanges,
        updatedScope: dedupeArray(
          correctionState.activeWritableSheets.filter(name => !finalDeletedSheets.includes(name))
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
        executionState.activeWritableSheets.filter(name => !executionState.deletedSheets.includes(name))
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
