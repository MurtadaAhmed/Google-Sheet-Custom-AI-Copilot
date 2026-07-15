# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Runtime Environment

This is a **Google Apps Script (GAS) add-on** — not a Node.js project. There is no `package.json`, no `npm`, no local test runner, and no build step. All `.js` files run inside Google's V8-based Apps Script runtime and are deployed via the `clasp` CLI.

**Deployment:** `clasp push` uploads all local `.js`/`.html` files to the Apps Script cloud project (script ID is in `.clasp.json`, which is gitignored). There is no compilation or transpilation.

## Configuration

All secrets live in Google Apps Script **Script Properties** — never in code or files:

| Property | Purpose |
|---|---|
| `AI_API_KEY` | Bearer token for the LLM API |
| `AI_API_ENDPOINT` | Full chat completions URL (any OpenAI-compatible endpoint) |
| `AI_MODEL` | Model identifier string |

Session state is stored in `PropertiesService.getUserProperties()` under `AI_SELECTED_SHEETS` and `AI_SESSION_HISTORY` (stringified JSON).

## Architecture

All files live at the repository root. The server-side is split into 8 modules; there is one HTML frontend.

### Request Lifecycle

1. **`aiAssistant_Main.js`** — `onOpen()` creates the "Agentic AI" menu; `showSidebar()` serves `Sidebar.html` as a 360px iframe. The central `processChat(sessionHistory, selectedSheets)` function orchestrates every step.

2. **`aiAssistant_State.js`** — Sidebar calls `getSidebarState()` on load via `google.script.run`. Persists selected sheets and session history across page loads.

3. **`aiAssistant_Context.js`** — `buildContextForSheets()` reads each in-scope sheet, computes true data bounds (ignoring blank-but-formatted cells), and emits a structured text block (metadata + first 40 rows + last 10 rows, with middle truncated). This becomes the LLM system prompt context.

4. **`aiAssistant_ApiClient.js`** — `fetchAiResponse()` posts to the configured OpenAI-compatible endpoint at `temperature: 0.1`. `parseAiResponse()` extracts the JSON object from raw text (handles models that wrap JSON in markdown fences by finding the outermost `{...}`).

5. **`aiAssistant_Validation.js`** — `validateEdits()` pre-flight checks the entire edit array before any sheet modification. Key checks: formula references must stay within the selected scope; newly created sheets (within the same response) are tracked so subsequent edits targeting them pass validation.

6. **`aiAssistant_Execution.js`** — `applyEditsToSheet()` dispatches each edit to one of four specialized handlers in order. After `SpreadsheetApp.flush()`, runs `collectFormulaIssues()` to detect error values (`#REF!`, `#NAME?`, etc.) and trigger an agentic correction loop (one extra API call with a hidden correction prompt).

7. **`aiAssistant_Handlers.js`** — Four handlers called by the execution engine:
   - `handleSheetStructureAction` — add/delete/rename/duplicate/move sheets
   - `handleFreezeClearAppendProtectActions` — freeze, clear, appendRow/s, protectRange
   - `handleMergeResizeNamedRangeActions` — merge/unmerge, resize, named ranges
   - `handleChartActions` — clearCharts, addChart (PIE/BAR/LINE/AREA/SCATTER/COMBO/TABLE/COLUMN)
   - `handleCellRangeActions` (fallback) — value/formula writes, formatting, validation, images, comments

8. **`aiAssistant_Formatting.js`** — `setRichFormatting()`, `applyDataValidation()`, and the conditional formatting builder.

9. **`aiAssistant_Utils.js`** — Pure helpers: `dedupeArray`, `clamp`, `overlapsRange`, `cleanupOverlappingMergedRanges`, `getSafeTargetRange`.

10. **`Sidebar.html`** — Complete frontend: HTML + CSS + vanilla JS. Communicates with server via `google.script.run`. Manages scope selection, session history rendering, and displays AI responses with status divs (scope changes, warnings, token usage).

### Critical Implementation Details

**Shatter Protocol** (`aiAssistant_Utils.js`): Before writing a `values2D` grid, `cleanupOverlappingMergedRanges()` breaks apart any merged cells in the target range. This prevents the fatal "Cannot overwrite part of a merged cell" Apps Script error.

**Formula/value split write** (`aiAssistant_Handlers.js`): When a grid contains a mix of formula strings and plain values, two passes are made — `setValues()` for plain cells, then `setFormula()` per formula cell — because `setFormulas()` requires all strings to be formulas.

**Agentic correction loop** (`aiAssistant_Main.js`): After applying edits, `collectFormulaIssues()` reads `getDisplayValues()` on all formula cells. If errors are found, the context is rebuilt from the modified sheets, a correction prompt is appended to the history, and a second API call is made. The user sees "(Self-corrected a formula/data issue during execution.)" in the response.

## Testing

There is no automated test framework. `Passed Tests.md` is the test record — it documents 12 manually executed scenarios. When making changes, verify against the scenarios in that file by deploying via `clasp push` and testing in a real Google Sheet.

## Key Files

| File | What to read it for |
|---|---|
| `Technical Documentation.js` | Developer reference doc (stored as a JS comment block, not executed) |
| `potential_changes.md` | Detailed hardening plan (not yet applied) — proposed features include LockService, audit logging, dry-run mode, destructive action gating, and a multi-attempt correction retry loop |
| `appsscript.json` | Manifest: runtime (V8), timezone (Europe/Sofia), logging (STACKDRIVER), OAuth scopes |
