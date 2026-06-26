/*
## 1. System Architecture & Lifecycle

The application operates on a stateless client-server model bridging the Google Sheets UI, the Google Apps Script (GAS) runtime environment, and an external LLM API (configured via Script Properties).

### The Request Lifecycle

1. **Client Trigger:** User types a message in `Sidebar.html` and hits send.
2. **State Sync:** Client calls `processChat` via `google.script.run`, passing the message history and currently selected sheet scope.
3. **Context Hydration:** GAS reads the live spreadsheet data for the selected scope to build the system prompt.
4. **LLM Inference:** GAS makes a synchronous POST request to the configured LLM API endpoint.
5. **JSON Extraction:** GAS parses the raw LLM response to extract the structured `edits` array.
6. **Validation Pre-flight:** The execution engine verifies all requested edits against the allowed sheet scope and formula safety rules.
7. **Execution:** Validated edits are applied to the spreadsheet via the `SpreadsheetApp` API.
8. **Auto-Correction (Agentic Loop):** GAS inspects newly written formulas. If it detects Google Sheets errors (e.g., `#REF!`, `#VALUE!`), it automatically triggers a hidden follow-up prompt to the LLM to fix the error.
9. **Client Update:** The final state, UI HTML, and execution summaries are returned to the sidebar.

---

## 2. Client-Side Implementation (`Sidebar.html`)

The frontend is a vanilla HTML/JS/CSS application designed to run entirely within the isolated `HtmlService` iframe.

### State Management

The frontend maintains three primary state variables:

* `allSheetNames` (Array): Every tab currently existing in the workbook.
* `selectedSheets` (Array): The active "scope" the AI is allowed to read/write.
* `sessionHistory` (Array): The chat log containing `role` (user/assistant) and `content`.

### Key Functions

* **`initializeSidebar()`:** Fires on load. Fetches persistent user settings from the backend to restore previous chat sessions and selected sheet scopes.
* **`sendMessage()`:** Handles UI locking (disabling buttons), appending optimistic UI messages ("Thinking..."), and invoking the backend `processChat()`.
* **`applyScopeChanges()`:** Allows the user to add/remove sheets from the AI's allowed scope dynamically without resetting the conversational context.

---

## 3. Server-Side Persistence & Initialization

The system uses Google's `PropertiesService.getUserProperties()` to store data across sessions.

* **`AI_SELECTED_SHEETS`:** Stringified JSON array of allowed sheet names.
* **`AI_SESSION_HISTORY`:** Stringified JSON array of the conversational memory.

*Developer Note:* Before returning state to the frontend, `getSidebarState()` always cross-references the saved `selectedSheets` against `spreadsheet.getSheets()` to automatically purge deleted tabs from the scope.

---

## 4. LLM API Orchestration (`fetchAiResponse` & `parseAiResponse`)

### API Connectivity

All three values below are read at runtime from Google Apps Script **Script Properties** (see **Setup** section at the bottom for how to add them):

* **`AI_API_KEY`:** Your API key, passed as a Bearer token in the `Authorization` header.
* **`AI_API_ENDPOINT`:** The full chat completions URL (e.g., `https://api.openai.com/v1/chat/completions`). Any OpenAI-compatible endpoint works.
* **`AI_MODEL`:** The model identifier string (e.g., `gpt-4o`, `claude-sonnet-4-6`).
* **Parameters:** `temperature: 0.1` (Low variance to ensure strict adherence to the JSON schema).

### Strict JSON Parsing

Because LLMs occasionally hallucinate markdown formatting (e.g., `json ... `) or append conversational text, `parseAiResponse()` does not rely on direct `JSON.parse()`.

1. It searches the raw string for the first `{` and the last `}`.
2. It extracts the substring and attempts parsing.
3. If parsing fails, it safely catches the error and returns a predefined error schema to the client rather than crashing the script.

---

## 5. Context Aggregation (The "Workbook RAG" Engine)

To provide the LLM with accurate data without exceeding token limits or memory limits, the script intelligently summarizes sheet content.

### `getActualDataBounds(sheet)`

Standard `getLastRow()` can be fooled by empty formatted cells. This function iterates through `getDisplayValues()` to find the *true* bottom-right boundary of actual data.

### `buildSheetContextString(sheet)`

This compiles the raw text injected into the LLM system prompt. It includes:

* **Metadata:** Dimensions, frozen rows, chart counts, and empty states.
* **Named Ranges:** Injects up to 20 relevant named ranges mapped to their A1 notation.
* **Data Truncation Strategy:** * Fetches the first 40 rows (Head).
* If the sheet is large (>60 rows), it skips the middle and explicitly tells the LLM: `[... X middle rows hidden ...]`.
* Fetches the last 10 rows (Tail).
* *Why?* This gives the LLM the header structure and the latest appended data while saving tokens.



---

## 6. The Validation Gatekeeper (`validateEdits`)

Before a single cell is modified, the JSON payload must pass strict validation to prevent malicious or hallucinated edits.

### Security Rules Enforced:

1. **Scope Containment:** If an edit targets a `sheetName` that is not in the `allowedWritableSheets` array, it is rejected.
2. **Creation Tracking:** If an action is `addSheet`, that sheet is immediately pushed to a temporary allowed list, so subsequent edits in the same JSON payload can write to the newly created sheet.
3. **Formula Sanitization (`sanitizeFormulaReferences`):** * Extracts all sheet references via Regex (`/'([^']+)'!/g` and `/\b([A-Za-z0-9_]+)!/g`).
* If a formula references a pre-existing sheet *outside* the user-approved scope, the entire edit batch is flagged and rejected.



---

## 7. Execution Engine: Action Handlers

The `applyEditsToSheet` function iterates through the validated JSON and routes each edit object to specialized handlers.

### A. Structure & Layout (`handleSheetStructureAction`)

* **Actions:** `addSheet`, `deleteSheet`, `renameSheet`, `duplicateSheet`, `moveSheet`.
* *Edge Case Handling:* Automatically resolves naming collisions. If `addSheet` is called for a name that already exists, it gracefully skips creation and simply brings the existing sheet into scope.

### B. Formatting & Protection (`handleFreezeClearAppendProtectActions`)

* **Actions:** `freeze`, `clear`, `clearContent`, `appendRow`, `appendRows`, `protectRange`.
* *Security Note:* `protectRange` automatically removes existing editors and adds the current `Session.getEffectiveUser()` unless instructed otherwise.

### C. Cell Value Injection (`handleCellRangeActions`)

This is the most complex handler, managing cell values, grids, and visual formatting.

* **Conflict Resolution (The Shatter Protocol):** Google Sheets API throws fatal errors if you try to inject an array into a grid that overlaps with merged cells. If a `values2D` array is detected, `cleanupOverlappingMergedRanges` calculates the mathmatical grid of the incoming data and proactively `breakApart()` any merged cells in its path.
* **Formula vs. String Separation:** Google Sheets requires formulas to be injected via `setFormulas()`, while standard data uses `setValues()`. If you mix them, data can corrupt.
* The engine separates the `values2D` grid into a plain-text grid (written atomically) and iterates over the remaining cells to write formulas individually.


* **Rich Formatting:** Maps LLM JSON keys (e.g., `backgroundColor`, `fontWeight`, `wrapStrategy`) directly to `SpreadsheetApp.Range` methods.

### D. Chart Generation (`handleChartActions`)

* **Fallback Anchors:** If the LLM requests a chart but forgets to specify where to place it (`range`), the engine calculates `getLastRow() + 5` and drops the chart at the bottom of the data set.
* **Dynamic Parsing:** Handles chart ranges whether the LLM provides them as an array `["A1:A10", "B1:B10"]` or a comma-separated string.

---

## 8. The Agentic Auto-Correction Loop

This is a critical reliability feature implemented in `processChat`.

1. **Queueing:** During execution, any time a formula is injected into a cell, that A1 notation is pushed to `state.formulaRangesToInspect`.
2. **Inspection (`collectFormulaIssues`):** After `SpreadsheetApp.flush()` applies all edits, the engine uses `getDisplayValues()` on the queued cells to check for `#ERROR`, `#REF!`, `#N/A`, `#VALUE!`, `#NAME?`, etc.
3. **Self-Correction:** If an issue is found, the system *pauses returning the response to the user*. Instead, it appends a hidden `user` message to the payload: *"Your last edit caused formula/data issues. Detected issues: [Error Details]. Fix these issues..."*
4. **Re-Execution:** It makes a second, silent call to the LLM API, validates the new response, applies the fixes, and only then returns the combined result to the user UI, logging the correction in the `machineSummary`.

---

## Setup: Configuring Script Properties in Google Apps Script

The script reads three values from **Script Properties** at runtime. These are never committed to source control — they live only inside your Apps Script project.

### Step-by-step

1. Open your Google Sheet and go to **Extensions → Apps Script**.
2. In the Apps Script editor, click the **⚙ Project Settings** gear icon in the left sidebar.
3. Scroll down to the **Script Properties** section and click **Add script property** for each of the three keys below.

### Required properties

| Property key | Description | Example value |
|---|---|---|
| `AI_API_KEY` | Your API key for the LLM provider. Sent as a `Bearer` token. | `sk-...` |
| `AI_API_ENDPOINT` | Full URL of the OpenAI-compatible chat completions endpoint. | `https://api.openai.com/v1/chat/completions` |
| `AI_MODEL` | The model identifier string accepted by the endpoint. | `gpt-4o` or `claude-sonnet-4-6` |

4. Click **Save script properties** after adding all three.
5. Reload the Google Sheet and open **🤖 AI Assistant → Open Chat** from the menu bar.

### Notes

* The script will throw a descriptive error in the chat if any of the three properties is missing, so you can tell at a glance which one to add.
* `AI_API_ENDPOINT` must point to an endpoint that accepts the OpenAI `POST /v1/chat/completions` message schema (i.e., a `messages` array with `role`/`content` pairs). Most major providers (OpenAI, Anthropic via proxy, Azure OpenAI, etc.) support this format.
* `temperature` is fixed at `0.1` in code to minimise variance in the structured JSON the model must return. You can change this in `fetchAiResponse` if needed.
*/