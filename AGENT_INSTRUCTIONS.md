# Agent Instructions — Workbook-Aware Context + Cross-Sheet Commands

## What was fixed

### Issue 1: Context button was NOT workbook-wide
**Before:** `getContextForAI()` only read the active sheet's selected range or used range. The AI had no idea other sheets existed or what data was on them.

**After:** `getContextForAI()` now:
1. Lists ALL sheets with their dimensions (rows × cols, used range address, hidden status)
2. Lists named ranges
3. Then includes the active sheet's data as before
4. Tells the AI to use `get_sheet_data` to read other sheets

### Issue 2: "Build dashboard on new sheet from sheet1" didn't work
**Root cause:** All Excel commands (`write_range`, `create_chart`, `format_range`, `get_range_csv`, etc.) operated ONLY on the active sheet. There was no way to:
- Read data from a specific sheet
- Write to a specific sheet
- Activate a different sheet

**After:** Added sheet-awareness to all key commands:
- **New command: `get_sheet_data`** — reads all data from a named sheet
- **New command: `activate_sheet`** — switches the active sheet
- **New `sheet` parameter** on: `write_cell`, `write_formula`, `write_range`, `create_chart`, `clear_range`, `format_range`, `get_range_csv`
- **New "WORKFLOW FOR BUILDING DASHBOARDS" section** in the system prompt with step-by-step instructions

### Issue 3: VBA macros & Power Query (assessed — not implemented)
**VBA Macros:** NOT feasible via Office.js. The Office.js API does not provide access to the VBA object model. VBA macros can only be created/edited through the VBA editor (Alt+F11), which is not accessible from a taskpane add-in. The `eval_js` escape hatch can run Office.js code but cannot create VBA macros.

**Power Query:** Technically possible via `ExcelApi 1.9+` (`workbook.queries`), but requires significant new infrastructure: M query syntax validation, query refresh management, and error handling for malformed queries. Too complex for this iteration. The AI can still GENERATE Power Query M code as text in its response (the user copies it manually), but cannot inject it programmatically.

## Files changed (4)

1. **`src/taskpane/services/office/ExcelService.ts`**
   - `getContextForAI()` — now lists all sheets + named ranges before active sheet data
   - `writeToRange()` — added optional `sheetName` parameter
   - `insertFormula()` — added optional `sheetName` parameter
   - `createChart()` — added optional `sheetName` parameter
   - `clearRange()` — added optional `sheetName` parameter
   - `formatRange()` — added optional `sheet` field in `opts` parameter
   - `getRangeAsCsv()` — added optional `sheetName` parameter
   - **New method: `activateSheet(name)`** — activates a worksheet by name

2. **`src/taskpane/hooks/useChat.ts`**
   - `executeExcelCommands` — passes `cmd.sheet` to all sheet-aware methods
   - **New command: `activate_sheet`** — calls `ExcelService.activateSheet(cmd.name)`
   - **New command: `get_sheet_data`** — calls `ExcelService.getSheetData(cmd.sheet)` and returns formatted data

3. **`src/taskpane/services/prompts.ts`**
   - Added "WORKBOOK AWARENESS" section explaining cross-sheet operations
   - Added `get_sheet_data` and `activate_sheet` to the command list
   - Added `sheet` parameter documentation to all relevant commands
   - Added "WORKFLOW FOR BUILDING DASHBOARDS ON NEW SHEETS" with 5-step example
   - Updated chaining examples to include cross-sheet workflows

## Verification

```
tsc --noEmit              → 0 errors
eslint --max-warnings 0   → 0 errors, 0 warnings
webpack --mode production  → succeeds, 962 KB
```

## How to deploy

```bash
unzip workbook-aware.zip -d YOUR_REPO
cd YOUR_REPO
npm install
git add -A
git commit -m "feat: workbook-wide context + cross-sheet commands

- Context now lists ALL sheets with dimensions, not just active sheet
- New commands: get_sheet_data, activate_sheet
- Sheet parameter added to: write_cell, write_formula, write_range,
  create_chart, clear_range, format_range, get_range_csv
- System prompt: WORKBOOK AWARENESS section + dashboard workflow guide
- VBA macros: not feasible via Office.js (documented)
- Power Query: too complex for this iteration (documented)"
git push origin main
```
