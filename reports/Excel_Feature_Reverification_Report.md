# Excel Feature Re-verification Report — AI Assistant for Microsoft Office v0.5.0

**Date:** 2026-08-29
**Branch:** `fix/adversarial-review-slice5-word-tools`
**Purpose:** Second pass over every claim in [`Excel_Feature_Verification_Report.md`](Excel_Feature_Verification_Report.md), with file:line evidence. Supersedes that report wherever the two disagree.

---

## A. What changed / was corrected during re-verification

1. **The working tree changed mid-review.** First shell listing showed `ExcelController.cs` at ~2,570 lines (`InsertText` at 2352); direct reads now show **3,282 lines** (`InsertText` at 3015, file ends 3282) with the full "Analysis / Import / Shapes / Rules" block present. `.git/HEAD` still resolves to `c5ba65828c419142f923a9cff08b8ee5e55c5576` — the same commit hash as at review start — so the analysis slice appears to be **uncommitted working-tree work** (shell capture broke during the session, so `git status` could not be run; confirm manually).
2. **`AI_Assistant_SSOT.md` was edited during the review:** line 471 changed from *"13 confirm-before-apply action types"* to *"38 … was 13, see commit history"*. Still short of the true count (the 9 analysis-slice types are missing from that figure).
3. **Correction to the first report's own number:** the prompt allow-list contains **48 action types mapping 1:1 to 48 registered `excel.*` tools** — the earlier "47" was an arithmetic miscount.

## B. Claim-by-claim re-verification (evidence)

| # | Claim | Re-verified | Evidence |
|---|---|---|---|
| 1 | 48 Excel tools registered, all with handlers | ✅ | Registry `excel.*` registrations `ToolRegistry.cs:224–611`; handlers read directly: `ExecuteAnalyzeRange` (2279), `ExecuteGetFormulaDetails` (~2390–2475), `ExecuteAddAnalysisColumn` (2478), `ExecuteImportWorksheet` (2575), `ExecuteCreateShape` (2732), `ExecuteUpdateShape` (2795), `ExecuteSetWorkbookRule` (2853), `ExecuteGetWorkbookRules` (2890), `ExecuteClearWorkbookRules` (2916); file ends cleanly at 3282 |
| 2 | Prompt advertises exactly the registered types | ✅ | `FormatActionTypesList("Excel")` (`ToolRegistry.cs:194`) lists all 48; `tests/Fixtures/golden_master_baseline.txt` pins the identical string — no prompt/registry drift |
| 3 | `WorkbookRulesStore` exists and compiles in | ✅ | `src/Core/WorkbookRulesStore.cs` (194 lines; JSON store + optional hidden `.Rules` sheet reader); **csproj line 206 includes it** — no "D-4 missing-csproj-entry" trap |
| 4 | PivotTable field configuration gap | ✅ **Still open** | `CreatePivotTable` (`ExcelController.cs:1103–1124`) reads only `destination:` and `name:`; `rows:/vals:/columns:` options the model may send are silently ignored — an empty pivot shell is created |
| 5 | Context is active-sheet only | ✅ **Still open** | `GetWorksheetSnapshot` (`ExcelController.cs:145–154`) reads `app.ActiveSheet` only; no workbook-wide enumeration |
| 6 | `=PY()` single-line only | ✅ **Still open** | `ExecuteWritePython` (2141–2142) explicitly rejects newlines; 8000-char cap; Formula2 probe; 1×1 target enforced |
| 7 | Only 4 chart types | ✅ | `CreateChart` (~1093–1099): line→4, pie→5, bar→57, else column (51) |
| 8 | Cross-host plan not wired into chat | ✅ | SSOT 852–854 states it; only test references exist (`CrossHostPlanCoordinatorTests`) |
| 9 | No sparklines | ✅ (action surface) | No sparkline type in the 48-type allow-list / registry |
| 10 | Docs stale | ✅ nuanced | `README.md:20` still "13"; `AI_Assistant_SSOT.md:471` now "38" (short by the 9 analysis types); `src/Help/UserManual.html:149` lists only the original 11 Excel capabilities |
| 11 | Test coverage of the 9 analysis tools | ⚠️ **None found** | `ToolRegistryTests` asserts only pre-slice tool names; no suite exercises the analysis/import/shape/rules handlers (not even null-controller failure paths à la `HostOperationResultTests`). Only guard: golden-master prompt hash |
| 12 | Cloud/enterprise exclusions | ✅ | `UserManual.html` "What is NOT included" (156–161): no Graph/SharePoint/OneDrive/tenant search, no telemetry |

## C. Net result after re-verification

- **Every material finding of the first report stands**, with the count corrected to **48/48** and the SSOT finding updated (38, not 13).
- **The top functional hole vs. consumer Copilot is unchanged and higher-stakes**: pivot field configuration — the prompt advertises the tool, so models will emit field options that are silently dropped.
- Docs are stale in **three places at three different counts** (README 13 / SSOT 38 / UserManual 11-by-name) against a true 48.

## D. New risks flagged by re-verification

1. **Uncommitted analysis slice** — build + test + commit before anything else (HEAD hash unchanged while files demonstrably newer).
2. **No headless tests for the 9 new tools** (`analyze_range`, `get_formula_details`, `add_analysis_column`, `import_worksheet`, `create_shape`, `update_shape`, `set_workbook_rule`, `get_workbook_rules`, `clear_workbook_rules`) — at minimum add null-controller failure-path tests per SSOT §8.
3. **`<LangVersion>5</LangVersion>`** (`src/MSOfficeAIAssistant.csproj:17`) — new code must stay C# 5-compatible (current slice correctly does).
4. **Build/test not run this session** (plan mode + broken shell capture) — SSOT §8 gate: build x86+x64 Release, run `tests\MSOfficeAIAssistant.Tests.exe`, expect exit 0.

## E. Suggested plan (act mode)

1. **Commit checkpoint** — build x86+x64, run the test suite, then `git add` the analysis slice (`ExcelController.cs`, `ToolRegistry.cs`, `Core/WorkbookRulesStore.cs`, SSOT, golden-master fixture, `ToolRegistryTests.cs` if touched) and commit.
2. **Docs sync** — `README.md:20` → 48; SSOT §3 → 48; UserManual Excel bullet → add analysis, formula-details, import, shapes, workbook rules, Python.
3. **Pivot fields** (top Copilot-parity gap) — parse `rows:/values:/columns:/filter:` in `ExecuteCreatePivotTable`, set `PivotField.Orientation`, extend `HostOperationResultTests` + `ToolRegistryTests`, bump golden master if prompt text changes.
4. Then in order: multi-sheet context → tests for the 9 analysis tools → chart types → multi-line `=PY()` → Excel "Analyze" chip.
