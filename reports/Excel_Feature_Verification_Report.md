# Excel Feature Verification Report — AI Assistant for Microsoft Office v0.5.0

**Date:** 2026-08-29
**Branch:** `fix/adversarial-review-slice5-word-tools`
**Scope:** Consumer **Microsoft 365 Copilot in Excel** features vs. this project. Cloud/enterprise features (Graph, OneDrive/SharePoint grounding, Work IQ, agents, Designer, tenant admin) are **excluded** per request.
**Method:** Source inspection only (no live build/COM run). Copilot feature set from documented consumer capabilities (Microsoft support pages were unreachable; project side is 100% verified from source).
**Read with:** [`Excel_Feature_Reverification_Report.md`](Excel_Feature_Reverification_Report.md) — corrects one count (48 types, not 47) and re-confirms every claim with file:line evidence.

---

## 1. Implemented Excel surface (verified from source)

**48 registered `excel.*` tools** in `src/Core/Actions/ToolRegistry.cs` (lines 224–611), each wired to a real `HostOperationResult` handler in `src/Hosts/ExcelController.cs` (~3,200 lines), plus context, safety, and UI layers. The prompt's action-type allow-list is **generated** from the registry (`ToolRegistry.FormatActionTypesList`, `ToolRegistry.cs:194`) and byte-pinned by the golden-master fixture (`tests/Fixtures/golden_master_baseline.txt`) — prompt and registry are in sync.

| Category | Tools |
|---|---|
| Data entry | `write_formula`, `write_value`, `fill_down`, `table` (markdown/CSV writes, bounded-range clipping), `create_table` (ListObject) |
| Analysis | `analyze_range` (local read-only per-column stats: min/max/mean, distributions, trends, outliers + chart/pivot suggestions), `get_formula_details` (formula + precedents/dependents), `add_analysis_column` (sentiment/classify/topic/summarize column structure) |
| Presentation | `conditional_format` (color_scale, data_bar, greater/less/equal, top_n, duplicates, contains, between, icon_set), `sort`, `filter`, `format_cells`, `merge_cells`, `autofit_columns`, `freeze_panes`, `apply_theme`, `create_chart` (column/bar/line/pie), `create_pivot_table`, `named_range`, `add_summary_row` |
| Data hygiene | `remove_duplicates`, `find_replace`, `set_case`, `trim_range`, `normalize_whitespace`, `text_to_columns`, `data_validation` (list/whole/decimal/date/between/custom) |
| Workbook structure | `add/rename/delete/duplicate_worksheet`, `set_tab_color`, `insert/delete_rows`, `insert/delete_columns`, `hide/unhide_rows`, `hide/unhide_columns` |
| Advanced / local | `write_python` (single-line `=PY()`), `import_worksheet` (local .xlsx via OpenXML, 20 MB cap, no cloud), `create_shape` / `update_shape` (incl. alt text), `set/get/clear_workbook_rules` (per-workbook local personalization via `Core/WorkbookRulesStore.cs`) |
| Context & safety | `GetWorksheetSnapshot(70×26)` with formula annotations, `GetSelectedRangeValues`, `GetContextReadout`, `NavigateToCell` (clickable `Sheet1!B7` citations), `ExcelChangeHighlighter` (Copilot-style green tab + grid + `clear_highlights`), risk levels 0–3, approval flow, `RollbackExecutor` before-state capture, DPAPI audit log, Chat/Plan/Edit mode gating |

> ⚠️ **Docs are stale against the code.** `README.md:20` says "13 confirm-before-apply action types"; `AI_Assistant_SSOT.md` §3 said 13 (updated to 38 during this review — still short of the 9 analysis-slice types); `src/Help/UserManual.html:149` lists only the original 11 capabilities by name. True count: **48**.

## 2. Feature-by-feature comparison vs. consumer Copilot in Excel

| Copilot in Excel feature (consumer) | This project | Verdict |
|---|---|---|
| Chat with your data (side panel, follow-ups) | ChatSidebar + worksheet snapshot/selection context, scope selector, streaming, clickable citations | ✅ Covered |
| Write formulas from a description | `excel.write_formula` + `fill_down` | ✅ Covered |
| Explain formula | `explain_formula` skill + `get_formula_details` (reads precedents/dependents — Copilot doesn't) | ✅ Covered, deeper |
| Add calculated/analysis column | `write_formula`+`fill_down`; `add_analysis_column` creates column structure for AI follow-up | ✅ Covered |
| Highlight / conditional formatting | 10 rule families incl. icon sets and data bars | ✅ Covered (fixed highlight color — see gaps) |
| Sort / filter | `excel.sort`, `excel.filter` | ✅ Covered |
| Insights (trends, outliers, suggested visuals) | `analyze_range` local stats + AI chat analysis + dashboard/Pareto skills | ✅ Covered |
| Create charts | 4 types: column, bar, line, pie | ⚠️ Partial |
| Create PivotTables | Real PivotCache pivot at a bounded destination — but **no field configuration** (`rows:/vals:` silently ignored) | ⚠️ **Real gap** |
| Summarize data | Summarize quick-prompt + skills + Insert into sheet | ✅ Covered |
| Multiple formula suggestions (one-click candidates) | Chat can list variants; no dedicated candidate-card UI | ⚠️ Partial |
| Data cleaning | 7 dedicated tools + "Clean Data" skill | ✅ Covered |
| Formatting & themes | `format_cells`, `apply_theme`, merge, freeze, autofit | ✅ Covered |
| Sheet management | add/rename/delete/duplicate/tab color | ✅ Covered |
| Python in Excel | `write_python` — single-line `=PY()`, single cell only | ⚠️ Partial |
| Cross-sheet / whole-workbook questions | Snapshot is **active-sheet only** (documented SSOT §2.6 gap); actions do accept `Sheet1!B7` targets | ⚠️ Partial |
| Table (ListObject) awareness in context | Can create/format tables; context doesn't enumerate existing tables | ⚠️ Partial |
| Import data from a file | `import_worksheet` (local, privacy-preserving — Copilot uses OneDrive) | ✅ Covered |
| Image of a table → data | Vision routing + write actions (user attaches) | ✅ Covered |
| Shapes/annotations | `create_shape`/`update_shape` — Copilot doesn't offer this | ✅ Beyond Copilot |
| Per-workbook memory/preferences | `set/get/clear_workbook_rules` (local JSON + optional `.Rules` sheet) | ✅ Equivalent, no cloud |
| Change highlighting + undo | Green tab/grid highlighter, undo, strict-LIFO rollback, audit log | ✅ **Stronger than Copilot** |
| Approval / verification / plan-before-run | Risk 0–3, preview cards, editable Plan mode, `ActionVerifier` | ✅ **Stronger than Copilot** |
| Export summary to Word | `CrossHostPlanCoordinator` exists but **not wired into chat** (tracked D3 follow-up) | ❌ Missing |
| Sparklines | Not present anywhere | ❌ Missing (Copilot lacks this too) |
| Multi-file/OneDrive/Graph grounding, agents, tenant admin, Designer | Deliberately out of scope (SSOT §1 "Out of scope", §7 rejected row) | ✗ By design |

## 3. Leftovers that CAN still be implemented (ranked by value/effort)

1. **PivotTable field configuration** — accept `rows:/values:/columns:/filter:` in `ExecuteCreatePivotTable` (`ExcelController.cs:1103`) and set `PivotField.Orientation` after `CreatePivotTable`. Highest payoff — the one place Copilot is genuinely ahead.
2. **Workbook-wide context** — compact multi-sheet overview appended to `GetWorksheetSnapshot` (closes SSOT §2.6's own "one real gap"); token-bound via `TokenCounter`.
3. **Table enumeration in context** — list `Workbook.ListObjects` (name + headers) so "add a column to the Sales table" works.
4. **Excel summary → Word export** — wire the already-built `CrossHostPlanCoordinator` into chat (documented open item).
5. **More chart types** — scatter, area, doughnut, stacked/combo: enum additions in `CreateChart` (`ExcelController.cs:1089`).
6. **Dedicated "Analyze" Excel chip** — `QuickPromptRegistry` entry firing `analyze_range` + chart/pivot batch (SSOT backlog #4, "a chip, not a project").
7. **Multi-line `=PY()`** — relax the single-line constraint in `ExecuteWritePython` (`ExcelController.cs:2136`).
8. **Sparklines** — new tool via `Range.SparklineGroups.Add` (cheap, unique vs. Copilot).
9. **Formula-variant cards** — several candidate formulas with per-variant Apply (Copilot's formula-suggestions UX).
10. **Configurable CF highlight color** — fill is hardcoded (`13551615`); expose a fixed palette like `apply_theme`.
11. **Already-tracked open items touching Excel:** `@` mention local files (backlog #3), WorkSession persistence (D4), "Planning" status (A7).

## 4. Deliberately excluded (cloud/enterprise — matches your request)

Graph grounding, Work IQ, Designer, agents, multi-file OneDrive/SharePoint analysis, web-search grounding (deferred post-Phase D), text-to-image (rejected), tenant admin/compliance. Formally recorded in SSOT §1 and §7; also stated to users in `UserManual.html` "What is NOT included".

## 5. Bottom line

Essentially the entire consumer Copilot-in-Excel surface is implemented — 48 tools, all with handlers, safety bounds, rollback, and audit — plus several capabilities Copilot doesn't have (shapes, workbook rules, local import, editability guards, rollback executor). Functional holes vs. consumer Copilot: **PivotTable field setup, multi-sheet context, multi-line Python, chart variety, summary→Word export**. Cloud/enterprise items are already formally excluded by design.
