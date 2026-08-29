# PowerPoint Gap Analysis & Local Implementation Roadmap

**AI Assistant for Microsoft Office add-in (v0.5.0) — what is thin, what can still be built locally, and what was deliberately rejected**

| | |
|---|---|
| **Generated** | 2026-08-28 |
| **Repository state** | branch `fix/adversarial-review-slice5-word-tools`, commit `c5ba65828c419142f923a9cff08b8ee5e55c5576` |
| **Companion report** | [PPT_Feature_Parity_Report.md](PPT_Feature_Parity_Report.md) |
| **Status note** | Point-in-time analysis report, **not** a planning document. `AI_Assistant_SSOT.md` remains the single source of truth (§0). Any of the "implementable" items below would need to go through SSOT §11 change control before implementation. |

---

## 1. Partial / thin spots found in the current implementation

1. **Structured rollback coverage is thin for PowerPoint.**
   `RollbackExecutor.cs:91-140` captures before-state only for `move_slide` (inverse coordinates) and `set_notes` (notes text). Every other PowerPoint mutation relies on app-level `CommandBars.ExecuteMso("Undo")`, which is outside the strict-LIFO, audited rollback guarantee that Phase C gives Excel.

2. **No PowerPoint plan template in Plan mode.**
   `Planner.cs` contains zero PowerPoint-specific templates (pattern search for `PowerPoint|ppt|slide|deck` returns nothing). Deck reorganization therefore flows through per-action approval cards, not the editable, reorderable Plan card that Word/Excel flows can use.

3. **No dedicated whole-deck translate.**
   The ribbon Translate submenu (9 languages) operates on selection/context. There is no slide-by-slide "translate this deck" loop, even though every primitive it needs (`GetPresentationText`, `set_shape_text`, `replace_text`) already exists and is risk-gated.

4. **Excel → slide chart/table is manual.**
   `.xlsx` attachment extraction (with real sheet-qualified provenance) and `powerpoint.add_table`/`add_chart` both exist, but there is no prompt/flow that wires attachment numbers into a slide table or chart automatically.

5. **SSOT errata — stale claim about `.pptx` extraction.**
   `AI_Assistant_SSOT.md` line 213 states that `.pptx` attachment extraction "still emits an ordinal `--- Slide N ---` (no per-slide title resolution)". **This is stale.** `AttachmentExtractor.cs:472-622` now resolves, per slide: correct slide order from `presentation.xml` (avoids the `slide10 < slide2` lexical sort), `[Title]` (title placeholder, deduped from body), `[Section]` (via `BuildSectionMap`), `[Layout]`, `[Speaker Notes]` (placeholder text filtered), plus a presentation-level `--- Sections ---` list. A §3 errata note should be added per the SSOT's own change-control rules; the extraction is now **more** informative than the `.docx`/`.xlsx` provenance row it was unfavorably compared against.

---

## 2. Leftovers you can implement — ranked (local, no cloud/enterprise)

| # | Feature | Effort | Hook points |
|---|---|---|---|
| 1 | **Whole-deck translate** — loop slides, propose per-slide `set_shape_text`/`replace_text` under one approval batch | Low | New "Translate deck" ribbon item (PowerPoint-visible); reuse `GetPresentationText` + existing risk-gated tools |
| 2 | **Expand PowerPoint rollback** — before-state capture for `set_shape_text`/`replace_text` (old shape text) and `set_font` (original font properties), mirroring the `move_slide`/`set_notes` pattern | Low–Med | `RollbackExecutor.CaptureBeforeState` + new inverse strategies; `RollbackExecutorTests` has the pattern to copy |
| 3 | **Deck consistency / audit chip** — extend the existing deterministic review brief into Finding cards: font-size outliers, bullet-level inconsistency, hidden slides, pictures missing alt text | Low | `GetPresentationReviewContext` already computes untitled/no-body/duplicate-title stats; EvidenceLevel Finding-card plumbing exists |
| 4 | **"Notes for all slides" one-shot** — chip that asks the model for speaker notes for every slide, emitted as multiple `set_notes` actions (each individually approved and rollbackable) | Low | QuickPrompt chip + prompt template; no new tools needed |
| 5 | **"Slides from attached deck" chip** — surface the rich `.pptx` extraction (titles/sections/notes) as "create slides from this deck" → `create_slide` actions | Low | QuickPrompt chip over `AttachmentExtractor.ExtractPptx` |
| 6 | **Excel data → slide table/chart** — "chart this Excel attachment on slide N": parse attached numbers → `add_table`/`add_chart` with bounded cells (mirror the Excel safety bounds) | Medium | New prompt + parser over xlsx extraction; `add_table` already accepts `List<List<string>>` data |
| 7 | **Wire cross-host plan into chat** — Excel analysis → Word report → PPT briefing. `CrossHostPlanCoordinator` is built and fully tested but deliberately **not wired into chat** (single-host today); the SSOT lists this as an open follow-up, not a drop | Medium | `AssistantSession.ProcessAssistantResponse` + `ChatSidebar` plan card; needs a host-scoped step gating UI |
| 8 | **PowerPoint plan template** for deck reorganization, so Plan mode produces an editable reorg plan (move/create/section steps) | Low | `Planner.cs`; all required tools already registered in `ToolRegistry` |
| 9 | **Alt-text audit / regenerate for existing pictures** — deterministic scan for `Picture` shapes without alt text; optionally suggest text via a vision model (BYOK custom endpoint; vision routing already exists), plus a small `powerpoint.set_alt_text` tool | Medium | New `ToolDefinition` + COM scan in `PowerPointController` |
| 10 | **Update the SSOT** — add the §3 errata for the stale `.pptx` extraction claim (see §1.5 above) | Trivial | `AI_Assistant_SSOT.md` |

**Best value first:** #2 (rollback breadth) and #7 (cross-host wiring) are the two most valuable — they close structural guarantees rather than adding surface area. #1, #3, #4, #5 are cheap wins that reuse tested infrastructure.

---

## 3. Deliberately NOT implementable (documented decisions, not gaps)

| Feature | Reason |
|---|---|
| Designer / auto-design polish, AI image generation, stock & brand imagery | Rejected in SSOT §8: paid image endpoints + content-policy surface; contradicts the deliberate "Insert image = a local file you approved" rule |
| Graph / Work IQ / org templates / brand kit / tenant licensing | Requires Microsoft cloud infrastructure and tenant licensing; contradicts Office 2010–2021 support and the BYOK-local design |
| Presenter Coach / rehearsal speech feedback | Cloud speech processing; out of scope |
| Meeting recap → deck, Outlook integration | Enterprise scope; the orphaned `OutlookController.cs` was deleted as dead code (SSOT D-4) |
| Web / stock search grounding | Deliberately deferred (post-Phase D, opt-in BYOK client-side search); documented in SSOT, not silently dropped |

---

## 4. Bottom line

All local Copilot-equivalent PowerPoint features are implemented and verified (40/40 test suites passing). The remaining local work is the ten ranked items in §2 — four of which (#1, #3, #4, #5) are low-effort compositions of already-tested parts, while the two structurally valuable ones (#2 rollback breadth, #7 cross-host wiring) close guarantees rather than add features.
