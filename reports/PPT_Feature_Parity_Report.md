# PowerPoint Feature Parity Report

**Microsoft 365 Copilot (PowerPoint) vs. AI Assistant for Microsoft Office add-in**

| | |
|---|---|
| **Generated** | 2026-08-28 |
| **Product version** | 0.5.0 |
| **Repository state** | branch `fix/adversarial-review-slice5-word-tools`, commit `c5ba65828c419142f923a9cff08b8ee5e55c5576` |
| **Scope** | Local, non-cloud, non-enterprise Copilot-in-PowerPoint features only. Cloud/tenant/enterprise features are listed in the companion gap report for completeness, marked not-implementable. |
| **Companion report** | [PPT_Gap_Analysis_Roadmap.md](PPT_Gap_Analysis_Roadmap.md) |
| **Status note** | Point-in-time analysis report, **not** a planning document. `AI_Assistant_SSOT.md` remains the single source of truth for the product (§0 of the SSOT). Source code outranks any claim made here. |

---

## 1. Verification evidence

### 1.1 Sources examined

- `src\Hosts\PowerPointController.cs` (~2,100 lines — full public surface inventoried)
- `src\Core\PowerPointActionParser.cs` (action model + `ParseStructuredActions` + `ParseSlideData`)
- `src\Core\Actions\ToolRegistry.cs` (18 registered PowerPoint tools, risk levels, rollback flags)
- `src\Core\Actions\ActionExtractor.cs` (unified `<office_actions>` + legacy `<powerpoint_actions>` extraction)
- `src\Core\Actions\RollbackExecutor.cs` (PowerPoint before-state capture + inverse execution)
- `src\Attachments\AttachmentExtractor.cs` (`ExtractPptx`, lines 472–622)
- `src\Core\PromptAssembler.cs` (host prompt + `BuildBriefingDeckPrompt`)
- `src\Core\QuickPrompts\QuickPromptRegistry.cs` (chips, `HostFilter = "PowerPoint"`)
- `src\Addin\Ribbon.xml` / `RibbonCallback.cs` (ribbon commands incl. `btnBuildDeck` → `OnBuildSlides`)
- `src\UI\ChatSidebar.xaml.cs` (deck application flow, action cards, insert-image button)
- `tests\` (suite list incl. `PowerPointActionParserTests`, `DocToDeckTests`, `RollbackExecutorTests`)
- `AI_Assistant_SSOT.md` (§2 architecture, §3 feature status, §7 phases, §8 backlog)

### 1.2 Test run (live)

```
bin\x86\Release\MSOfficeAIAssistant.Tests.exe
...
Running PowerPointActionParserTests... [PASS]
Running DocToDeckTests... [PASS]
Running RollbackExecutorTests...
  [PASS] PowerPoint slide move captures exact inverse coordinates
  [PASS] PowerPoint speaker notes capture and restore state verified
...
ALL TEST SUITES PASSED (40/40)
```

---

## 2. Parity comparison

Legend: ✅ Implemented · ⚠️ Partial / thin · ❌ Not implemented (see gap report)

| # | Copilot (M365) PowerPoint feature | Status | Where it lives in this repo |
|---|---|---|---|
| 1 | Draft a new deck from a prompt | ✅ | "Build deck" chip (PowerPoint-only, `QuickPromptRegistry`); ribbon **Slides** → `OnBuildSlides`; PowerPoint system prompt (`PromptAssembler.cs:60-63`) requests structured slide blocks; `ParseSlideData` + `ExecuteCreateDeckFromOutline` (`ChatSidebar.xaml.cs:1414,1505`) |
| 2 | Draft deck from a Word/document file (doc-to-deck) | ✅ **Verified** | `PromptAssembler.BuildBriefingDeckPrompt` (executive format: title / bullets / notes / visual suggestion), slide-outline parser, preview dialog, view-state guards; `DocToDeckTests` passes |
| 3 | "Add a slide about X" | ✅ | `powerpoint.create_slide` (risk 2) → `ExecuteCreateSlide(title, bullets, layout, slideIndex, speakerNotes)` |
| 4 | Reorganize deck (move / delete / duplicate / hide) | ✅ | `move_slide` (risk 2, **rollbackable**), `delete_slide` (risk 3, non-rollbackable by design), `duplicate_slide`, `hide_slide` / `unhide_slide` |
| 5 | Summarize deck / Q&A over the deck | ✅ | `GetPresentationText` (48k cap, sections + speaker notes), `GetPresentationReviewContext` — deterministic brief with untitled / no-body / duplicate-title statistics |
| 6 | Rewrite text on a slide / tighten it | ✅ | `powerpoint.replace_text` (selection-aware, true replacement), `powerpoint.set_shape_text` (named or indexed shape) |
| 7 | Generate / add speaker notes | ✅ | `powerpoint.set_notes` (risk 1, **rollbackable** via before-state capture); notes included in review context |
| 8 | Sections: create / rename | ✅ | `powerpoint.create_section` (incl. trailing `AddSection` fallback at end of deck), `powerpoint.rename_section` |
| 9 | Insert image with alt text | ✅ | "Insert image" panel button (PowerPoint-guarded), local file picker (.png/.jpg/.jpeg/.webp/.bmp), alt text derived from filename; AI never generates or fetches images (deliberate) |
| 10 | Tables / charts / shapes on slides | ✅ | `powerpoint.add_table` (bounded, data-filled), `powerpoint.add_chart` (typed), `powerpoint.add_shape` |
| 11 | Formatting (font, fit-to-slide) | ✅ | `powerpoint.set_font` (name/size/bold/italic/color on selection), `powerpoint.fit_content` |
| 12 | Apply slide layout / master | ✅ | `powerpoint.apply_layout` by name → `ApplyLayoutToSlide` |
| 13 | Cite slide locations & click-to-navigate | ✅ | `[Slide #N: Title]` / `--- Slide N ---` recognized as clickable citations → `NavigateToSlide`; "Slide N of M" context readout (`GetContextReadout`) |
| 14 | Empty-deck / image-only-slide safety | ✅ | `GetOrCreateActiveSlide` creates slide 1 (ppLayoutText, fallback ppLayoutBlank) on an empty deck; "image-only slides do not crash" is a named release gate (SSOT §8) |
| 15 | Structured undo/rollback breadth | ⚠️ | App-level `CommandBars.ExecuteMso("Undo")` covers all mutations; structured LIFO rollback only for `move_slide` + `set_notes` (`RollbackExecutor.cs:91-140`) |
| 16 | Whole-deck translate | ⚠️ | 9-language Translate submenu targets selection/context only — no slide-by-slide deck loop |
| 17 | Excel data → slide table/chart | ⚠️ | `.xlsx` attachment extraction and `add_table`/`add_chart` both exist; no dedicated "chart this attachment's data on slide N" flow |
| 18 | Add slides from another deck | ⚠️ | `.pptx` attachment extraction is rich (titles/sections/notes/layouts) but there is no dedicated chip/flow — user must ask in chat |
| 19 | Theme/Designer polish, AI image generation, brand kit | ❌ | Rejected in SSOT §8 (cloud + content-policy surface; contradicts "local file you approved" rule) |
| 20 | Presenter Coach / rehearsal feedback | ❌ | Cloud speech; out of scope |
| 21 | Org templates / tenant licensing / Graph / Work IQ | ❌ | Requires Microsoft cloud infrastructure; contradicts Office 2010–2021 support and BYOK-local design |
| 22 | Meeting recap → deck, Outlook integration | ❌ | Enterprise scope; orphaned Outlook host deleted as dead code (SSOT D-4) |
| 23 | Web / stock image search grounding | ❌ | Deliberately deferred (post-Phase D, opt-in BYOK); documented, not silently dropped |

---

## 3. Conclusion

Every **local** Copilot-style PowerPoint capability in the M365 AI Assistant's wheelhouse — draft, doc-to-deck, reorganize, summarize/Q&A, rewrite, speaker notes, sections, tables/charts/shapes, image insert, and slide citations — is **implemented**, and the suite is green (40/40, including the PowerPoint parser, doc-to-deck, and rollback suites).

The genuine remaining local gaps are catalogued and ranked in the companion report:

➡ **[PPT_Gap_Analysis_Roadmap.md](PPT_Gap_Analysis_Roadmap.md)** — thin spots, an SSOT errata, ten ranked implementable features with effort and hook points, and the documented non-implementable list.
