# Word Feature Leftovers & Deferred Items

**Audit date:** 2026-08-29 · **Product version:** 0.5.0 · **Branch:** `fix/adversarial-review-slice5-word-tools`
**Companion files:** [Word_Feature_Audit.md](Word_Feature_Audit.md) · [Word_Feature_Parity_vs_Copilot.md](Word_Feature_Parity_vs_Copilot.md)

Two categories: **(A)** features genuinely implementable locally with Word COM that were verified *absent* from the codebase, and **(B)** items the project itself has documented as deferred — backlog entries, not missing implementations of a committed feature. Cloud/enterprise features are excluded entirely (out of scope per SSOT §1).

---

## A. Implementable leftovers — verified missing in source

All items below were probed with direct source scans of `src\Hosts\WordController.cs` (pattern searches for the corresponding COM API) and returned **no matches**: `TablesOfContents`, `ExportAsFixedFormat`, `SaveAs`, `Watermark`, `ReadabilityStatistics`, `ProofingLanguage`, `InsertCaption`, `InsertFile`, `ListStyles` — plus a scan of `TrackRevisions` usage (only internal per-insert toggling, no exposed action). Ranked by value-for-effort.

| # | Feature | Word COM API | Why it fits this project | Effort |
|---|---|---|---|---|
| 1 | **`word.insert_toc` / `word.update_toc`** — table of contents | `Document.TablesOfContents.Add(Range, …)` / `.Item(1).Update()` | Official letters and reports need TOCs; the AI already builds the outline via `TryGetLiveDocumentOutline` | Small |
| 2 | **`word.export_pdf`** (and optionally `word.save_as`) | `Document.ExportAsFixedFormat` / `SaveAs2` | Natural end-of-workflow step; fully local, no cloud | Small |
| 3 | **`word.toggle_track_changes`** — expose tracking on/off as an action | `Document.TrackRevisions = true/false` | Users cannot ask the AI to switch tracking mode today | Small |
| 4 | **`word.list_styles`** — enumerate style names | `Document.Styles` | Risk 0; stops `word.apply_style` from guessing style names the document lacks | Small |
| 5 | **`word.set_proofing_language`** | `Range.LanguageID = wd…` | Pairs with the 9-language Translate flow (spell-check in target language) | Small |
| 6 | **`word.merge_document`** — append another local document | `Document.InsertFile` | Complements `word.compare_documents` (both take an absolute local path) | Small |
| 7 | **`word.set_watermark`** — text watermark (DRAFT / CONFIDENTIAL) | Header `Shapes.AddTextEffect` | Common in government-office documents, the product's primary audience | Medium |
| 8 | **`word.insert_caption`** — captions for tables/figures | `Range.InsertCaption` | Pairs with existing `insert_table`/`insert_image` actions | Medium |
| 9 | **`word.delete`** — targeted deletion (paragraph index / selection / table index) | `Range.Delete` / `Table.Delete` | Deletion is only reachable today via find-replace-with-empty; a typed risk-2 delete action is safer | Medium |
| 10 | **`word.apply_list`** — convert existing paragraphs to bullets/numbering | `Range.ListFormat.ApplyBulletDefault` / `ApplyListTemplate` | Renderer only formats AI-generated content; users can't ask the AI to listify existing text | Medium |
| 11 | **Readability/statistics readout** | `Document.ReadabilityStatistics` | Risk-0 readout backing a "readability" quick prompt; fits the evidence-model philosophy | Small |

### Implementation notes for whichever items are picked up

- Follow the established pattern: `ToolDefinition` registration in `ToolRegistry.cs` (Word section, after `word.insert_image`, ~line 1107) → `Execute*` method in `WordController.cs` returning `HostOperationResult` → append to `FormatActionTypesList("Word")` and the Word system prompt in `PromptAssembler.BuildHostAwareSystemPrompt` → **bump `ExpectedGoldenMasterSha256`** following the documented fixture-diff procedure → add a COM-free test suite in `tests\` and register it in `tests\Program.cs`.
- `word.delete` and `word.export_pdf` are the only ones that plausibly warrant risk 2; everything else fits risk 0–1.
- Update SSOT §2 and §3 in the same change (change-control rule 2).

---

## B. Deferred by the project itself (documented, non-cloud)

These are **not** missing implementations — they are recorded decisions. Starting any of them requires a formal SSOT §11 revision.

| Item | Where documented | State |
|---|---|---|
| Web search / external grounding (opt-in BYOK Brave/SearXNG) | SSOT §3, backlog | **Deferred (post-Phase D)** — touches the zero-middleman privacy guarantee, stays off by default |
| AI Pages, model routing (fast vs. reasoning), local knowledge library, local feedback capture | SSOT §3, §7 | **Deferred past Phase D** — requires §11 change control |
| Cross-host plan coordination in live chat (`CrossHostPlanCoordinator`) | SSOT Phase D follow-up | Backend exists and is tested (`CrossHostPlanCoordinatorTests`); **not wired into chat** because a session's action extractor is single-host today |
| `WorkSession` persistence in live chat | SSOT Phase D follow-up | Backend tested (`WorkSessionStoreTests`); active plans live only in memory on the `ChatMessage` (`[JsonIgnore]`), not saved/reloaded across sessions |
| Persistent "active skill" composition | SSOT §B3 | Skill chips are one-shot prompts; no persisted active skill the AI stays aware of turn-to-turn |
| Doc-to-deck initiated from inside Word | SSOT §7.12–7.18 | Explicitly excluded from v1 — requires cross-process PowerPoint COM automation with its own lifetime/visibility/failure modes |
| Outlook as a fourth host | SSOT backlog rank 6 | `OutlookController.cs` was orphaned dead code and **deleted** (D-4); starting over means Explorer/Inspector lifecycles + a new registration surface |
| Cross-document persistent memory / user-editable fact list | SSOT backlog rank 5 | "Largely already done" via per-document DPAPI `ConversationStore`; only a cross-document fact store would be new |
| Live verification across the full Office/bitness matrix (incl. Excel 2010) | SSOT §3 | The only §3 row marked **Not Implemented** — needs real Office runs, not code |
| Excel live-snapshot sheet-qualified addresses; `.pptx` attachment slide-title resolution | SSOT §2 context table | Cross-host gaps, noted only because they are the context engine's remaining rough edges (not Word-specific) |

## C. Permanently out of scope (cloud / enterprise — excluded per audit request)

GSSMS/CMMS/IoT telemetry integration · external enterprise connectors · MCP · organizational knowledge graphs · RAG beyond single-session file grounding · cloud middleware · Office.js/VSTO rearchitecture · unrestricted silent agentic editing · browser automation as a provider · Graph/Work IQ/Designer · Microsoft 365 Copilot tenant services.

These require Microsoft cloud infrastructure and/or tenant licensing, which would abandon both Office 2010–2021 support and the local BYOK design. They are listed as out of scope in SSOT §1 and were not considered gaps in this audit.

---

**Bottom line:** nothing in the committed, non-cloud Word scope is missing. Category A is an optional enhancement backlog (11 items, mostly small); category B is a recorded decision list, not a defect.
