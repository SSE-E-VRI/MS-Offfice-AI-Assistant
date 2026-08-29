# Word Feature Parity — This Add-in vs. Copilot in Word (non-cloud features only)

**Audit date:** 2026-08-29 · **Product version:** 0.5.0 · **Branch:** `fix/adversarial-review-slice5-word-tools`
**Companion files:** [Word_Feature_Audit.md](Word_Feature_Audit.md) · [Word_Feature_Leftovers_and_Deferred.md](Word_Feature_Leftovers_and_Deferred.md)

Scope rule: cloud/enterprise Copilot capabilities that require Microsoft 365 infrastructure (Graph grounding, Work IQ, Designer, tenant licensing, org knowledge) are **excluded** — they contradict this project's Office 2010–365, BYOK, zero-middleman design (SSOT §1) and were never candidates.

---

## Parity table

| Copilot-in-Word feature (local-equivalent) | This add-in | Notes |
|---|---|---|
| Draft with Copilot (new document from a prompt) | ✅ | Generate ribbon button + free chat (`RibbonCallback.cs:45`) |
| Draft grounded in an existing file | ✅ | Attachment extraction (.docx/.pdf/.txt/…) + `@` mention picker (`MentionResolver`, `MentionPicker`) |
| Chat about the document (side panel Q&A) | ✅ | Chat mode + context engine: excerpts, live outline, cursor context (`WordController.cs:1815–1828`) |
| Rewrite selection | ✅ | Rewrite button + `word.find_replace` / selection replace |
| Rewrite with alternatives + regenerate | ✅ | **3 Variants** carousel — insert/regenerate/discard (`RewriteVariantsWindow.cs`) |
| Tone adjustment | ✅ | Tone menu: Official Register, Formal Letter, Executive/Concise, Technical (`Ribbon.xml:46–53`) |
| Visualize as a table | ✅ | Dedicated Visualize as Table button (`RibbonCallback.cs:205`) |
| Summarize document | ✅ | Summarize button + quick-prompt chip |
| Expand / shorten content | ✅ | Expand and Shorten buttons (`RibbonCallback.cs:117,135`) |
| Continue writing / auto-completion | ✅ | Continue button (`RibbonCallback.cs:63`) |
| Coaching / review feedback without mutation | ✅ | Review button — read-only coaching prompt (`WordSearchService.BuildCoachingPrompt`) |
| Outline generation | ✅ | Outline button + live outline context |
| Action-item extraction | ✅ | Actions button + `BuildActionItemContext` |
| Translate | ✅ | 9-language submenu (`Ribbon.xml:54–66`) |
| Track-changes-aware editing | ✅ **exceeds Copilot** | Tracked insert/replace with revision-count readout; accept/reject per-revision, in-selection, or all |
| AI comments | ✅ **exceeds Copilot** | add / list / edit / delete comments as structured actions |
| Document compare | ✅ | Native `CompareDocuments` action **plus** local diff UI (`DocumentCompareWindow` + `DiffEngine`) |
| Formatting automation (styles, fonts, tables, headers/footers, page setup, hyperlinks, bookmarks, images) | ✅ **exceeds Copilot** | 24 typed, risk-gated `word.*` actions with preview + audit + rollback envelope |
| Find & replace automation | ✅ | `word.find_replace` |
| Source citations that navigate | ✅ | `[Source: …]` convention + `[¶N]`/`~Paragraph N` clickable tags jump to the paragraph |
| Grammar-clean insertion (no critique report pasted into the document) | ✅ | Prompt contract + `ResponseContentCleaner` defense-in-depth |
| Requires Microsoft cloud | ✗ correctly rejected | Graph / Work IQ / Designer / AI Pages / tenant services are out of scope by design (SSOT §1) |

## Conclusion

**The add-in is feature-complete for the non-cloud surface it committed to**, and in three areas (track-changes-aware insertion, comment management, and typed formatting actions with approval/audit/rollback) it deliberately goes beyond what Copilot-in-Word exposes. The remaining gap is not parity with Copilot's *shipped* consumer feature set — it is an 11-item list of Office-COM capabilities this project never scoped (TOC, PDF export, watermark, captions, and similar), catalogued in [Word_Feature_Leftovers_and_Deferred.md](Word_Feature_Leftovers_and_Deferred.md).

## What was deliberately rejected (cloud/enterprise, per request)

| Copilot capability | Why rejected here |
|---|---|
| Microsoft Graph / Work IQ / organizational knowledge grounding | Requires tenant licensing + Microsoft cloud; abandons Office 2010–2021 support and local-key design (SSOT backlog: "Not feasible") |
| Designer (slide design intelligence) | Same cloud dependency |
| AI Pages / org-wide AI memory | Deferred past Phase D; requires §11 change control even in local form |
| MCP / external enterprise connectors / RAG beyond single-session file grounding | Out of scope §1, "permanently unless formally re-activated" |
| PowerPoint text-to-image | Rejected — paid image endpoints contradict the "insert a local file you approved" rule |
| Browser automation as a provider | Out of scope §1 |
