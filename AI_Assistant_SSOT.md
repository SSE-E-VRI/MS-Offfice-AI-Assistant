# AI Assistant for Microsoft Office — Single Source of Truth (SSOT)

**Document:** `AI_Assistant_SSOT.md`
**Product:** AI Assistant for Microsoft Office
**Shipped version:** 0.5.0
**SSOT revision date:** 2026-08-21
**Developer credit:** Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302

---

## 0. Purpose and precedence

This is the **only** planning/architecture document in this repository. It supersedes and replaces
`AI_Assistant_SSOT_final.md`, `SSOT_Addendum_Post_v0.0.0_Roadmap.md`,
`Copilot_Style_Railway_Agentic_Implementation_Plan.md` (and `_v2`),
`MS_Office_AI_Assistant_Copilot_Improvement_Plan.md`, `Mistral_Office_Addin_Requirements.md`,
`ROOT_CAUSE_AND_FIXES.md`, `FIXES_SUMMARY.txt`, and the former `docs/` folder — all of which were
deleted when this file was created. `README.md` remains as the public-facing project overview.

### Precedence when information conflicts

1. Current source code in this repository
2. Verified runtime behavior and passing tests
3. This SSOT
4. General assumptions

**Never silently convert a planned feature into an implemented one.** Section 3 marks every feature
`Implemented`, `Verified`, `Planned`, or `Not Implemented`. Per §11, nothing may be marked `Verified`
from source inspection alone.

> **Note on the superseded SSOT.** The previous `AI_Assistant_SSOT_final.md` described the original
> `MSOfficePlugin.rar` archive, not the shipped product, and had drifted badly: it recorded streaming
> as "Not implemented", settings in `HKCU\Software\MistralAIOffice`, a 2-button "Mistral AI" ribbon,
> a 12,000-character context cap, the task-pane ProgID as `MistralAI.ChatPane`, and the add-in ProgID
> as `MistralAI.Connect`. **All six are wrong against v0.4.0.** Section 2 below is written from the
> code and replaces those claims.

---

## 1. Product identity and scope

- **Supported hosts:** Microsoft Word, Excel, PowerPoint — Office 2010 → Microsoft 365, 32-bit and 64-bit.
- **AI providers:** Mistral AI, Google Gemini, Groq, Custom OpenAI-compatible endpoint (incl. local Ollama).
- **Model:** Bring-Your-Own-Key. Direct HTTPS to the provider. No intermediary server, no telemetry.
- **File inputs:** `.docx`, `.xlsx`, `.pptx`, `.pdf`, images, and plain-text formats.
- **Credentials:** Windows DPAPI, `CurrentUser` scope, local only.

### Out of scope — permanently, unless formally re-activated under §11

- GSSMS, CMMS, IoT telemetry, asset-master or work-order integration
- External enterprise connectors, MCP, organizational knowledge graphs
- RAG beyond single-session file grounding
- Cloud middleware, Office.js or VSTO-only rearchitecture
- Unrestricted or silent agentic Office editing
- Browser automation as a provider

---

## 2. Current architecture (v0.4.0, as built)

```text
Word │ Excel │ PowerPoint
        │
        ▼
Connect  (IDTExtensibility2, IRibbonExtensibility, ICustomTaskPaneConsumer)
        │
        ├── Ribbon.xml ──► RibbonCallback
        └── CustomTaskPaneManager
                 │  (picks one of three hosting strategies — see §2.3)
                 ▼
           TaskPaneControl (WinForms ActiveX shim, IObjectSafety)
                 └── ElementHost ──► ChatSidebar (WPF)
                                        │
                                        ├── ChatOrchestrator ──► IAIProvider
                                        │        ├── MistralProvider ─┐
                                        │        ├── GroqProvider ────┤► OpenAICompatibleClient
                                        │        ├── CustomApiProvider┘
                                        │        └── GeminiProvider (own HTTP/REST)
                                        ├── AttachmentExtractor (OpenXML / PdfPig / vision)
                                        ├── ConfigManager · ConversationStore · ActionAuditStore (DPAPI)
                                        └── WordController │ ExcelController │ PowerPointController
```

### 2.1 COM identity — authoritative

| Item | Value | Source |
|---|---|---|
| Add-in class | `MSOfficeAIAssistant.Addin.Connect` | `src/Addin/Connect.cs:81` |
| Add-in CLSID | `{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}` | `Connect.cs:81` |
| Add-in ProgID | **`MSOfficeAIAssistant.Addin`** | `Connect.cs:82` |
| Task-pane class | `MSOfficeAIAssistant.Addin.TaskPaneControl` | `CustomTaskPaneManager.cs:100` |
| Task-pane CLSID | `{9B3C7624-5A1D-4C5E-8C9B-12D3E4F5A6B7}` | `CustomTaskPaneManager.cs:100` |
| Task-pane ProgID | `MSOfficeAIAssistant.TaskPaneControl` | `CustomTaskPaneManager.cs:101` |
| Type library GUID | `{A7E4C9B1-3F82-4D16-9E5A-71B08C4D2F90}` | `AssemblyInfo.cs:18` |
| Assembly / file version | `0.5.0.0` | `AssemblyInfo.cs:20-21` |
| Assembly name | `MSOfficeAIAssistant` | csproj |
| Root namespace | `MSOfficeAIAssistant` (matches `RootNamespace` and `AssemblyName`) | — |

Declared Office interface GUIDs (`Connect.cs:32,66`; `CustomTaskPaneManager.cs:18,30,45,75`):
`IDTExtensibility2 {B65AD801-ABAF-11D0-BB8B-00A0C90F2744}`,
`IRibbonExtensibility {000C0396-…}`, `ICustomTaskPaneConsumer {000C033D-…}`,
`ICTPFactory {000C033E-…}`, `_CustomTaskPane {000C033B-…}`,
`IObjectSafety {CB5BDC81-93C1-11CF-8F20-00805F2CD064}`.

> ⚠️ **The csproj `ProjectGuid` is deliberately the same value as the add-in CLSID**
> (`install.ps1:80` depends on it). Never change it.

#### Identity migration (v0.4.0 → v0.5.0)

The Mistral-specific identifiers were renamed in one pass. **The CLSIDs and the type-library
GUID did not change** — only the human-readable names did, so COM class identity is stable.

| Layer | Was | Now |
|---|---|---|
| C# namespace | `MistralOfficeAddin.*` | `MSOfficeAIAssistant.*` |
| Add-in ProgID | `MistralAI.Addin` | `MSOfficeAIAssistant.Addin` |
| Task-pane ProgID | `MistralAI.TaskPaneControl` | `MSOfficeAIAssistant.TaskPaneControl` |
| Ribbon tab id | `tabMistralAI` | `tabAIAssistant` |
| User data folder | `%LOCALAPPDATA%\MistralOfficeAddin` | `%LOCALAPPDATA%\MSOfficeAIAssistant` |
| Log fallback | `%TEMP%\MistralAddinLog.txt` | `%TEMP%\MSOfficeAIAssistant.log` |

Because a renamed ProgID still resolves to the *same* CLSID, an orphaned old ProgID would let
Office load the add-in twice under two names — the same class of failure as the Office 2021
incident in Appendix A. Two mitigations are therefore mandatory and are implemented:

1. **`install.ps1` purges the legacy identities** (`MistralAI.Addin`, `MistralAI.Connect`,
   `MistralAI.TaskPaneControl`, `MistralAI.ChatPane`) from `Software\Classes`, from every
   `Addins` key (versioned and unversioned), and from `DoNotDisableAddinList`, before
   registering the new names. `uninstall.ps1` removes both old and new sets.
2. **`AppPaths` migrates user data** (`src/Core/AppPaths.cs`). On first access it moves the
   legacy folder to the new one, falling back to a non-overwriting file-by-file copy if the
   move is blocked. DPAPI blobs are user-scoped, not path-scoped, so credentials, conversation
   history and the audit trail all survive. `AppPaths` deliberately takes no dependency on
   `Logger`, because `Logger`'s static constructor calls it.

### 2.2 Ribbon

One tab, `tabAIAssistant`, labelled **AI Assistant** (`src/Addin/Ribbon.xml`, embedded resource loaded
in `Connect.cs:322-340`). Four groups, 14 buttons, all handled in `src/Addin/RibbonCallback.cs`:

| Group | Buttons |
|---|---|
| Chat | Open AI Chat, New Chat, Configure |
| Draft | Generate, Continue, Summarize, Rewrite |
| More | Expand, Shorten, Outline, Actions, Review, Slides |
| Help | User Manual |

Every callback is `ShowPane()` + `ExecutePrompt(<hardcoded prompt>, <title>)`.

### 2.3 Task pane — three hosting strategies

`CustomTaskPaneManager.CreateAndShowPane()` (`:578-658`) selects at runtime:

1. **Excel 2010** (`app.Version` starts `"14."`) → docked pane, because the native CTP path is unreliable there.
2. **Otherwise** → native CTP via `ICTPFactory.CreateCTP`, docked right at 380 px.
3. **Fallback** → `OfficeDockedPane` — a `Form` re-parented as `WS_CHILD` into the Office document window.
4. **Last resort** → `ChatFloatingWindow`.

All four paths construct the same `ChatSidebar` WPF control.

> ⚠️ **`OfficeDockedPane` installs a `WH_GETMESSAGE` keyboard hook** (`:435-526`) that intercepts key
> messages and dispatches them into the WPF `HwndSource`, preventing Excel from starting in-cell
> editing while the user types in the prompt box. Together with the focus/HWND plumbing in
> `ChatSidebar.xaml.cs:1356-1440`, this is **load-bearing**. Validate any new focusable control
> against it, and do not refactor it casually.

### 2.4 Chat engine

- `ChatMessage` (`src/API/Models/ChatRequest.cs:8-141`) is simultaneously the wire format
  (Newtonsoft `[JsonProperty]`) and the view model (`INotifyPropertyChanged`). Roles: `system`,
  `user`, `assistant`. `FullContent` carries the real prompt while `Content` carries the short
  display title.
- History is `ObservableCollection<ChatMessage>` in `ChatSidebar`, persisted per document by
  `ConversationStore` (DPAPI `.dat` under `%LOCALAPPDATA%\MistralOfficeAddin\Conversations\`).
- **Token budget is 24,000**, applied by `TokenCounter.TruncateToFit` — a heuristic estimator
  (`src/API/TokenCounter.cs`), not a real tokenizer. The budget is hardcoded at the call site.
- Streaming is live SSE. The delta callback throttles to **every 5th delta** and marshals via
  `Dispatcher.BeginInvoke` at `Background` priority.

### 2.5 Providers

`IAIProvider` (`src/Providers/IAIProvider.cs:113-123`) exposes exactly: `ProviderType`,
`Capabilities`, `TestConnectionAsync`, `ListModelsAsync`, `ChatAsync`, `StreamChatAsync`,
`CheckVisionSupport`.

Mistral, Groq and Custom share `OpenAICompatibleClient` (endpoint validation, HTTPS enforcement,
3× exponential backoff on 429/5xx). Gemini has its own client and payload shape. Two independent SSE
parsers exist (`src/API/StreamingParser.cs`, `src/Providers/GeminiStreamingParser.cs`).

`ChatOrchestrator` (235 lines) is **only** a thread-safe provider-lifetime wrapper — provider swap,
cancellation, streaming pass-through. It contains no prompt, Office, or action logic.

> **Provider Capabilities:** `AICapabilities` (`IAIProvider.cs:17-29`) supports `StructuredOutput`, `ToolCalling`, and `JsonMode` flags. `OpenAICompatibleClient.BuildPayload` and `GeminiProvider.BuildGeminiPayload` support structured JSON mode (`response_format`/`responseMimeType`), `tools`, `tool_choice`, and arbitrary extra parameters. `StreamingParser.TryParseLine` extracts `choices[0].delta.content`.

### 2.6 Office context extraction

| Host | Method | Location fidelity |
|---|---|---|
| Excel | `GetWorksheetSnapshot(70, 26)` — sheet name, ActiveCell, Selection, UsedRange, per-cell `A1=value` with formula annotations | **Good** — the only place addresses survive. Active sheet only; addresses not sheet-qualified. |
| Word | `GetRelevantDocumentContext` → chunk-and-score against prompt terms, plus live outline and cursor context | **Poor** — output labelled only `[Excerpt i of n]`; no paragraph numbers or offsets. |
| PowerPoint | `GetPresentationReviewContext`, `GetPresentationText`, `GetPresentationOutline` | **Partial** — slide numbers in some paths, absent in others. |
| Attachments | `AttachmentExtractor` — OpenXML, PdfPig | **Mixed** — PDF keeps real page numbers; `.pptx`/`.xlsx` emit ordinal indices only; `.xlsx` **drops cell addresses entirely**; `.docx` has no paragraph index. |

Scope is chosen by the user: `Selection`, `CurrentFile`, `SelectionAndFile`, `AttachmentsOnly`.

Limits: 20 MB/file, 30 MB total, 10 files, 50,000 chars, 100 MB decompressed (zip-bomb guard).
Legacy `.doc/.xls/.ppt/.rtf` are rejected outright.

### 2.7 Structured actions and write-back

Three **mutually incompatible** extraction mechanisms exist today:

1. **Excel — `<excel_actions>` XML.** `SpreadsheetActionParser.ExtractActions`
   (`src/Core/SpreadsheetAction.cs:252-364`). 13 action types. Hardened `XmlReader`
   (`DtdProcessing.Prohibit`, `XmlResolver = null`). Safety bounds: max 25 actions, 12,000 chars per
   action, 100,000 cells, 200 KB block, and **bounded-A1 targets only** — sheet-qualified refs
   (`Sheet1!A1`), whole columns (`B:B`) and multi-area ranges are deliberately rejected.
   This is the only parser wired into the chat response path.
2. **PowerPoint — `<powerpoint_actions>` XML.** `PowerPointActionParser` (4 allowed types).
   Invoked from inside `PowerPointController.InsertText`, **not** from the response pipeline.
3. **Excel free text.** `ExcelController.ExtractCleanExcelContent` — regex heuristics for the
   Insert button.

**Word has no structured action format at all.** The model returns prose and the whole response is
inserted.

Approval is `System.Windows.MessageBox` in every case. The entire risk model is one boolean,
`SpreadsheetAction.IsUndoable`, which is `false` for RemoveDuplicates, CreateTable, Chart,
PivotTable and NamedRange.

### 2.8 Undo

Three unrelated mechanisms, none unified:

- **Word** — real Office undo grouping via `UndoRecord.StartCustomRecord` / `EndCustomRecord`
  (`WordController.cs:641,660`), degrading gracefully on Word 2007. **This is the pattern to generalize.**
- **Excel** — bare `app.Undo()`. Unreliable from COM; useless for the five non-undoable types. A
  batch apply produces N separate undo steps, or none.
- **PowerPoint** — `app.CommandBars.ExecuteMso("Undo")`.

There is **no before-state capture, no snapshot, and no custom rollback anywhere.**

### 2.9 Storage and security

| Store | Path (under `%LOCALAPPDATA%\MSOfficeAIAssistant\`) | Protection |
|---|---|---|
| Settings + API keys | `config.dat` | Whole file DPAPI, `CurrentUser`, no entropy |
| Conversations | `Conversations\{key}.dat` | DPAPI; legacy `.json` migrated on load |
| Action audit | `action-audit.dat` | DPAPI; 250 entries max, 2,000 chars/field; `.bak` quarantine on corruption |
| Log | `addin.log` (fallback `%TEMP%\MSOfficeAIAssistant.log`) | Plain text |

`ActionAuditEntry` records `TimestampUtc, Host, ActionType, Target, Summary, Undoable, Prompt,
SourceContext, Model, FullProposedAction, ApplyResult`. It is **append-only, written only after
success, and never read back for rollback** — it is displayed as text and nothing more.

Security invariants: never log API keys; HTTPS enforced for remote providers (HTTP permitted only
for loopback); context is sent only when the user has enabled it; no telemetry leaves the machine.

### 2.10 Threading

Office object-model access happens on the Office UI thread. HTTP runs on the thread pool. Results
marshal back through the WPF `Dispatcher`. All COM interop must respect this — Office object model
calls from a background thread will fail or corrupt state.

### 2.11 Interop strategy — mixed, and it matters

- **Word** uses **NetOffice 1.9.10 typed wrappers** (`using Word = NetOffice.WordApi`), with a lazy
  `GetApp()` guard. `GetSelectedText` is deliberately dual-path (tries `dynamic` first) because
  NetOffice wrapping can trigger COM event subscriptions on Word's UI thread.
- **Excel and PowerPoint** use **pure `dynamic` late binding** with raw magic integers for every
  Office enum (e.g. `ExcelController.cs:501` `5 // xlGreater`; `PowerPointController.cs:96`
  `slides.Add(1, 2) // ppLayoutText`).

Consequence: new Excel/PowerPoint operations need no references but get **zero compile-time
checking** — a typo becomes a runtime `RuntimeBinderException`. Enum constants must be looked up by
hand. No Office PIAs are used anywhere.

#### Host isolation and controller exclusivity

Each Office host (Word, Excel, PowerPoint) executes as an independent OS process and instantiates
its own isolated add-in COM instance. Consequently, `_wordCtrl`, `_excelCtrl`, and `_pptCtrl` are
**strictly mutually exclusive** — live Word and PowerPoint controllers never coexist in memory.
Cross-document workflows like doc-to-deck operate via COM-free file extraction (`AttachmentExtractor`),
never through cross-process live COM.

### 2.12 Error handling conventions

**No `COMException` is caught anywhere in the codebase** — everything catches bare `Exception`. Three
patterns are in use:

1. **Swallow-and-warn** on read paths — log, return empty/false. The caller cannot distinguish
   "empty document" from "COM failed".
2. **Log-and-rethrow** on mutating paths — surfaces as a `MessageBox`.
3. **Status object** — `ExcelController.ApplySpreadsheetAction` (`:317-323`) sets
   `Status`/`ErrorMessage` and returns false. **This is the only structured error channel in the
   controller layer** and the natural seed for a tool-result envelope.

Additionally ~60 bare `try { … } catch { }` blocks wrap individual COM property reads. This is
intentional — it absorbs Office 2010↔365 API differences — but it means **a failed sub-step is
invisible**, which any future verification layer must account for.

### 2.13 Build, install, verify

> `build.bat` and `register.cmd` **do not exist.** They were deleted in commit `97caea8` (v0.3.0).
> There is also **no `.sln`**.

- **Build only (fastest loop):** MSBuild directly on `src/MSOfficeAIAssistant.csproj`, once per
  platform, `Configuration=Release`, `Platform=x86` then `Platform=x64`. Output to
  `bin\{plat}\Release\`.
- **Build + install:** `install.cmd` → `install.ps1`. Locates MSBuild → downloads `nuget.exe` if
  absent → restores → builds x86 and x64 → cleans stale CLSID trees → `RegAsm /codebase /regfile`
  with `HKCR`→`HKCU\Software\Classes` rewriting → writes
  `HKCU\Software\Microsoft\Office\{Word,Excel,PowerPoint}\Addins\MSOfficeAIAssistant.Addin\LoadBehavior = 3`.
  It **cannot** be run build-only. x86 regfile entries are routed to `Wow6432Node` deliberately.
- **Uninstall:** `uninstall.cmd` / `uninstall.ps1`.
- **Distribution:** `installer/setup-x86.iss`, `installer/setup-x64.iss` (Inno Setup 6).
- **Smoke check:** `tools/verify.ps1` — COM-creates each host and asserts the add-in is loaded.
  Requires Office installed.

#### Project file constraints

- Old-style non-SDK MSBuild, `TargetFrameworkVersion v4.8`, **`LangVersion 5`**.
  No string interpolation, `nameof`, `?.`, tuples, pattern matching, or switch expressions.
- Four explicit configurations (`Debug|x86`, `Release|x86`, `Debug|x64`, `Release|x64`).
  **No `AnyCPU`.** `RegisterForComInterop=false` — registration is external.
- References are `<Reference>` + `<HintPath>` into `..\packages\`. No `<PackageReference>`.
- **Source files are listed explicitly. There is no wildcard globbing.** Every new `.cs` needs a
  `<Compile Include>` entry; a new WPF window needs **two** entries (`Compile` + `Page`).

### 2.14 Tests

`tests/MSOfficeAIAssistant.Tests.csproj` — a **hand-rolled console runner**, no NUnit/xUnit/MSTest.
`tests/Program.cs` runs 5 suites and returns exit 0/1.

Run: build the tests csproj, execute `bin\{plat}\Release\MSOfficeAIAssistant.Tests.exe`.

`InternalsVisibleTo("MSOfficeAIAssistant.Tests")` is wired (`AssemblyInfo.cs:22`), and the tests
csproj **deliberately omits every NetOffice reference** — that is what keeps the suite Office-free.
Anything testable must therefore be COM-free. Existing COM-free seams:
`WordDocumentContextBuilder`, `SpreadsheetActionParser`, `PowerPointActionParser`, `ActionAuditStore`
(via its custom-path constructor), `ExcelController.ExtractCleanExcelContent`.

#### Golden Master test fixture (Phase 0.0 Gate)

`BuildHostAwareSystemPrompt` inspects `_hostType` and null-checks controller references without
calling COM, allowing prompt construction to be lifted to a pure function immediately. A headless
**Golden Master test fixture** records baseline prompt strings, XML action parsing DTOs, and DPAPI
audit serialization outputs into a JSON baseline fixture before refactoring. The console test runner
asserts zero diff against this baseline to mechanistically verify behavior parity during Phase 0.1 extraction.

There is **no CI**.

---

## 3. Feature status

| Feature | Status |
|---|---|
| Shared COM add-in, Word/Excel/PowerPoint, Office 2010→365 | Implemented |
| Ribbon (14 buttons), custom task pane, 3 hosting strategies | Implemented |
| Provider abstraction + Mistral / Gemini / Groq / Custom | Implemented |
| SSE streaming with cancellation | Implemented |
| Attachments: docx/xlsx/pptx/pdf/images + vision routing | Implemented |
| DPAPI config, conversation, and audit stores | Implemented |
| Word: insert, replace, Track Changes, accept/reject, Markdown→table | Implemented |
| Excel: 13 confirm-before-apply action types with safety bounds | Implemented |
| PowerPoint: deck build, slide move, sections, speaker notes, image insert | Implemented |
| Embedded offline User Manual | Implemented |
| Office 2021 stale-COM-registration fix | Implemented (see Appendix A) |
| Mistral-neutral identity rename + legacy purge + data migration | Implemented |
| From document → briefing deck (doc-to-deck) | Planned — builds on Phase 0 foundation |
| Live verification across the full Office/bitness matrix | **Not Implemented** |
| Chat / Plan / Edit modes | Planned — Phase A |
| Context bar, source citations, response cards | Planned — Phase A |
| Skills and domain packs | Planned — Phase B |
| Unified action schema, tool registry, risk levels, verification, rollback | Planned — Phase C |
| Multi-step planner, cross-host workflows | Planned — Phase D |
| Web search / external grounding | Deferred (Post-Phase D) — opt-in BYOK client-side search |
| AI Pages, model routing, knowledge library, feedback capture | Not Implemented — deferred past Phase D |
| GSSMS / CMMS / IoT / connectors / RAG | **Out of scope** (§1) |

---

## 4. Known defects and structural debt

| ID | Severity | Issue |
|---|---|---|
| ~~D-1~~ | — | **RESOLVED.** `PowerPointController.InsertText` strips action XML without executing. `<powerpoint_actions>` XML is parsed on response completion into `ChatMessage.PowerPointActions`, and actions are reviewed and approved via dedicated UI cards with typed status and audit logging. |
| **D-2** | Medium | `GetOrCreateActiveSlide(createIfNone:true)` and `GetActivePresentation(createIfNone:true)` are read-shaped names that **create a presentation or slide as a side effect** (`PowerPointController.cs:34,96`). Any risk classification must treat them as mutating. |
| ~~D-3~~ | — | **RESOLVED.** `tools/verify.ps1` checked ProgID `MistralAI.Connect` while the add-in registered `MistralAI.Addin`, so the smoke check always reported FAIL and pointed at the deleted `register.cmd`. Now checks `MSOfficeAIAssistant.Addin` and recommends `install.cmd`. |
| **D-4** | Medium | `src/Hosts/OutlookController.cs` is on disk and git-tracked but **absent from the csproj compile list** — it is silently never built. Either compile it or delete it. |
| **D-5** | Medium | Action-type allow-lists are duplicated in **four** places — `SpreadsheetAction.cs:366`, `PowerPointActionParser.cs:36`, the `ExcelController.cs:250` switch, and a **prompt string literal** at `ChatSidebar.xaml.cs:729` — and must be kept in sync by hand. |
| ~~D-6~~ | — | **RESOLVED.** Session orchestration extracted from `ChatSidebar.xaml.cs` into `src/Core/Session/` (`AssistantSession`, `StreamCoordinator`, `PromptAssembler`). View code-behind handles rendering and HWND routing only. Headlessly verified with dedicated session tests and Golden Master hash gate. |
| **D-7** | Low | Streaming re-parses the **entire** markdown string every 5th delta, and `VirtualizingStackPanel.IsVirtualizing` is explicitly `False` (`ChatSidebar.xaml:104`). Both degrade long conversations. |
| **D-8** | Low | `MarkdownToFlowDocumentConverter` (`src/UI/Converters/MarkdownConverter.cs:12`) is fully written but referenced by nothing — dead code. Markdig is a live dependency used only for Word insertion. |
| **D-9** | Low | `RibbonCallback.OnTranslate` (`:121-151`) implements 9 languages but **has no corresponding ribbon XML** — orphaned and unreachable. |
| **D-10** | Low | Dead methods with no callers: `ExcelController.CreatePreviewDescription` (duplicates the UI's own `DescribeSpreadsheetAction`), `ExcelController.WriteFormula`, `PowerPointController.GetPresentationOutline`, `PowerPointController.SetSpeakerNotes`. |
| ~~D-11~~ | — | **RESOLVED.** `README.md` documented the non-existent `build.bat` and `register.cmd`; corrected to `install.cmd` plus a direct-MSBuild loop. |
| ~~D-12~~ | — | **RESOLVED.** Implemented `OleMessageFilter` (`CoRegisterMessageFilter` on STA thread) to handle Excel in-cell edit busy rejections (`0x800AC472` / `RPC_E_SERVERCALL_RETRYLATER`) with 10s retry window; previous filter restored on add-in disconnect. Added typed `SafeOfficeProbe` helpers. |

---

## 5. Target architecture

```text
User request
     ▼
Chat │ Plan │ Edit          ← mode governs mutation permission
     ▼
Context Engine ──► ranked, location-tagged fragments
     ▼
Skill Engine (domain pack: general │ railway)
     ▼
Prompt Assembler
     ▼
AI Provider (Mistral │ Gemini │ Groq │ Custom │ local)
     ▼
free-form response OR native tool call
     ▼
Action Extractor ──(on failure)──► ExtractionFailure ──► editable Plan mode
     ▼
Structured Action Schema  ← the safety contract
     ▼
Schema Validator ──► Risk Validator ──► Approval
     ▼
Host Controller (Word │ Excel │ PowerPoint)
     ▼
Verification ──► Audit / Undo
```

**The Structured Action Schema — not native provider tool-calling — is the safety contract.** Native
tool-calling, where a provider supports it reliably, is an optimization layered on top; it is never a
requirement for a provider to remain usable for structured actions.

### 5.1 Interaction modes

| Mode | Purpose | Mutation |
|---|---|---|
| **Chat** | Q&A, explain, summarize, analyze | **None.** Enforced in the session layer, not the UI. |
| **Plan** | Convert a complex request into ordered, editable steps | None until approved |
| **Edit** | Execute approved actions through the schema | Risk-gated, previewed, verified, audited |

The mode is always visible and user-controlled. **The AI may never switch modes.**

### 5.2 Risk classification

| Level | Description | Examples | Requirement |
|---|---|---|---|
| **0** | Read-only | Read range, outline, slide text | Silent execution |
| **1** | Low | Insert text at cursor, new sheet/slide, styling | Lightweight confirm |
| **2** | Medium | Overwrite formulas, replace data, reorder slides | Preview + diff + explicit approval |
| **3** | High | Delete sheet/slide, bulk clear, structural change | Strong confirm + snapshot + audit |

### 5.3 Structured action schema

```json
{
  "action_id": "uuid",
  "host": "Word | Excel | PowerPoint",
  "operation": "excel.create_chart",
  "target": { "sheet": "Dashboard", "range": "A4:C18", "slide": null, "paragraph": null },
  "input": {},
  "expected_result": "Chart object created at Dashboard!E4",
  "risk_level": 2,
  "requires_approval": true,
  "rollback_info": { "strategy": "delete_created_shape", "shape_name": "…" },
  "source_reason": "Top 3 stations contribute 68% of incidents",
  "evidence": [
    { "location": "Failure Data!B4:B18", "extracted_value": "42", "evidence_level": "DIRECTLY_OBSERVED" }
  ]
}
```

### 5.4 Evidence model

Every factual claim must map to an extractable Office location and carry an evidence level —
**never model self-report**:

`DIRECTLY_OBSERVED` · `CALCULATED` · `STRONG_INFERENCE` · `POSSIBLE_INFERENCE` · `INSUFFICIENT_EVIDENCE`

`confidence_in_extraction` is tracked **separately**, because an extraction can be
`DIRECTLY_OBSERVED` in intent yet unreliable in execution — for example a `COUNTIF` over a range
containing merged cells or mixed types.

---

## 6. Domain packs

The domain layer ships as **two swappable packs** so one binary serves both railway and
non-railway users (Tamil Nadu government departments, corporate and private sector). The pack is
selected in Settings, persisted in `ConfigDto`, and defaults to `general`. Packs are **additive
manifests, not code forks.**

### 6.1 `general` — default

Official Letter · Minutes of Meeting · Inspection Report · Technical Note · Management Summary ·
Root Cause Analysis · Management Dashboard · Improve Official Language · Document Comparison.

Neutral government/corporate register. No railway vocabulary.

### 6.2 `railway`

Everything in `general`, plus DRM Briefing · Failure Analysis & Pareto · Substation/OHE Asset
Health · Deficiency Tracker — and a terminology layer: Depot, Station, Substation, Asset, Failure,
Deficiency, Inspection, PM, CM, Breakdown, Root Cause, Corrective/Preventive Action, OHE, TRD, DRM,
Sr.DEE, SSE, JE.

### 6.3 Domain rules — enforced in both packs

- Do **not** invent inspection data, equipment specifications, failure causes, or official references.
- Distinguish observed evidence from inference; label every finding with an evidence level.
- Preserve all numerical values exactly.
- Flag missing information rather than filling it in.
- Every important claim carries a real, resolvable source location.
- Never present a forecast as a guaranteed result.

### 6.4 Skill definition

```json
{
  "id": "failure_analysis",
  "name": "Failure Analysis",
  "description": "…",
  "required_context": ["current_file", "selection"],
  "preferred_host": "Excel",
  "prompt_template": "…",
  "output_structure": ["findings", "ranking_table", "pareto", "recommendations"],
  "default_mode": "Plan",
  "risk_ceiling": 2,
  "domain_pack": "railway"
}
```

A Skill is a **structured prompt template with declared required context — not an agent.** Skills
feed the Planner; they do not execute.

---

## 7. Implementation roadmap

Ordering is deliberate: Step 0 merges pending work and bumps version; Step 1 fixes live approval defects;
Phase 0 establishes the test baseline, extracts orchestration, hardens COM resilience, and provides the
clean core upon which doc-to-deck and subsequent phases are built.

### Immediate Execution Sequence

1. **Step 0 — Merge & Version Bump:** Merge pending work; bump assembly and package version `0.4.0` → `0.5.0`.
2. **Step 1 — Fix D-1 (PowerPoint Approval Hole):** Pull `ApplyStructuredActions` out of `InsertText`; route
   through the response pipeline and approval dialog showing parsed actions before execution.
3. **Phase 0.0 — Golden Master Baseline Fixture:** Lift `BuildHostAwareSystemPrompt` to a pure function.
   Record prompt strings, XML parsing DTOs, and DPAPI audit serialization outputs into a JSON baseline fixture
   in the COM-free test runner.
4. **Phase 0.1 — Extract Orchestrator (D-6):** Extract orchestration out of `ChatSidebar.xaml.cs` into
   `src/Core/Session/` (`AssistantSession`, `PromptAssembler`, `StreamCoordinator`). Verify bit-for-bit parity
   against the Phase 0.0 baseline.
5. **Phase 0.2 — COM Resilience (D-12):** Implement `IOleMessageFilter` (`CoRegisterMessageFilter`) to handle
   Excel `0x800AC472` (`VBA_E_IGNORE`) modal formula-edit rejections. Replace bare swallows with typed
   `SafeOfficeProbe<T>` helpers; mutations must never catch bare `Exception`.
6. **Next Slice — From Document → Briefing Deck (Doc-to-Deck):** Build the "From document…" entry point and
   briefing deck generator directly on the clean `AssistantSession` core with a structured slide preview card.
7. **Phase 0.3 — UI Foundation & Performance:** Design system (`Tokens.xaml` + `Controls.xaml`), Win32/ElementHost
   DPI awareness handling, re-enable list virtualization, and incremental markdown rendering (D-7, D-8).
8. **Phase C — Safe Execution:** Unified JSON `OfficeAction` schema, single-source Tool Registry (D-5),
   risk levels 0–3, preview cards, post-execution verification engine, and before-state rollback.
9. **Phase A / B / D — Copilot UX, Domain Packs & Multi-Step Planner:** Chat/Plan/Edit modes, context bar,
   source citations, `general` and `railway` domain packs, multi-step execution state machine.

---

### Next slice — From document → briefing deck (doc-to-deck)

Built on the clean `AssistantSession` core following Phase 0.2. The pipeline already exists end to end:

**Already built — verified against the code:**

| Capability | Where |
|---|---|
| `.docx` / `.pdf` / `.pptx` / `.xlsx` extraction | `AttachmentExtractor.ExtractAsync` |
| Word outline + prompt-relevant context | `WordController.GetDocumentOutline`, `GetRelevantDocumentContext` |
| Outline → slide model (title, bullets, notes, visual) | `PowerPointActionParser.ParseSlideData:113` |
| Layout-aware slide creation | `PowerPointController.AddSlideUsingPresentationLayout:708` |
| Speaker notes | `PowerPointController.SetSpeakerNotesForSlide:471` |
| Deck apply | `PowerPointController.CreateOrUpdateDeckFromOutline:368` |
| **A "Build deck" chip, PowerPoint-only** | `ChatSidebar.xaml:307`, gated at `ChatSidebar.xaml.cs:239` |
| **Confirm → apply already wired** | `ChatSidebar.xaml.cs:1125` |

**Scope to build in this slice:**

1. A **"From document…"** chip and ribbon button sourcing content from a **selected attachment**
   (`.docx` / `.pdf`) rather than from the chat transcript.

   > **Scope boundary.** Controllers are host-exclusive (§2.11), so from the PowerPoint host there
   > is no live `WordController` to read an open document through. Attachment-sourced generation is
   > the v1 scope. A Word-host-initiated *"turn this document into a deck"* is a different and
   > heavier feature — it requires launching and automating a **separate PowerPoint process** over
   > cross-process COM, with its own lifetime, visibility and failure modes. It is explicitly **not**
   > in this slice; schedule it separately if wanted.
2. A fixed **briefing-deck prompt** ("5–10 slides: executive summary, findings, actions…") that
   emits the outline shape `ParseSlideData` already understands.
3. A **preview card** listing the proposed slide titles before anything is created — reusing the
   Excel action-card pattern rather than the current prose-only `MessageBox`.

---

### Phase 0 — Foundation breakdown

| # | Work | Addresses | Status |
|---|---|---|---|
| 0.0 | **Golden Master Baseline:** Lift prompt assembly to pure static function; create headless test fixture recording prompt strings, action parsing DTOs, and audit serialization. Canonical SHA-256 (`1f3971ed6790ed842fe73df80479d9e18f2ca723efd29ccde72454023918e4ca`) and `golden_master_baseline.txt` fixture committed to gate Phase 0.1 onward against byte-for-byte drift. | Verification Gate | Implemented |
| 0.1 | **Extract Orchestrator:** Move prompt, streaming, and session logic into `src/Core/Session/` (`AssistantSession`, `PromptAssembler`, `StreamCoordinator`). View keeps rendering & HWND hooks only. | D-6 | Implemented |
| 0.2 | **COM Resilience:** Implement `IOleMessageFilter` for Excel busy rejection (`0x800AC472`); implement typed `SafeOfficeProbe<T>` for 2010↔365 version probing; unswallow mutation errors. | **D-12**, §2.12 | Implemented |
| 0.3 | **UI Foundation & Theme:** Design system `Tokens.xaml` + `Controls.xaml`; ElementHost DPI handling; re-enable list virtualization; incremental markdown. | D-7, D-8 | Planned |
| 0.4 | **Controller Interface:** `IOfficeHostController` over the three controllers, replacing null-check dispatch. Resolve `OutlookController`. | D-4 | Planned |
| 0.5 | **Provider Capabilities:** Add `StructuredOutput` / `ToolCalling` / `JsonMode` to `AICapabilities`; make `BuildPayload` extensible. | §2.5 | Implemented |

**Exit:** identical responses, actions and audit entries before and after extraction; PowerPoint
structured actions visible in approval dialog; Excel in-cell edit rejections handled gracefully.

### Phase A — Copilot UX (low risk, no new mutation pathways)

- **A1** Chat/Plan/Edit selector, enforced in `AssistantSession` — Chat hard-blocks risk ≥1.
- **A2** Context bar: explicit toggles (Selection / Current file / Attachments / Entire document) plus
  a live host readout. Backed by a `ContextEngine` emitting `ContextFragment { SourceType, Location, Text }`.
- **A3** Response cards via `DataTemplateSelector`: Finding, Recommendation, Plan, ActionPreview,
  Warning, Summary/KPI, Table. The existing Excel action card is the visual precedent.
- **A4** Source tags on claims. Requires fixing the extractors that lose provenance (§2.6) — `.docx`
  paragraph index, `.xlsx` real sheet name + cell address, Word paragraph offsets, sheet-qualified
  Excel addresses. Click-to-navigate where the host allows.
- **A5** Suggested-prompt chips generated from the Skill registry, replacing the 8 hardcoded buttons.
- **A6** Conversation history UI (session list) and a real action-history panel over `ActionAuditStore`.
- **A7** Status indicators (Ready / Thinking / Reading / Planning / Awaiting approval / Applying /
  Verifying / Completed / Failed / Cancelled) and a full accessibility pass — keyboard navigation,
  visible focus, screen-reader labels, no colour-only status.

**Exit:** the user can see exactly what context is sent, switch modes, and receive structured cards
with working source links. Mutation behavior unchanged.

### Phase B — Skills and domain packs (low risk, parallel with A)

- **B1** `src/Core/Skills/` — `Skill`, `SkillRegistry`, JSON manifests as embedded resources.
- **B2** Ship the `general` and `railway` packs (§6). Selector in Settings.
- **B3** Move every inline prompt literal — `BuildHostAwareSystemPrompt`, the 13 `RibbonCallback`
  prompts, the `ConfigManager` default — into `PromptAssembler`, composed as
  *base + host rules + domain pack + selected skill*.
- **B4** Evidence levels rendered on Finding cards, backed by real `Source` locations from A4.
- **B5** Context-aware skill promotion, surfaced through A5 chips.

**Exit:** the same prompt under `general` and `railway` produces an appropriately different register;
domain rules demonstrably suppress invented references.

### Phase C — Safe execution (medium risk)

- **C1** `src/Core/Actions/OfficeAction.cs` implementing §5.3. `SpreadsheetAction` becomes a
  view-model projection. **Keep the `<excel_actions>` XML parser as a compatibility path** so stored
  conversations keep working.
- **C2** `src/Core/Tools/` — the Tool Registry. Each entry declares name, input schema, risk level,
  host method, rollback strategy and validation rules, and the registry **generates** the prompt's
  action catalog, the validator allow-list and the dispatch table — retiring all four duplicated
  allow-lists (D-5). Must be written in C# 5 style: `Dictionary<string, ToolDefinition>` with explicit
  delegates. Initial set includes **a Word action set, which does not exist today** (§2.7).
- **C3** Action Extractor accepting native tool calls or embedded JSON; on malformed output emit a
  typed `ExtractionFailure` and route into **editable Plan mode**, never a silent chat-only fallback.
  `StreamingParser.TryParseLine` must stop discarding non-content deltas.
- **C4** Risk gating with in-panel risk-badged preview cards replacing every `MessageBox`. Preserve
  and extend the existing safety bounds (§2.7).
- **C5** Verification engine — compare observed against `expected_result`; detect `#REF!`/`#VALUE!`;
  offer Fix / Review / Cancel. **Prerequisite:** make the ~60 silent `catch {}` swallows reportable
  (§2.12), or verification cannot see partial failures.
- **C6** Audit v2 — add `ActionId`, `BeforeState`, `RiskLevel`, `PlanId`; capture before-state at
  apply time; generalize Word's `UndoRecord` grouping so a batch apply is **one** undo step in every
  host, not N or zero.

**Exit:** a multi-step plan is reviewable, approvable per-step or wholesale, executed safely,
verified and undoable — identically across all four providers.

### Phase D — Agentic plan-then-execute (medium–high risk)

- **D1** `src/Core/Planning/` — Planner producing ordered `PlanStep`s bound to Tool Registry entries.
  Editable before any risk ≥1 action runs.
- **D2** Execution state machine: Queued / Running / Awaiting approval / Paused / Completed / Failed /
  Cancelled / Rolled back. Progress, cancellation, partial-failure recovery, continue-from-step-N.
- **D3** Cross-host workflows (Excel analysis → Word report → PowerPoint briefing) as sequential
  approved actions over `IOfficeHostController`. No silent cross-application mutation.
- **D4** `WorkSession` — persist Conversation + Context + Plan + Actions + Sources as one reopenable unit.

**Exit:** "Analyze these failures, build a dashboard, then draft a briefing" yields an editable plan
that executes across hosts under approval with a full audit trail.

### Deferred past Phase D

AI Pages · model routing (fast vs reasoning) · local knowledge library · document-comparison skill ·
local feedback capture. Documented as **not activated**; requires §11 change control to start.

### Feature backlog — ranked

Assessed for fit against this add-in's actual constraints (COM, Office 2010+, BYOK, provider-neutral,
confirm-before-write), not against Copilot's feature list.

| Rank | Feature | Verdict |
|---|---|---|
| 1 | **Doc-to-deck** | Ship first (above). Pipeline exists; only the entry point is missing. |
| 2 | **Word Review → native comments** | Build next. Genuinely new — `grep` finds **no** Comments API in any host controller. `Document.Comments.Add(Range, Text)` works from Office 2010 onward, and it is non-destructive, so it fits the existing safety model without new risk tiers. |
| 3 | **`@` mention local files** | Same local grounding, better UX. Pure WPF work over `AttachmentExtractor`. Delivers the "Copilot `/file`" feel with no cloud dependency. Helps Word, Excel and doc-to-deck alike. |
| 4 | Chart/Pivot "explain + suggest" | Cheap. Prompt work over the existing `excel_actions` chart/pivot types. A chip, not a project. |
| 5 | Persistent memory | **Largely already done** — `ConversationStore` is per-document DPAPI history. Only worth revisiting for cross-document memory or a user-editable fact list. |
| 6 | Outlook summarise/reply | Feasible (`OutlookController` exists, see D-4) but widens the product to a fourth host with its own Explorer/Inspector lifecycles and registration surface. Tighten the three-app loop first. |
| 7 | Plan/agent loop | Highest effort; this is Phases C–D. Bounded action XML already covers the common cases. |
| ↓ | **Web search / external grounding** | **Deferred (Post-Phase D).** Client-side, provider-neutral, opt-in BYOK search (e.g. Brave Search / SearXNG HTTP API), off by default to maintain zero-middleman and privacy guarantees. Enables drafting tasks citing external statutes, tax codes, and ISO standards without cloud lock-in. |
| ✗ | PowerPoint text-to-image | Rejected. Needs paid image endpoints, adds content-policy surface, and contradicts the deliberate "Insert image = a local file you approved" rule. |
| ✗ | Graph / Work IQ / Designer | Not feasible. Requires Microsoft cloud infrastructure and tenant licensing, abandoning both Office 2010–2021 support and the local-key design. |

---

## 8. Verification requirements

Applies to every phase before merge.

1. **Build** — MSBuild `src/MSOfficeAIAssistant.csproj` for **both** `x86` and `x64`, Release, with no
   new warnings. Confirm every new `.cs`/`.xaml` was actually added to the csproj: a missing entry
   compiles clean and silently omits the file (see D-4).
2. **Tests** — build and run `MSOfficeAIAssistant.Tests.exe`; expect exit 0. Register a new suite in
   `tests/Program.cs` for each new parser, registry or validator. Keep the logic COM-free.
3. **Install and smoke** — `install.cmd`, then `tools/verify.ps1`, then manual
   exercise in Word, Excel and PowerPoint — and specifically **Excel 2010**, whose docked-pane path
   and keyboard hook are the highest-regression-risk area in the product.
4. **Provider matrix** — exercise the feature against Mistral, Gemini, Groq **and** a Custom/local
   endpoint. Phase C's `ExtractionFailure` path is only meaningfully tested on a weak model.

### Release gates

**Build:** clean, correct version, no warnings.
**COM:** add-in and task-pane registration succeed; 32-bit Office verified; 64-bit Office verified.
**Per host:** ribbon loads · pane opens · chat works · context works · insert works · shutdown/restart stable.
**Excel additionally:** large and sparse sheet behavior; protected-sheet rejection; merged-cell handling.
**PowerPoint additionally:** image-only slides do not crash.
**Network:** valid key · invalid key · 401 · 404 · 429 · timeout · offline · TLS failure.
**Security:** API key absent from plaintext registry, logs, and source; context transmission matches
what the context bar declares.

---

## 9. Privacy model

The user must always be able to see what is sent. The UI states the current context scope, the active
provider and model, and that Office changes require approval.

- Context is sent **only** when enabled in the context bar.
- API keys are DPAPI-encrypted, never logged, never in error messages, never in source control.
- Chat history and the audit trail are local and DPAPI-encrypted.
- Never log full document context, full prompts, or full model responses by default.
- No telemetry leaves the machine.

---

## 10. Canonical terminology

**Office Host** — Word, Excel, PowerPoint · **Shared COM Add-in** — the extension model ·
**Connect** — COM entry point · **TaskPaneControl** — the ActiveX shim hosting the WPF chat ·
**ChatSidebar** — the WPF chat control · **Provider** — an AI backend implementing `IAIProvider` ·
**Orchestrator** — provider lifetime and streaming · **Session** — conversation, context and mode
state (Phase 0) · **Context** — Office content supplied to the AI · **Fragment** — one
location-tagged piece of context · **Source** — a resolvable Office location backing a claim ·
**Evidence Level** — the classification in §5.4 · **Skill** — a structured prompt template ·
**Domain Pack** — `general` or `railway` · **Tool** — a registered host capability ·
**Action** — one instance of a Tool with concrete parameters · **Plan** — an ordered list of steps ·
**Risk Level** — 0–3 per §5.2 · **WorkSession** — conversation + context + plan + actions + sources.

---

## 11. Change-control rules

1. Update source code first.
2. Update this SSOT whenever architecture, behavior, hosts, providers, security, registration or major
   features change.
3. Mark every feature `Implemented`, `Verified`, `Planned`, or `Not Implemented` (§3).
4. **Never mark a feature `Verified` from source inspection alone** — it requires a live run.
5. Preserve backward compatibility unless a breaking change is deliberately documented here.
6. Update `install.ps1` / `uninstall.ps1` / the Inno Setup scripts together with any COM change.
7. Update `tools/verify.ps1` whenever a host or registration contract changes.
8. Never add a provider-specific feature to the core abstraction without documenting its capability
   requirement.
9. Anything listed as out of scope in §1 requires an explicit revision of this document before any
   implementation begins.

---

## Appendix A — Resolved defects

### Office 2021 add-in failure — stale COM registration (RESOLVED 2026-08-21, severity CRITICAL)

**Symptom.** The add-in failed to load on Office 2021 and newer.

**Root cause — not a version incompatibility.** The CLSID `{2F8D4B61-…}` under `HKCU\Software\Classes`
had accumulated two `InprocServer32\<version>` subkeys:

| Subkey | Assembly identity | Status |
|---|---|---|
| `InprocServer32\0.4.0.0` | `MSOfficeAIAssistant, Version=0.4.0.0` | current |
| `InprocServer32\1.0.0.0` | `MistralOfficeAddin, Version=1.0.0.0` | **stale** |

The CLR COM activation path always selects the **highest** version subkey. `1.0.0.0` no longer exists
in the current assembly, so the loader threw
`FileLoadException: manifest definition does not match the assembly reference` — **HRESULT 0x80131040**.

The cause was `RegAsm /regfile`, which **merges** rather than replaces registry entries, compounded by
a loop that blindly stamped the current CodeBase onto every child key. The failure was therefore
**registration-history-specific, not version-specific**: a clean machine had no orphan; a machine
upgraded from a pre-0.4.0 build did.

**Fixes applied.**

1. `install.ps1` — delete the entire CLSID trees before importing the new regfile.
2. `install.ps1` — stamp CodeBase only on the parent key and the current version subkey, never on
   arbitrary children.
3. `src/Core/VersionDetector.cs` — discriminate Office versions by the second build-number component
   (`16.0.2xxxx+` = 2021, `16.0.1xxxx+` = 365/2019, below = 2016).
4. `installer/setup-x86.iss`, `setup-x64.iss` — run `regasm /unregister` before `/codebase`.

---

## Appendix B — Troubleshooting

**Diagnostic logs:** `%LOCALAPPDATA%\MSOfficeAIAssistant\addin.log`, fallback
`%TEMP%\MSOfficeAIAssistant.log`.

**1. Ribbon tab does not appear.** Usually a bitness mismatch or a disabled `LoadBehavior`.
Register the 32-bit DLL for 32-bit Office and the 64-bit DLL for 64-bit Office. Check
**File → Options → Add-ins → COM Add-ins → Go…** and confirm the add-in is checked. Verify
`HKCU\Software\Microsoft\Office\<App>\Addins\MSOfficeAIAssistant.Addin\LoadBehavior` is DWORD `3`. Also check
the Office *disabled items* and *crash-disabled* lists — `install.ps1` clears both.

**2. "Task pane factory is not available."** The host has not yet exposed `ICustomTaskPaneConsumer`,
or Office was launched in embedded/preview mode. Restart Office normally and open a document before
clicking **Open AI Chat**.

**3. API error 401.** The API key is wrong, expired, or missing. Regenerate it in the provider
console, then **Configure → paste → Test Connection → Save**.

**4. API error 429.** Provider rate limit. The client already retries with exponential backoff; for
sustained load, switch to a smaller model or raise your provider tier.

**5. Office 2021+ fails to load the add-in.** Stale COM registration — see Appendix A. Re-run
`install.cmd`, which now cleans the CLSID trees before registering.

**6. Office 2010.** Requires the .NET Framework 4.8 runtime. If the ribbon fails to render, confirm
the Visual Studio 2010 Tools for Office Runtime is present. Excel 2010 deliberately uses the docked-pane
host rather than the native CTP.

**7. Typing in the chat box edits the worksheet instead.** The Excel keyboard hook failed to install.
Close and reopen the pane; if it persists, restart Excel.

---

## Appendix C — Historical baseline

The product began as `MSOfficePlugin.rar` — a Mistral-only, non-streaming, `HttpWebRequest` +
`JavaScriptSerializer` COM add-in with a 2-button ribbon, registry-stored settings under
`HKCU\Software\MistralAIOffice`, a 12,000-character context cap, and simple single-response insertion
(`Selection.TypeText`, `ActiveCell.Value2`, a new PowerPoint text box).

Every one of those characteristics has since been replaced. The original requirements document
(`Mistral_Office_Addin_Requirements.md`) and the archive-era SSOT were deleted when this file was
created; their content is superseded by §2. Git history retains them if needed.

---

*End of SSOT. This document is the single reference for the AI Assistant for Microsoft Office. Any
expansion beyond what is described here requires change control under §11.*
