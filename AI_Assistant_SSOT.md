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

### 2.15 Design tokens (Phase 0.3)

`src/UI/Theme/Tokens.xaml` (color primitives + semantic brushes) and `src/UI/Theme/Controls.xaml`
(keyed Card/Badge/Chip/Button styles) are the first `ResourceDictionary` files in the repo. Both are
merged into `ChatSidebar.xaml` and `SettingsWindow.xaml` via `MergedDictionaries` in each file's own
`Resources` — there is no `App.xaml`/`Application` instance to host a global dictionary (this is a
COM add-in hosted through `ElementHost`, not a standalone WPF app), so the merge is per-root-element
by design. Every inlined hex literal in those two files was replaced 1:1 with a `{StaticResource}`
token reference (mechanical substitution — same colors, same properties, same element order — so the
`OfficeDockedPane` keyboard-hook HWND resolution in §2.3 is unaffected). `Controls.xaml` styles are
all `x:Key`'d, not `TargetType`-implicit, and none are wired into existing controls yet, so merging
it changes nothing visually; Phase A's visual design pass is expected to consume them.
`MarkdownHelper.cs` mirrors the same six colors it needs as frozen `SolidColorBrush` statics
(with a comment cross-referencing the `Tokens.xaml` keys) rather than loading the `ResourceDictionary`
at runtime — it sits on the streaming hot path (see D-7 resolution above), and per-call XAML resource
resolution has no reason to be on that path.

### 2.16 ElementHost DPI awareness — investigated, not implemented (Phase 0.3)

Per-Monitor v2 DPI handling was scoped as an *investigation* for Phase 0.3, not an implementation.
Findings:

- **DPI awareness is a process-level declaration**, made by whichever EXE calls
  `SetProcessDpiAwarenessContext` (or carries the equivalent application manifest) before any window
  is created. For this add-in, that process is **`EXCEL.EXE` / `WINWORD.EXE` / `POWERPNT.EXE`** — the
  Office host — not `MSOfficeAIAssistant.dll`. A COM in-process server loaded into that host inherits
  whatever DPI-awareness mode the host already declared; it cannot change it after the fact, and
  `app.config` (checked — `src/app.config`, currently only `<startup>`) has no DPI-related element
  that would do anything even if added, because `.config` files govern the CLR, not Win32 DPI
  awareness, and are read by whichever process loads the CLR (the Office host) regardless.
- Confirmed no DPI awareness is declared anywhere in this repo today: no manifest, no
  `<application><windowsSettings>` block, no `SetProcessDpiAwarenessContext`/`SetProcessDPIAware`
  call. Whatever the three hosting strategies render at currently comes entirely from the Office
  host process's own DPI-awareness mode (Office itself has been Per-Monitor-v2-aware since fairly
  recent 365 builds, System-DPI-aware on older ones) — this add-in has simply never had an opinion.
- Where a real fix *would* live, if this add-in ever needs to react to a DPI change independent of
  the host (e.g. dragging a floating `ChatFloatingWindow` between monitors with different scaling):
  **`ElementHost`/WinForms DPI-change handling**, not process-level awareness. Concretely:
  `OfficeDockedPane` (`Form`) and `ChatFloatingWindow` (`Form`) would need `AutoScaleMode` review and
  a `WM_DPICHANGED` handler that re-applies `ElementHost.Font`/scaling and asks the child WPF
  `HwndSource` to re-layout — WinForms `ElementHost` does not automatically rescale its hosted WPF
  content on a monitor-to-monitor DPI change the way a native WPF window does. The native-CTP path
  (`ICTPFactory.CreateCTP`, §2.3 strategy 2) is Office-managed chrome and is out of scope for any such
  fix regardless — only the `OfficeDockedPane` and `ChatFloatingWindow` fallback paths would need it.
- This is a self-contained slice (WinForms/ElementHost-specific, touches the same `Form` subclasses
  as the load-bearing keyboard hook) and was deliberately **not** bundled into Phase 0.3 — it has no
  dependency on the tokens/virtualization/markdown work done here and deserves its own careful pass
  given `OfficeDockedPane`'s fragility (§2.3).

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
| From document → briefing deck (doc-to-deck) | Implemented (Verified in slice) |
| Live verification across the full Office/bitness matrix | **Not Implemented** |
| Chat / Plan / Edit modes | Implemented (Phase A1: SessionMode gate hard-blocks RiskLevel ≥1 in Chat mode. Plan mode is now real, not a no-op alias of Edit: `AssistantSession.ProcessAssistantResponse` builds a `Planner`-produced `Plan` instead of populating `OfficeActions`, rendered as an editable, executable `PlanTemplate` card — reorder/skip/remove steps, per-step approve, run, rollback, wired to the already-tested `PlanExecutor`. Single-host only; `CrossHostPlanCoordinator` is not wired into chat, since a single chat session's `ActionExtractor` is already host-scoped) |
| Context bar, source citations, response cards | Implemented (Phase A2–A4: checkbox context scope + live host readout; DataTemplateSelector over Text/ActionPreview/Warning/Finding/Recommendation/Summary; paragraph/cell/sheet provenance tags in extracted text — click-to-navigate UI wiring not yet built) |
| Skills and domain packs | Implemented (Phase B1–B5: `SkillRegistry` loading `general` (9 skills) and `railway` (13 skills) JSON manifests; `AppendDomainPackRules` prompt composition; evidence levels on Finding cards; context-aware skill-chip promotion) |
| Unified action schema, tool registry, risk levels, verification, rollback | Implemented (Phase C0–C5 complete, 16/16 unit test suites passing) |
| Multi-step planner, cross-host workflows | Implemented (Phase D1–D4 complete: `Planner`, `PlanExecutor`, `CrossHostPlanCoordinator`, `WorkSession`; verified in `PlannerTests`, `PlanExecutorTests`, `CrossHostPlanCoordinatorTests`, `WorkSessionStoreTests`) |
| Web search / external grounding | Deferred (Post-Phase D) — opt-in BYOK client-side search |
| AI Pages, model routing, knowledge library, feedback capture | Not Implemented — deferred past Phase D |
| GSSMS / CMMS / IoT / connectors / RAG | **Out of scope** (§1) |

---

## 4. Known defects and structural debt

| ID | Severity | Issue |
|---|---|---|
| ~~D-1~~ | — | **RESOLVED.** `PowerPointController.InsertText` strips action XML without executing. `<powerpoint_actions>` XML is parsed on response completion into `ChatMessage.PowerPointActions`, and actions are reviewed and approved via dedicated UI cards with typed status and audit logging. |
| ~~D-2~~ | — | **RESOLVED.** `GetActivePresentation(bool)`/`GetOrCreateActiveSlide(bool)` split into honestly-named pairs with no boolean flag: `GetActivePresentation()`/`GetOrCreateActivePresentation()` and `GetActiveSlide()`/`GetOrCreateActiveSlide()`, sharing a private `*Core(bool)` implementation. Every call site now states its own mutation intent by name alone. |
| ~~D-3~~ | — | **RESOLVED.** `tools/verify.ps1` checked ProgID `MistralAI.Connect` while the add-in registered `MistralAI.Addin`, so the smoke check always reported FAIL and pointed at the deleted `register.cmd`. Now checks `MSOfficeAIAssistant.Addin` and recommends `install.cmd`. |
| ~~D-4~~ | — | **RESOLVED.** Deleted orphaned `OutlookController.cs` and unified controller dispatch under `IOfficeHostController`. |
| ~~D-5~~ | — | **RESOLVED.** All 4 allow-list sites unified to `ToolRegistry` (`PromptAssembler.cs`, `SpreadsheetActionParser.cs`, `PowerPointActionParser.cs`, `ExcelController.cs` execution dispatch via `ToolRegistry.Execute`). Eliminated 13-case hardcoded switch in `ExcelController`. |
| ~~D-6~~ | — | **RESOLVED.** Session orchestration extracted from `ChatSidebar.xaml.cs` into `src/Core/Session/` (`AssistantSession`, `StreamCoordinator`, `PromptAssembler`). View code-behind handles rendering and HWND routing only. Headlessly verified with dedicated session tests and Golden Master hash gate. |
| ~~D-7~~ | — | **RESOLVED (Phase 0.3).** `MessagesItemsControl` now virtualizes (`VirtualizingStackPanel.IsVirtualizing=True`, `VirtualizationMode=Standard`, `ScrollUnit=Pixel` — Recycling was tried first and produces visible scroll jitter because message bubbles vary widely in height, plain text vs. an action card with N sub-items; Standard avoids that at the cost of not reusing containers). `MarkdownHelper` (`src/UI/Helpers/MarkdownHelper.cs`) now renders incrementally while `ChatMessage.IsStreaming` is true — only newly-*completed* (newline-terminated) lines are parsed and appended; the trailing, possibly mid-marker line is deliberately held back. On stream end (`IsStreaming` flips false) it always does one full from-scratch re-parse, so final output is guaranteed identical to a single full parse regardless of chunk boundaries — verified by `tests/MarkdownHelperTests.cs`, including a fence-marker split mid-token. |
| ~~D-8~~ | — | **RESOLVED (Phase 0.3).** Dead `MarkdownToFlowDocumentConverter` deleted from `src/UI/Converters/MarkdownConverter.cs` (kept `BooleanToVisibilityConverter`, which lives in the same file and *is* used). Its sole dependency, `Markdig.Wpf`, is now unused and was removed from the csproj and `packages.config`; core `Markdig` remains as a live dependency for `WordMarkdownRenderer`. |
| ~~D-9~~ | — | **RESOLVED.** Added a `Translate` menu (9 language buttons) to `grpMore` in `Ribbon.xml`, each wired to `RibbonCallback.OnTranslate`. The degraded-mode fallback ribbon hardcoded in `Connect.cs` is deliberately minimal (2 buttons) and was left as-is. |
| ~~D-10~~ | — | **RESOLVED.** Deleted all four dead methods (`ExcelController.CreatePreviewDescription`, `ExcelController.WriteFormula`, `PowerPointController.GetPresentationOutline`, `PowerPointController.SetSpeakerNotes`). Verified each shared private helper (`GetSlideTitle`, `GetSectionName`, `AppendBounded`, `CleanMarkdown`) still has other live callers before removing the dead wrapper. |
| ~~D-11~~ | — | **RESOLVED.** `README.md` documented the non-existent `build.bat` and `register.cmd`; corrected to `install.cmd` plus a direct-MSBuild loop. |
| ~~D-12~~ | — | **RESOLVED.** Implemented `OleMessageFilter` (`CoRegisterMessageFilter` on STA thread) to handle Excel in-cell edit busy rejections (`0x800AC472` / `RPC_E_SERVERCALL_RETRYLATER`) with 10s retry window; previous filter restored on add-in disconnect. Added typed `SafeOfficeProbe` helpers. |
| ~~D-13~~ | — | **RESOLVED (Tier 1+2; Tier 3 deferred, non-blocking).** Every UI-reachable mutation method across `WordController`/`ExcelController`/`PowerPointController` now has an `Execute*` wrapper returning `HostOperationResult` (Tier 1, Phase C4: `ExecuteMoveSlide`/`ExecuteCreateSectionBeforeSlide`/`ExecuteRenameSectionInPlace`/`ExecuteSetSpeakerNotesInPlace`; Tier 2, this pass: `ExecuteAcceptRevisionsInSelection`/`ExecuteRejectRevisionsInSelection`/`ExecuteUndoLastChange` on Word, `ExecuteUndoLastAction`/`ExecuteInsertText` on Excel, `ExecuteUndo`/`ExecuteInsertText` on PowerPoint), each covered by headless tests in `HostOperationResultTests.cs`. Also removed 7 additional dead methods found during the sweep beyond D-10's original scope (`WordController.ApplyTrackChanges`/`ConvertSelectedMarkdownTableToWordTable`/`InsertMarkdownTableAtCursor`/`ConvertSelectedTextToWordTable`, `ExcelController.ApplySpreadsheetAction`, `PowerPointController.ApplyPowerPointAction`/`ExecutePowerPointAction`/`UndoLastChange`). The original blocking criterion — "Phase C5 verification cannot ship while any UI-reachable mutation swallows" — is satisfied: Phase C5 shipped and every mutation entry point now reports a structured result. Remaining bare `catch { }` swallows (~67, Word 10 / Excel 18 / PowerPoint 39) are internal read/probe fallbacks nested *inside* already-wrapped methods (COM cross-version property probing, best-effort range/target resolution) — not unreported top-level mutation failures. Converting them to the `SafeOfficeProbe<T>` pattern (D-12) for cleaner logging is left as a **Tier 3 follow-up**, tracked but not gating Phase A. |
| Note (Phase A forward guidance) | — | `src/UI/ChatSidebar.xaml.cs` calls several legacy controller methods directly (e.g. `_wordCtrl.InsertTextAtCursor`, `_excelCtrl.UndoLastAction`, `_pptCtrl.Undo`) instead of the `Execute*`/`HostOperationResult` wrappers above. Left as-is deliberately — this UI surface is being redesigned by Phase A (mode selector, response cards, approval flow), so rewiring it now would likely be partially thrown away. Phase A's action-approval UI should route through `Execute*` wrappers, not legacy methods. |

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
7. **Phase 0.3 — UI Foundation & Performance:** Design system (`Tokens.xaml` + `Controls.xaml`),
   re-enabled list virtualization, and incremental markdown rendering (D-7, D-8). Win32/ElementHost
   Per-Monitor v2 DPI handling was investigated (§2.16) and deliberately deferred to its own slice.
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

| 0.0 | **Golden Master Baseline:** Lift prompt assembly to pure static function; create headless test fixture recording prompt strings, action parsing DTOs, and audit serialization. Canonical SHA-256 (`88e58388...`) and `golden_master_baseline.txt` fixture committed to gate changes against byte-for-byte drift. | Verification Gate | Verified |
| 0.1 | **Extract Orchestrator:** Move prompt, streaming, and session logic into `src/Core/Session/` (`AssistantSession`, `PromptAssembler`, `StreamCoordinator`). View keeps rendering & HWND hooks only. | D-6 | Verified |
| 0.2 | **COM Resilience:** Implement `IOleMessageFilter` for Excel busy rejection (`0x800AC472`); implement typed `SafeOfficeProbe<T>` for 2010↔365 version probing; unswallow mutation errors. | **D-12**, §2.12 | Verified |
| 0.3 | **UI Foundation & Theme:** Design system `Tokens.xaml` + `Controls.xaml`; re-enabled list virtualization; incremental markdown. ElementHost Per-Monitor v2 DPI handling investigated (§2.16) and deliberately deferred — out of scope, host-process-level concern. | D-7, D-8 | Verified |
| 0.4 | **Controller Interface:** `IOfficeHostController` over the three controllers, replacing common dispatch; deleted orphaned `OutlookController` (D-4). | D-4 | Verified |
| 0.5 | **Provider Capabilities:** Add `StructuredOutput` / `ToolCalling` / `JsonMode` to `AICapabilities`; make `BuildPayload` extensible. | §2.5 | Verified |

**Exit:** identical responses, actions and audit entries before and after extraction; PowerPoint
structured actions visible in approval dialog; Excel in-cell edit rejections handled gracefully.

### Phase A — Copilot UX (Implemented & Verified in 6/6 New Unit Suites; low risk, no new mutation pathways)

- **A1** ✅ Chat/Plan/Edit selector, enforced in `AssistantSession` — Chat hard-blocks risk ≥1.
  `SessionMode` enum + `AssistantSession.IsActionAllowed(OfficeAction)`, wired as the first check in
  `ChatSidebar.ExecuteOfficeAction`, defaults to `Edit` (zero regression). `AssistantSessionModeTests`.
- **A2** ✅ Context bar: `ChkIncludeSelection`/`ChkIncludeCurrentFile` checkboxes mapped onto the
  unchanged `PromptContextScope` enum, plus a live host readout (`GetContextReadout()` per controller —
  `Sheet1!B2:B10`, `Slide 3 of 12`, `Section: X`/`Document: X`). The `ContextEngine`/`ContextFragment`
  abstraction named in the original plan was not built as a separate layer — the checkbox→enum mapping
  and per-controller readout getters deliver the same user-visible behavior more simply.
- **A3** ✅ Response cards via `ResponseCardTemplateSelector` (`src/UI/Cards/`): Text, ActionPreview,
  **Plan**, Warning, Finding, Recommendation, Summary — driven by a pure `ResponseCardCategoryClassifier`
  (`HasPlan` first, then `HasOfficeActions`, then a `**Warning:**`/`Warning:`-style content-prefix
  marker). Plan card added post-Phase-B once the chat-flow `PlanExecutor` wiring landed (see below) —
  step list, reorder/skip/remove, per-step approve, run, rollback. **Table** is still NOT a distinct
  card type — already handled inline by `MarkdownHelper` today with no distinct card chrome, so a
  dedicated Table card remains low-value. `ResponseCardCategoryTests`.
- **A4** ✅ Source-tag provenance (backend half only — click-to-navigate UI wiring is open): `.docx`
  paragraphs tagged `[¶N]`, `.xlsx` cells tagged with their real address and sheet name (resolved from
  `xl/workbook.xml`, falling back to an ordinal per-sheet on failure), Word excerpt labels carry a
  `~Paragraph N` line-based approximation. `AttachmentExtractorProvenanceTests`.
- **A5** ✅ Data-driven quick-prompt chips (`QuickPromptRegistry`, `Core.QuickPrompts` — deliberately
  **not** `Core.Skills`, since the real Skill registry is Phase B and doesn't exist yet). Same 6 prompts,
  same text, `HostFilter` restores the original PowerPoint-only visibility of "Build deck".
  `QuickPromptRegistryTests`.
- **A6** ✅ Conversation history UI (`ConversationHistoryWindow` over `ConversationStore.ListSessions()`)
  and a real action-history panel (`ActionHistoryWindow` over `ActionAuditStore.GetRecent`), both
  replacing prior `MessageBox` dumps. `ConversationStoreSessionTests`.
- **A7** ✅ Status indicator (Ready / Thinking / Reading / Awaiting approval / Applying / Verifying /
  Done / Failed / Cancelled — 9 of 10 named states, each tied to a real code transition) plus
  `AutomationProperties.Name`/`LiveSetting="Polite"` accessibility labels. **Planning** (the 10th state)
  is still not implemented — the Plan card's Run/Approve handlers call `PlanExecutor` synchronously on
  the UI thread with no progress-callback-driven intermediate status update, so there's no real
  "actively planning" moment distinct from the surrounding Applying/Verifying states to hang a Planning
  status on. `AssistantStatusTests`.

**Exit:** the user can see exactly what context is sent, switch modes, and receive structured cards.
Mutation behavior unchanged except the new Chat-mode hard block (additive safety, not a regression).
**One item from the original two is now closed:** Plan mode is wired to the Phase D `PlanExecutor`
(see the Chat/Plan/Edit modes row above and the Phase D integration entry below) — A3's Plan card and
this wiring shipped together, after Phase B, as a separate follow-up slice. **Click-to-navigate on
source tags (A4) remains open.**

### Phase B — Skills and domain packs (Implemented & Verified in 5/5 New Unit Suites; low risk, parallel with A)

- **B1** ✅ `src/Core/Skills/` — `Skill` (plain Newtonsoft.Json model, snake_case properties),
  `SkillRegistry.LoadPack`/`GetAllPacks` (embedded JSON manifests, case-insensitive pack name, never
  throws/never null), `ConfigManager.DomainPack` (persisted, defaults `general`), Settings pack
  selector. `SkillRegistryTests`.
- **B2** ✅ Real catalog content: `general.json` (9 skills — Official Letter, Minutes of Meeting,
  Inspection Report, Technical Note, Management Summary, Root Cause Analysis, Management Dashboard,
  Improve Official Language, Document Comparison) and `railway.json` (all 9 general + 4 railway-
  specific — DRM Briefing, Failure Analysis & Pareto, Substation/OHE Asset Health, Deficiency
  Tracker — 13 total). Every skill whose output could misrepresent source data carries an explicit
  anti-fabrication instruction (preserve values exactly, flag missing information, distinguish
  evidence from inference). All 10 railway terms (Depot, Substation, OHE, TRD, PM/CM, Breakdown, DRM,
  Sr.DEE, SSE, JE) used accurately across the 4 railway-specific skills.
- **B3** ✅ `PromptAssembler.AppendDomainPackRules(systemPrompt, domainPack)` — appends railway
  vocabulary guidance when `ConfigManager.DomainPack == "railway"`, no-op for `general`; wired into
  `AssistantSession.PreparePayloadAsync`. `QuickPromptRegistry.GetRibbonPrompts()` consolidates the 10
  simple `RibbonCallback` hardcoded prompts into the same data-driven catalog Phase A5 introduced.
  **"Selected skill" composition is NOT wired** — no skill-selection UI exists in the live chat flow
  yet (B5 only surfaces skill chips as one-shot prompts, not as a persisted "active skill" the AI
  stays aware of turn-to-turn); this remains open for a future slice. First deliberate change to
  `PromptAssembler` since the Golden Master gate was established — hash updated following the
  documented fixture-diff procedure, confirmed purely additive both times (once for the initial
  change, once more after a terminology-accuracy fix caught in review).
- **B4** ✅ Evidence levels (`DirectlyObserved`/`Calculated`/`StrongInference`/`PossibleInference`/
  `InsufficientEvidence`) rendered on Finding cards (A3), classified from real Phase-A4 citation
  patterns in the message text (`[¶N]`, `~Paragraph N`, `Sheet1!B7`, `B7=value`, `Slide N of M`) as
  the primary signal, an explicit bracketed structured tag as a secondary signal, and
  `InsufficientEvidence` as the honest default — **never from the model's self-reported confidence**,
  per the plan's explicit requirement.
- **B5** ✅ `SkillPicker.SelectChips` surfaces up to 3 skill-derived chips through the existing A5
  quick-prompt row (not a new picker UI — out of scope for this slice), with context-aware promotion:
  an Excel sheet whose column headers contain failure-related keywords (failure/breakdown/fault/
  defect) promotes `failure_analysis_pareto` to the front when the railway pack is active.

**Exit:** the same prompt under `general` and `railway` produces an appropriately different register
(railway vocabulary guidance is real and wired); domain rules demonstrably suppress invented
references (every content-generating skill carries an explicit anti-fabrication instruction).

### Phase C — Safe execution (Implemented & Verified in 16/16 Unit Suites)

- **C0** `HostOperationResult` structured execution envelopes with HRESULT capture and error classification.
- **C1** `src/Core/Actions/OfficeAction.cs` implementing §5.3 unified schema with backward-compatible adapters.
- **C2** `src/Core/Actions/ToolRegistry.cs` single-source tool registry for actions, risk levels, and prompt generation (D-5).
- **C3** `src/Core/Actions/ActionExtractor.cs` native/embedded extraction and unified single-card UI (`ChatSidebar.xaml`).
- **C4** `src/Core/Actions/ActionVerifier.cs` pre/post verification engine, Excel calculation error literal detection (`#REF!`, `#VALUE!`), and busy state (`0x800AC472`) classification.
- **C5** `src/Core/Actions/RollbackExecutor.cs` & Audit v2: formula-preserving `BeforeState` capture, programmatic inverse execution, strict LIFO batch unwinding, capacity limits (5,000 cells), and additive `ActionAuditStore`.

**Exit:** All structured actions across Word, Excel, and PowerPoint are unified, risk-gated, verified, rollbackable in strict LIFO order, and audited with full forensic provenance.

### Phase D — Agentic plan-then-execute (medium–high risk) (Implemented & Verified in 4/4 Unit Suites)

- **D1** `src/Core/Planning/` — Planner producing ordered `PlanStep`s bound to Tool Registry entries.
  Editable before any risk ≥1 action runs.
- **D2** Execution state machine: Queued / Running / Awaiting approval / Paused / Completed / Failed /
  Cancelled / Rolled back. Progress, cancellation, partial-failure recovery, continue-from-step-N.
- **D3** Cross-host workflows (Excel analysis → Word report → PowerPoint briefing) as sequential
  approved actions over `IOfficeHostController`. No silent cross-application mutation.
- **D4** `WorkSession` — persist Conversation + Context + Plan + Actions + Sources as one reopenable unit.

**Exit:** "Analyze these failures, build a dashboard, then draft a briefing" yields an editable plan
that executes across hosts under approval with a full audit trail.

**Post-Phase-B follow-up — chat integration.** D1/D2's backend (`Planner`, `PlanExecutor`) is now wired
into the live chat send flow (`AssistantSession.ProcessAssistantResponse` + `ChatSidebar`'s `PlanTemplate`
card, see the A1/A3 entries above) — Plan mode produces a real, editable, executable step list instead
of behaving like Edit mode. This integration is **single-host only**: `CrossHostPlanCoordinator` (D3) is
NOT wired into chat, since a single chat session's `ActionExtractor` only ever proposes same-host
actions today, so there is no live multi-host plan for it to coordinate yet. `WorkSession` (D4)
persistence is also not wired into chat — an active `Plan` lives only in memory on its `ChatMessage`
(`[JsonIgnore]`), not saved/reloaded across sessions. Both remain genuinely open follow-ups, not
silently dropped.

### Deferred past Phase D

AI Pages · model routing (fast vs reasoning) · local knowledge library · document-comparison skill ·
local feedback capture. Documented as **not activated**; requires §11 change control to start.

### Feature backlog — ranked

Assessed for fit against this add-in's actual constraints (COM, Office 2010+, BYOK, provider-neutral,
confirm-before-write), not against Copilot's feature list.

| Rank | Feature | Verdict |
|---|---|---|
| 1 | **Doc-to-deck** | **Verified**. Attachment-sourced briefing generator, slide outline parser, preview dialog, and view-state guards active. |
| 2 | **Word Review → native comments** | **Implemented.** `WordController.ExecuteAddComment` wraps `Document.Comments.Add(Range, Text)`, risk-gated and tested (`HostOperationResultTests`). |
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
10. **Validate regression guards by deliberate negative mutation.** A regression test or safety guard must
    be validated by temporarily reintroducing the defect it claims to prevent and confirming that the test
    suite fails. Testing helper methods in isolation does not validate adapter call sites.

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
