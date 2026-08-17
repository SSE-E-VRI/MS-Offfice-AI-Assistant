# AI Assistant — Single Source of Truth (SSOT)

**Document:** `AI_Assistant_SSOT.md`  
**Version:** 0.0.0  
**SSOT Date:** 2026-08-17  
**Project:** AI Assistant  
**Primary artifact analyzed:** `MSOfficePlugin.rar`  
**Reference capability documents:** Microsoft 365 Copilot Word, Excel and PowerPoint feature notes uploaded with the project

---

## 0. Purpose

This document is the authoritative working description of the **AI Assistant** project and its required v0.0.0 product baseline.

It separates:

1. **Implemented in the analyzed source archive**
2. **Implemented but requiring live verification**
3. **Design/review recommendations not yet implemented**
4. **Target capabilities / future roadmap**
5. **Microsoft 365 Copilot reference capabilities**

Future development agents should use this file before modifying architecture, interfaces, provider behavior, Office-host behavior, registration, or feature scope.

## Product Identity — Mandatory v0.0.0 Baseline

- **Product name:** AI Assistant
- **Version:** 0.0.0
- **Supported products:** Microsoft Word, Microsoft Excel, Microsoft PowerPoint only
- **AI providers:** Mistral AI, Google Gemini, Groq, Custom API
- **File inputs:** PDF, Word, Excel, PowerPoint and images
- **Developer credit:** Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302

These are product requirements for the updated SSOT. They take precedence over legacy Mistral-only product naming in the analyzed source.

## Versioning Rule

For the v0.0.0 product baseline, the displayed application version is **0.0.0**. Build metadata, assembly/file version and About/Settings version display should remain consistent with the product release version unless a technical build-number suffix is required by the build system.

### SSOT precedence

When information conflicts:

1. Current source code in the project repository/archive
2. Current tests and verified runtime behavior
3. This SSOT
4. Design documents/review notes
5. Copilot capability reference documents
6. General assumptions

Do not silently convert a planned feature into an implemented feature.

---

# 1. Executive Summary

The analyzed project is a **classic shared C# COM add-in** for desktop Microsoft Office.

### Supported Office products

The product scope is intentionally limited to:

- Microsoft Word
- Microsoft Excel
- Microsoft PowerPoint


### AI provider options

The product shall support four provider choices:

1. **Mistral AI** — existing provider behavior retained.
2. **Google Gemini** — new provider.
3. **Groq** — new provider.
4. **Custom API** — user-configurable OpenAI-compatible or documented HTTP chat-completions-style endpoint.

All providers must be accessed through a common provider abstraction. The user selects the active provider from Settings. Credentials remain local and provider-specific configuration must not leak between providers.

### Office integration model

- `IDTExtensibility2` shared COM add-in
- Ribbon XML
- Custom Task Pane
- WinForms ActiveX-compatible task-pane control
- Office object model accessed through `dynamic`
- No Office PIA dependency
- Per-user HKCU registration
- AnyCPU target
- .NET Framework 4.x compiler/runtime strategy

### Current user-facing features

- Mistral AI Ribbon tab
- Chat pane
- Multi-turn chat history
- Optional current-document/workbook/presentation context
- Insert last response into current Office host
- Clear chat
- Settings dialog
- DPAPI-encrypted API key storage
- Model/base URL/timeout/system-prompt settings
- Test Connection
- Provider selection: Mistral AI / Gemini / Groq / Custom API
- Provider-specific API key, base URL and model configuration
- Upload/import source documents: PDF, Word, Excel, PowerPoint and images
- Background HTTP request with UI-thread marshaling
- Basic error handling for authentication, endpoint, rate-limit and network failures

### Product identity and credit

- **Product name:** AI Assistant
- **Version:** 0.0.0
- **Developer credit:** Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302

The developer credit shall be displayed in the About/Settings area and may also appear in the application footer or Help/About command, subject to the final UI design.

---

# 2. Source Inventory

## 2.1 Project archive

`MSOfficePlugin.rar`

The archive contains:

### Documentation

- `Design Plan.md`
- `Design Review.md`
- `mistral-office-addin/README.md`

### Source

- `AssemblyInfo.cs`
- `ComInterfaces.cs`
- `Connect.cs`
- `RibbonXml.cs`
- `SettingsStore.cs`
- `MistralClient.cs`
- `TaskPaneControl.cs`
- `SettingsForm.cs`

### Build/deployment

- `build.cmd`
- `register.cmd`
- `unregister.cmd`

### Verification/tools

- `tools/verify.ps1`
- `tools/DumpInterfaces.ps1`

### Binary

- `bin/MistralOfficeAddin.dll`

---

# 3. Architecture — Current State

```text
Microsoft Word
Microsoft Excel
Microsoft PowerPoint
        |
        v
IDTExtensibility2 / Connect
        |
        +--------------------+
        |                    |
        v                    v
   Ribbon XML          Custom Task Pane
                             |
                             v
                     TaskPaneControl
                             |
             +---------------+----------------+
             |               |                |
             v               v                v
        Word context    Excel context    PowerPoint context
             |               |                |
             +---------------+----------------+
                             |
                             v
                    ChatMessage history
                             |
                             v
                     AI Orchestrator
                             |
                    AI Provider Interface
                             |
       +---------------------+---------------------+----------------+
       |                     |                     |                |
       v                     v                     v                v
 MistralProvider       GeminiProvider        GroqProvider    CustomApiProvider
       |                     |                     |                |
       +---------------------+---------------------+----------------+
                             |
                         HTTPS APIs
```

## 3.1 Architectural principles

- One shared add-in DLL for multiple Office hosts.
- Avoid Office PIAs by hand-declaring required COM interfaces.
- Use `dynamic` for host object-model access.
- Avoid Visual Studio as a build requirement.
- Compile with the .NET Framework C# compiler.
- Keep API credentials local to the Windows user profile.
- Perform network operations away from the Office UI thread.
- Keep the initial implementation dependency-light.

---

# 4. Host Support Matrix

| Capability | Word | Excel | PowerPoint |
|---|---:|---:|---:|
| COM add-in connection | YES | YES | YES |
| Ribbon tab | YES | YES | YES |
| Chat task pane | YES | YES | YES |
| Current content context | YES | YES | YES |
| Insert AI response | YES | YES | YES |
| Settings | YES | YES | YES |
| AI chat | YES | YES | YES |

---

# 5. COM / Office Integration

## 5.1 Main COM add-in

The product-facing name is **AI Assistant**. The analyzed source currently uses legacy Mistral-specific COM class/ProgID names; renaming those identifiers is a separate implementation task and must preserve COM registration compatibility.

Class:

`MistralOfficeAddin.Connect`

Key attributes:

- `ComVisible(true)`
- GUID: `{2F8D4B61-7C3E-4A59-9B2D-6E1F0A3C5E78}`
- ProgID: `MistralAI.Connect`

Implements:

- `IDTExtensibility2`
- `IRibbonExtensibility`
- `ICustomTaskPaneConsumer`

## 5.2 CTP control

Class:

`MistralOfficeAddin.TaskPaneControl`

Key attributes:

- `ComVisible(true)`
- `ClassInterface(ClassInterfaceType.AutoDual)`
- GUID: `{9C4E7A15-2D6B-4F83-B5C9-7A2E1D4F6B83}`
- ProgID: `MistralAI.ChatPane`

Implements:

- `IObjectSafety`

The source contains ActiveX/CTP registration logic and `register.cmd` also adds the relevant control categories.

## 5.3 Office interface GUIDs

The source declares:

| Interface | GUID |
|---|---|
| `IDTExtensibility2` | `{B65AD801-ABAF-11D0-BB8B-00A0C90F2744}` |
| `IRibbonExtensibility` | `{000C0396-0000-0000-C000-000000000046}` |
| `ICustomTaskPaneConsumer` | `{000C033D-0000-0000-C000-000000000046}` |
| `ICTPFactory` | `{000C033E-0000-0000-C000-000000000046}` |
| `_CustomTaskPane` | `{000C033B-0000-0000-C000-000000000046}` |
| `IObjectSafety` | `{CB5BDC81-93C1-11CF-8F20-00805F2CD064}` |

These are source-declared COM contracts and must be validated against the target Office type library during live verification.

---

# 6. Ribbon

The current Ribbon XML uses:

`http://schemas.microsoft.com/office/2006/01/customUI`

Current tab:

**Mistral AI**

Current group:

**AI Assistant**

Current buttons:

1. **Chat Pane**
   - Callback: `OnChatButtonClick`
2. **Settings**
   - Callback: `OnSettingsButtonClick`

The Ribbon is intentionally minimal.

---

# 7. Custom Task Pane

## 7.1 Current UI

The task pane contains:

- Host-aware title
- Settings link
- Include document context checkbox
- Scrollable chat area
- Multiline input box
- Send
- Insert
- Clear
- Status text

## 7.2 Keyboard behavior

- `Ctrl+Enter` → Send
- `Shift+Enter` → Send
- `Escape` → clears input

Note: Escape currently **clears the input**; it does not actually close the task pane despite the source comment.

This distinction is authoritative for current behavior.

---

# 8. Chat Engine

## 8.1 Message model

```text
ChatMessage
  Role
  Content
```

Roles currently used:

- `system`
- `user`
- `assistant`

## 8.2 History

- Maximum retained history: `30` messages.
- Oldest messages are removed when the limit is exceeded.
- The current history snapshot is sent with every request.

## 8.3 System prompt

Default:

> You are a helpful AI assistant embedded in Microsoft Office. Be concise, accurate and practical.

A configured system prompt may be replaced/edited through Settings.

When document context is enabled, the context is appended inside:

```text
<context>
...
</context>
```

with host-specific instructions.

---

# 9. Office Context Extraction

## 9.1 Word

Current source extracts:

`ActiveDocument.Content.Text`

Processing:

- carriage returns normalized
- paragraph/end markers cleaned
- context capped

## 9.2 Excel

Current source reads:

`ActiveSheet.UsedRange.Value2`

Limits:

- up to 120 rows
- up to 40 columns
- overall context cap of 12,000 characters

Values are serialized as tab-separated rows.

## 9.3 PowerPoint

Current source reads text from shapes across the active presentation.

Limits:

- up to 80 slides
- text-only extraction
- overall context cap of 12,000 characters

## 9.4 Context hard limit

Current source constant:

`MaxContextChars = 12000`

This is a **character** limit, not a token limit.

Future token-aware context management must not be described as implemented until added.

---

# 10. Insert / Write-Back Behavior

## Word

Uses the current selection:

`Selection.TypeText(text)`

Result is inserted at the current selection.

## Excel

Writes to:

`ActiveCell.Value2`

This replaces the active cell value.

## PowerPoint

Creates a new horizontal text box on the current slide and places the response into it.

### Important design implication

The current implementation is not a general-purpose Office editing engine. It has a simple **single-response insertion** capability.

It does not currently provide:

- tracked/reversible edits
- range-aware editing
- multi-cell transformation planning
- style-preserving rewrite
- paragraph-level Word editing
- slide-layout-aware editing
- workbook-wide edit planning

---

# 11. AI Provider Architecture — Required

The existing `MistralClient` must be refactored behind a provider-neutral contract. The Office UI and orchestration layer must not call a vendor-specific client directly.

## 11.1 Provider options

| Provider | Required | Configuration |
|---|---|---|
| Mistral AI | YES | API key, base URL, model, timeout |
| Google Gemini | YES | API key, API base/endpoint, model, timeout |
| Groq | YES | API key, base URL, model, timeout |
| Custom API | YES | API key if required, custom base URL/endpoint, model, timeout, optional headers |

## 11.2 Common provider contract

Conceptually the provider layer shall expose:

- `SendChatAsync(...)`
- `TestConnectionAsync(...)`
- `GetModelsAsync(...)` where supported
- provider metadata/capabilities
- normalized response/error types

The orchestrator owns conversation history, document/image context, cancellation, retries where appropriate, and UI marshaling. Providers own HTTP/API-specific serialization and authentication.

## 11.3 Custom API

Custom API must support configurable endpoint and model information rather than assuming the Mistral endpoint. The first compatibility target should be an OpenAI-compatible chat-completions-style HTTP interface. Provider-specific headers and authentication behavior must be configurable only to the extent required by the implementation and must be stored securely.

# 12. Document and Image Upload / Context

The AI Assistant shall provide an **Upload Document / Image** option in the chat interface. Supported input types:

- PDF (`.pdf`)
- Microsoft Word (`.doc`, `.docx`)
- Microsoft Excel (`.xls`, `.xlsx`)
- Microsoft PowerPoint (`.ppt`, `.pptx`)
- Images (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.gif`, `.webp` where supported by the selected provider)

## 12.1 Upload behavior

1. User selects one or more supported files.
2. The add-in validates extension and file size before processing.
3. Text/table/slide content is extracted where possible.
4. Images are retained as image inputs when the selected provider supports multimodal input; otherwise an appropriate extraction/fallback path is used.
5. The extracted content and/or file representation is passed to the AI orchestration layer.
6. The chat UI displays the attached files and their processing status.
7. The original local files are not uploaded until the user sends a request requiring them.

## 12.2 Security and privacy

- Never execute uploaded files.
- Do not permanently copy uploaded documents into the project directory.
- Use temporary storage only when required by an API or extraction library.
- Remove temporary material after processing where practical.
- Do not transmit a document/image to a provider unless the user has initiated an operation that requires it.
- Provider-specific file upload APIs must remain encapsulated in the provider implementation.

## 12.3 Office-host context versus uploaded files

The current-document/workbook/presentation context feature remains separate from explicit file upload. A user can either ask about the active Office document or attach external PDF/Word/Excel/PowerPoint/image files to the conversation.

# 11. Mistral API Layer

Class:

`MistralClient`

## 11.1 Current API operations

### Chat

`POST {BaseUrl}/chat/completions`

### Model test

`GET {BaseUrl}/models`

## 11.2 Request behavior

Current chat request includes:

- model
- messages
- temperature = `0.7`
- stream = `false`

### Streaming

**Not implemented.**

The current architecture explicitly uses non-streaming responses.

## 11.3 HTTP

Uses:

`HttpWebRequest`

Headers include:

- `Content-Type: application/json`
- `Accept: application/json`
- `Authorization: Bearer <API key>`
- User-Agent: `AIAssistant/0.0.0`

## 11.4 JSON

Uses:

`System.Web.Script.Serialization.JavaScriptSerializer`

The source explicitly sets:

`MaxJsonLength = int.MaxValue`

Response parsing expects the chat-completions structure:

```text
choices[0].message.content
```

---

# 12. TLS

The source attempts to enable:

- TLS 1.2 via numeric value `3072`
- TLS 1.3 via numeric value `12288`

The project documentation requires a modern enough .NET Framework/runtime and operating-system TLS stack for communication with the Mistral endpoint.

Live validation is required on the oldest supported target machine.

---

# 13. Settings and Credential Storage

Settings registry location:

`HKCU\Software\MistralAIOffice`

Stored values:

- `ApiKeyEnc`
- `Model`
- `BaseUrl`
- `SystemPrompt`
- `TimeoutSeconds`
- `IncludeContextByDefault`

## 13.1 API key protection

The API key is protected with Windows DPAPI:

`DataProtectionScope.CurrentUser`

Entropy:

`MISTRAL`

The API key is therefore not stored as plaintext by the settings implementation.

## 13.2 Default settings

| Setting | Current default |
|---|---|
| Active Provider | `Mistral AI` |
| Model | `mistral-small-latest` |
| Base URL | `https://api.mistral.ai/v1` |
| Timeout | 60 seconds |
| Include context | enabled |
| Temperature | 0.7 |
| Max history | 30 messages |
| Max context | 12,000 characters |
| Streaming | disabled |

---

# 14. Error Handling

Current API handling explicitly recognizes:

- HTTP 401 → API key problem
- HTTP 404 → base URL/model problem
- HTTP 429 → rate limiting, including `Retry-After` where available
- Other HTTP failures
- Network failures
- JSON/response parsing failures

The chat UI shows an error bubble instead of throwing the API error into the Office host.

---

# 15. Threading Model

## UI thread

Office object-model operations are performed on the Office/UI thread.

Examples:

- Reading document context
- Reading workbook UsedRange
- Reading presentation shapes
- Inserting AI output

## Worker thread

HTTP requests run through the thread pool.

## Return path

The source captures a `SynchronizationContext` and posts the result back to the UI context.

A `BeginInvoke` fallback exists when no synchronization context is available.

---

# 16. Build System

Build command:

`build.cmd`

Compiler:

`.NET Framework csc.exe`

References include:

- System
- System.Core
- System.Drawing
- System.Windows.Forms
- System.Web.Extensions
- System.Security
- Microsoft.CSharp

Target:

- library
- AnyCPU
- optimized
- warnings as errors

## Strong-name status

The build script now attempts to generate:

`MistralAI.snk`

and compile with:

`/keyfile:MistralAI.snk`

Therefore the earlier design-review concern about missing strong-name support is **addressed in the current source**, subject to live verification that `sn.exe` is actually available on the build machine.

---

# 17. Registration

`register.cmd` performs per-user registration.

Target registry area:

`HKCU`

It registers:

- main COM add-in
- task-pane control
- 32-bit registration path
- 64-bit registration path
- Word add-in key
- Excel add-in key
- PowerPoint add-in key

Office add-in registration uses:

`LoadBehavior = 3`

No administrator elevation is intended.

---

# 18. Unregistration

`unregister.cmd` removes:

- Word add-in registration
- Excel add-in registration
- PowerPoint add-in registration
- main COM CLSID
- task-pane CLSID
- main ProgID
- task-pane ProgID
- 32-bit equivalents

The unregistration script should be kept synchronized with any future typelib or additional COM registrations.

---

# 19. Verification

Current verification script:

`tools/verify.ps1`

It attempts to launch:

- Word
- Excel
- PowerPoint

It checks for:

`MistralAI.Connect`

and whether the COM add-in is connected.

## Verification status

The archive contains verification tooling, but the archive itself does not prove a successful live run on all target Office versions.

Therefore:

**Source-level support ≠ runtime-certified support.**

A release should require live verification on representative:

- Office 2010 32-bit
- modern Office 64-bit
- Word
- Excel
- PowerPoint

---

# 20. Design Review Status — Reconciled

The included Design Review identified three critical issues and several recommendations.

## 20.1 CTP ActiveX registration

### Review status

Originally identified as a showstopper.

### Current source status

**Addressed in source.**

Evidence in current source:

- `TaskPaneControl` is COM-visible.
- Task pane has a GUID and ProgID.
- `IObjectSafety` is implemented.
- `ComRegisterFunction` / `ComUnregisterFunction` exist.
- `register.cmd` creates ActiveX-related registry categories.
- `Connect` calls `CreateCTP("MistralAI.ChatPane", ...)`.

### Remaining requirement

Live verification is still mandatory.

---

## 20.2 Strong-name signing

### Review status

Originally identified as missing.

### Current source status

**Addressed in build script.**

`build.cmd`:

- looks for `sn.exe`
- generates `MistralAI.snk`
- uses `/keyfile:MistralAI.snk`

### Remaining requirement

Verify the build actually produces a strong-named assembly on the target build environment.

---

## 20.3 JavaScriptSerializer size issue

### Review status

Originally identified as a critical issue.

### Current source status

**Addressed for JSON size.**

The serializer sets:

`MaxJsonLength = int.MaxValue`

### Still not implemented

- SSE streaming
- incremental token display
- streaming cancellation

---

# 21. Remaining Technical Gaps / Risks

These are the highest-priority items visible from the analyzed source.

## P0 — Runtime validation of CTP registration

The CTP mechanism is complex because the task pane control must be discoverable as the required COM/ActiveX control.

**Action:**

- Build
- Register
- Launch Word/Excel/PowerPoint
- Confirm `CreateCTP` succeeds
- Confirm `ContentControl` is the expected `TaskPaneControl`
- Confirm pane remains stable during Office startup/shutdown

---

## P0 — 32-bit/64-bit registration verification

The scripts attempt both registry views, but this must be tested with actual Office bitness.

**Required matrix:**

- 32-bit Office on 64-bit Windows
- 64-bit Office on 64-bit Windows
- if retained as a claim, 32-bit Windows

---

## P1 — Use the declared `_CustomTaskPane` / `ICTPFactory` interfaces directly

`ComInterfaces.cs` declares `ICTPFactory` and `_CustomTaskPane`, but `Connect.cs` currently stores the factory as `object` and uses `dynamic`.

This is inconsistent with the explicit COM-interface design.

### Preferred direction

Use the declared interfaces directly where practical:

```text
ICustomTaskPaneConsumer
        |
        v
ICTPFactory
        |
        v
_CustomTaskPane
```

This reduces late-binding ambiguity for the CTP layer.

---

## P1 — 32-bit Windows build script fallback

`build.cmd` and `register.cmd` primarily expect the 64-bit .NET Framework tool path and then conditionally handle 32-bit paths.

The 32-bit-Windows case should be explicitly tested and, if still supported, made deterministic.

---

## P1 — Model discovery instead of hard-coded model assumptions

The Settings dialog contains known model names, but the model field is editable and the client already has `/models` access.

Recommended direction:

- retrieve available models
- populate model selector dynamically
- preserve manual model entry
- avoid treating old/deprecated model names as guaranteed

Do not hard-code provider-specific model availability into the core architecture.

---

## P1 — Provider abstraction

The analyzed source implementation is Mistral-specific.

There is currently no provider abstraction such as:

```text
IAIProvider
AIRequest
AIResponse
ProviderCapabilities
AuthenticationMode
```

Therefore Gemini, Groq and Custom API cannot yet be added cleanly without refactoring `MistralClient` and the settings model.

### Architectural target

```text
Office UI
   |
   v
AI Orchestrator
   |
   v
IAIProvider
   |
   +---- MistralProvider
   +---- GeminiProvider
   +---- GroqProvider
   +---- CustomApiProvider
```

Provider-specific authentication, endpoints, model catalogs and capabilities should remain below the abstraction boundary.

---

## P1 — Context/token budgeting

Current context management is character-based.

Future architecture should introduce:

```text
ContextCollector
ContextNormalizer
TokenBudget
ContextTruncator
PromptAssembler
```

This is especially important for large Word documents, large Excel UsedRanges and long PowerPoint decks.

---

## P1 — Office edit safety

Current Insert behavior writes directly into the active document/cell/slide.

Future edit-mode functionality should support:

- preview
- plan
- confirmation
- reversible changes where feasible
- scoped target selection
- structured operation results

---

# 22. Microsoft 365 Copilot Capability Reference

The three uploaded Copilot reference documents are treated as **target capability references**, not proof that the current add-in implements those features.

## 22.1 Word reference capabilities

The uploaded Word reference describes:

- prompt-based drafting
- structured document generation
- Word Agent
- inline editing/refinement
- direct document editing
- writing suggestions
- text-to-table
- document summaries
- content Q&A
- citation display
- audio overview
- recent activity summaries
- work-context grounding
- chat-only mode
- model selection
- memory

The source specifically describes direct/reversible inline editing and document-grounded drafting as Copilot capabilities. fileciteturn0file2L3-L17

### Current project parity

Current add-in:

- Chat: YES
- Context Q&A: YES
- Basic insertion: YES
- Full drafting engine: NO
- Inline edit engine: NO
- Track/reversible edit engine: NO
- Summarization workflow: NO dedicated command
- Citations: NO
- Audio overview: NO
- Work-context grounding: NO
- Model picker: PARTIAL, Mistral-only
- Memory: NO persistent AI memory

---

## 22.2 Excel reference capabilities

The uploaded Excel reference describes:

- worksheet editing
- formula generation
- chart/PivotTable/shape creation
- insight detection
- sorting/filtering
- sentiment analysis
- financial modeling
- workbook import
- PDF extraction
- web search/citations
- federated connectors
- Python integration
- Work IQ context
- Edit Mode
- Plan Mode
- Chat Only Mode
- Custom Skills
- Agent Mode
- model selection
- workbook rules/guidelines

These capabilities are explicitly described in the uploaded Excel reference. fileciteturn0file1L8-L24 fileciteturn0file1L49-L75

### Current project parity

Current add-in:

- Workbook context extraction: YES
- Ask questions about workbook content: YES
- Write response to active cell: YES
- Formula generation: possible through free-form chat, but no dedicated tool
- Chart creation: NO
- PivotTable automation: NO
- Workbook restructuring: NO
- Plan mode: NO
- Edit mode: NO
- Agent mode: NO
- Python integration: NO
- Web search: NO
- Citations: NO
- Connectors/MCP: NO
- Workbook rules: NO

---

## 22.3 PowerPoint reference capabilities

The uploaded PowerPoint reference describes:

- slide deck generation
- document-to-presentation
- Researcher-to-PowerPoint
- PowerPoint Agent
- conversational Edit Mode
- standardize format
- control length/tone/style/images
- brand kit integration
- review presentation
- visualize slide
- prepare for questions
- AI image editing
- AI image generation
- enterprise assets
- slide explanations
- speaker notes
- Work IQ context
- custom skills
- model selection
- web search with citations
- notebooks/connectors
- multi-turn refinement

The uploaded reference explicitly includes slide generation, edit mode, built-in review/visualization skills, speaker notes, model selection and web search/citations. fileciteturn0file0L8-L25 fileciteturn0file0L30-L45 fileciteturn0file0L50-L66

### Current project parity

Current add-in:

- Presentation text extraction: YES
- Chat/Q&A: YES
- Insert response into slide: YES
- Full deck generation: NO
- Slide rewrite: NO
- Layout optimization: NO
- Standardization: NO
- Speaker-note generation: NO dedicated feature
- Presentation review: NO dedicated skill
- Visualize slide: NO
- Image generation/editing: NO
- Brand integration: NO
- Web research/citations: NO
- Model selection: Mistral-only in current source; multi-provider model selection is a required v0.0.0 change
- Agent mode: NO

---

# 23. Target Product Direction

The product direction is:

**AI Assistant**

into:

**Provider-neutral Office AI Agent Platform**

The Office host should become a presentation/editing surface for a common AI orchestration layer.

## Target layers

```text
+------------------------------------------------------+
|                    Office Hosts                      |
| Word | Excel | PowerPoint |
+----------------------------+-------------------------+
                             |
+----------------------------v-------------------------+
|             Office Capability Layer                  |
| Context | Selection | Read | Write | Undo | Preview |
+----------------------------+-------------------------+
                             |
+----------------------------v-------------------------+
|                 AI Orchestrator                     |
| Prompting | Context | Planning | Tool Calls | Policy|
+----------------------------+-------------------------+
                             |
+----------------------------v-------------------------+
|                 Provider Abstraction                  |
| IAIProvider | Models | Auth | Capabilities | Limits |
+----------------------------+-------------------------+
                             |
       +----------+----------+----------+----------+
       |          |                     |          |
     Mistral   Gemini        Groq        Custom API
```

---

# 24. Provider Architecture — Target Contract

The core project should eventually define provider-neutral contracts.

## Suggested conceptual interfaces

```csharp
interface IAIProvider
{
    string Id { get; }
    ProviderCapabilities Capabilities { get; }

    Task<AIResponse> ChatAsync(
        AIRequest request,
        CancellationToken cancellationToken);
}
```

Supporting concepts:

```text
AIRequest
AIResponse
AIMessage
AIModel
ProviderCapabilities
AuthenticationMode
UsageInfo
ToolDefinition
ToolCall
StreamingEvent
```

### Provider capability flags

At minimum:

- Chat
- Streaming
- Vision
- Structured output
- Tool calling
- Web search
- Embeddings
- Image generation
- Image editing
- Reasoning
- Long context
- OAuth
- API key
- Local model

A provider must advertise what it actually supports.

---

# 25. Authentication Architecture — Target

Authentication must not be embedded in the Office UI or individual provider classes.

Target:

```text
AuthenticationManager
    |
    +-- ApiKeyCredential
    +-- OAuthCredential
    +-- LocalCredential
    +-- BrowserSession (experimental only)
```

Security rule:

**Official API/OAuth flows are preferred. Browser automation must not be treated as equivalent to an official API integration.**

Credentials should remain local and encrypted where appropriate.

---

# 26. Web / Browser Automation Boundary

If browser-based access is later explored, keep it isolated:

```text
WebAutomationProvider
        |
        +-- BrowserEngine
        +-- SessionManager
        +-- PageAdapter
        +-- ChatTransport
```

This must not contaminate the core provider interface.

The architecture must distinguish:

1. Official API
2. Official OAuth
3. Local model
4. Browser/web-chat automation

These are different integration classes with different reliability, security and compatibility properties.

---

# 27. Feature Model

Every AI capability should be represented as a feature/tool rather than as ad-hoc UI code.

Examples:

### Word

```text
summarize_document
rewrite_selection
change_tone
draft_document
insert_table
explain_selection
```

### Excel

```text
analyze_range
generate_formula
modify_range
create_chart
create_pivot
filter_data
summarize_workbook
build_dashboard
```

### PowerPoint

```text
summarize_deck
rewrite_slide
generate_slides
standardize_format
generate_speaker_notes
review_presentation
visualize_slide
prepare_questions
```

### Common

```text
chat
explain
summarize
draft
rewrite
plan
preview
apply
undo
```

---

# 28. Edit/Agent Execution Model

Future agentic editing should use:

```text
User request
     |
     v
Intent classification
     |
     v
Context collection
     |
     v
Plan generation
     |
     v
Validation / policy
     |
     v
Preview
     |
     v
User confirmation
     |
     v
Office operation execution
     |
     v
Result + audit information
```

Do not let raw model output directly execute unrestricted Office operations.

---

# 29. Security Baseline

Required principles:

- Never log API keys.
- Never include API keys in error messages.
- Keep credentials out of source control.
- Use DPAPI or equivalent secure credential storage.
- Validate provider endpoints.
- Use HTTPS for remote providers.
- Avoid transmitting document context unless the user explicitly enables the feature.
- Clearly identify what context is being sent.
- Do not silently send unrelated Office content.
- Minimize context.
- Add cancellation for long-running requests.
- Avoid storing sensitive document content in persistent logs.
- Keep provider credentials isolated from chat history.

---

# 30. Privacy Model

Current implementation:

- Document context is only collected when the checkbox is enabled.
- Context is sent to the configured Mistral API endpoint.
- Chat history is held in memory for the current task-pane instance.
- API key is persisted encrypted in HKCU.

Future privacy UX should expose:

- Provider
- Model
- Endpoint
- Context being shared
- Whether web search is enabled
- Whether external connectors are enabled
- Retention behavior
- Telemetry status

---

# 31. Observability

Current logging:

`%TEMP%\MistralAddinLog.txt`

Logging is used mainly by `Connect`.

Future logging should be structured:

```text
timestamp
host
operation
provider
model
request_id
duration
status
error_code
```

Never log:

- API keys
- full document context
- full user prompts by default
- full model responses by default

---

# 32. Release Gates

A build should not be considered release-ready until:

## Build

- [ ] Clean build
- [ ] Strong-name verification
- [ ] No compiler warnings
- [ ] Correct assembly version

## COM

- [ ] Main COM registration succeeds
- [ ] Task pane ActiveX registration succeeds
- [ ] 32-bit Office verified
- [ ] 64-bit Office verified

## Word

- [ ] Ribbon loads
- [ ] Task pane opens
- [ ] Chat works
- [ ] Context works
- [ ] Insert works
- [ ] Shutdown/restart stable

## Excel

- [ ] Ribbon loads
- [ ] Task pane opens
- [ ] UsedRange extraction works
- [ ] Large/sparse sheet behavior tested
- [ ] Active-cell insertion works

## PowerPoint

- [ ] Ribbon loads
- [ ] Task pane opens
- [ ] Slide text extraction works
- [ ] Text-box insertion works
- [ ] Image-only slides do not crash

## Network

- [ ] Valid key
- [ ] Invalid key
- [ ] 401
- [ ] 404
- [ ] 429
- [ ] timeout
- [ ] offline machine
- [ ] TLS failure

## Security

- [ ] API key not present in plaintext registry
- [ ] API key not present in logs
- [ ] API key not present in source
- [ ] Context transmission behavior confirmed

---

# 33. Current Project Status

## Implemented

- [x] Shared COM add-in
- [x] Word support
- [x] Excel support
- [x] PowerPoint support
- [x] Ribbon
- [x] Custom Task Pane
- [x] ActiveX/CTP registration code
- [x] IObjectSafety
- [x] Mistral chat API (existing source baseline)
- [ ] Provider abstraction
- [ ] Gemini provider
- [ ] Groq provider
- [ ] Custom API provider
- [ ] PDF/Word/Excel/PowerPoint/image upload
- [x] Model test
- [x] Settings
- [x] DPAPI credential storage
- [x] Document/workbook/presentation context
- [x] Chat history
- [x] Insert response
- [x] Clear chat
- [x] Background API execution
- [x] UI synchronization
- [x] Basic API error handling
- [x] Strong-name build support

## Required v0.0.0 changes not yet implemented in the analyzed source

- [ ] Rename product/UI identity to **AI Assistant**
- [ ] Set product version to **0.0.0**
- [ ] Add developer credit
- [ ] Provider abstraction
- [ ] Gemini provider
- [ ] Groq provider
- [ ] Custom API provider
- [ ] Provider selector in Settings
- [ ] Provider-specific credential/model/base-URL storage
- [ ] PDF/Word/Excel/PowerPoint/image upload
- [ ] Upload processing and provider routing
- [ ] Streaming
- [ ] Tool calling
- [ ] Agent framework
- [ ] Plan mode
- [ ] Preview/apply edit model
- [ ] Undo/transaction layer
- [ ] Web search
- [ ] Citations
- [ ] RAG
- [ ] Embeddings
- [ ] Connector framework
- [ ] MCP integration
- [ ] Image generation
- [ ] Image editing
- [ ] Audio
- [ ] Persistent AI memory
- [ ] Work-context aggregation
- [ ] Custom skills framework

---

# 34. Recommended Implementation Sequence

## Phase 1 — Stabilize current COM add-in

1. Live-verify CTP ActiveX registration.
2. Live-verify 32/64-bit Office compatibility.
3. Replace dynamic CTP factory/pane calls with declared COM interfaces.
4. Harden shutdown/restart behavior.
5. Fix 32-bit Windows build/registration edge cases if that platform remains in scope.
6. Add automated smoke tests for all three hosts.

## Phase 2 — Provider abstraction

1. Introduce `IAIProvider`.
2. Move Mistral into `MistralProvider`.
3. Add provider/model registry.
4. Add provider capabilities.
5. Add provider-specific authentication.
6. Preserve current Mistral behavior through the abstraction.

## Phase 3 — AI orchestration

1. Context manager.
2. Token budget.
3. Prompt assembler.
4. Tool registry.
5. Plan/preview/apply pipeline.
6. Cancellation.
7. Structured responses.

## Phase 4 — Office agent capabilities

1. Word rewrite/summarize/draft.
2. Excel formula/range/chart operations.
3. PowerPoint slide operations.
4. Preview and confirmation.
5. Undo/reversible operations.

## Phase 5 — Optional capabilities outside v0.0.0

The following are deliberately outside the v0.0.0 provider/product scope and must not be implemented as additional AI-provider choices without an explicit SSOT revision:

1. RAG.
2. Connectors/MCP.
3. Advanced web-search integration.
4. Browser automation.

---

# 35. Non-Goals for the Current Baseline

Do not accidentally expand the current baseline into:

- Microsoft 365 Copilot itself
- Office.js architecture
- VSTO-only architecture
- cloud middleware server
- unrestricted browser automation
- unrestricted agentic Office editing
- automatic external data access

The analyzed baseline is a **local desktop COM add-in with direct Mistral API communication**. The required v0.0.0 product architecture adds provider selection for Mistral AI, Google Gemini, Groq and Custom API.

---

# 36. Change-Control Rules

When changing the project:

1. Update source code.
2. Update this SSOT if architecture, behavior, supported hosts, providers, security, registration or major features change.
3. Mark features as `Implemented`, `Verified`, `Planned`, or `Not Implemented`.
4. Do not mark a feature `Verified` from source inspection alone.
5. Preserve backward compatibility unless a deliberate breaking change is documented.
6. Update build/register/unregister scripts together with COM changes.
7. Update verification scripts whenever a host or registration contract changes.
8. Never add a provider-specific feature to the core abstraction without documenting its capability requirements.

---

# 37. Canonical Terminology

Use these names consistently:

- **Office Host** — Word, Excel and PowerPoint
- **Shared COM Add-in** — the main extension model
- **Connect** — COM add-in entry point
- **TaskPaneControl** — the hosted chat UI
- **MistralClient** — current provider-specific HTTP client
- **Provider** — AI backend implementation
- **Context** — Office content supplied to the AI
- **Chat History** — in-memory conversation messages
- **Insert** — current simple write-back operation
- **Agent** — future plan/tool execution layer
- **SSOT** — this document

---

# 38. Final Architecture Decision

### Current baseline

**Keep the shared COM add-in architecture for the Office 2010 → modern desktop compatibility requirement.**

### Immediate architectural priority

**Refactor the AI layer to provider-neutral interfaces without disturbing the Office COM layer.**

### Long-term architecture

```text
COM Office Adapter
        |
        v
Office Capability API
        |
        v
AI Orchestrator
        |
        v
Provider Abstraction
        |
        +--> Mistral
        +--> Gemini
        +--> Groq
        +--> Custom API
```

The Office integration and AI provider integration must remain independently replaceable.

---

# 39. Reference Documents

### Project-provided

- `Design Plan.md`
- `Design Review.md`
- `mistral-office-addin/README.md`
- `src/*.cs`
- `build.cmd`
- `register.cmd`
- `unregister.cmd`
- `tools/verify.ps1`
- `tools/DumpInterfaces.ps1`

### Capability references

- `Microsoft_365_Copilot_Word_Features.md`
- `Office365_Copilot_Excel_Features.md`
- `Office365_Copilot_PowerPoint_Features.md`

The Copilot documents describe the reference feature set as of August 2026; they are not evidence that those features exist in this codebase.

---

## SSOT Status

**Current source baseline:** Mistral-only shared COM add-in for Word, Excel and PowerPoint.

**Required product baseline:** AI Assistant v0.0.0 with Mistral AI, Google Gemini, Groq and Custom API provider options; Word, Excel and PowerPoint only; document/image upload support; developer credit as specified in Section 1.

**Primary technical objective:** establish the AI Assistant v0.0.0 identity and scope, add the four required AI providers and file-upload capability, while preserving the existing Word/Excel/PowerPoint COM/CTP foundation.

**Last analyzed:** 2026-08-17
