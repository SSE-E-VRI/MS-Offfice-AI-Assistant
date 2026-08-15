# Mistral AI Office Assistant — Project Requirements Specification

> **Version:** 1.0  
> **Date:** 2026-08-15  
> **Target:** Office 2010 → Office 365 (Word, Excel, PowerPoint, Outlook)  
> **IDE:** Google Antigravity IDE (VS Code fork)  
> **API:** Mistral AI (OpenAI-compatible, BYOK — Bring Your Own Key)

---

## Table of Contents

1. [Executive Summary](#1-executive-summary)
2. [Development Environment](#2-development-environment)
3. [Architecture Overview](#3-architecture-overview)
4. [Technology Stack](#4-technology-stack)
5. [Application Coverage](#5-application-coverage)
6. [UI/UX Requirements](#6-uiux-requirements)
7. [Mistral API Integration](#7-mistral-api-integration)
8. [Feature Matrix (Copilot Parity)](#8-feature-matrix-copilot-parity)
9. [Security & Privacy](#9-security--privacy)
10. [Deployment & Installation](#10-deployment--installation)
11. [Performance & Reliability](#11-performance--reliability)
12. [Version-Specific Limitations](#12-version-specific-limitations)
13. [Project Structure](#13-project-structure)
14. [Testing Matrix](#14-testing-matrix)

---

## 1. Executive Summary

Build a **cross-version Microsoft Office COM Add-in** that brings Copilot-like AI capabilities to **Office 2010 through Office 365** using the user's own **Mistral API key** (BYOK). The add-in must be **free from vendor lock-in**, run entirely on the user's machine, communicate directly with Mistral's API, and support **Word, Excel, PowerPoint, and Outlook**.

### Key Principles
- **No intermediary servers** — direct HTTPS from client to `api.mistral.ai`
- **No subscription gates** — user brings their own free/paid Mistral API key
- **Open architecture** — source code is the user's property
- **Graceful degradation** — detect Office version and disable unsupported features

---

## 2. Development Environment

| Requirement | Specification |
|-------------|---------------|
| **IDE** | Google Antigravity IDE (VS Code fork with agentic AI) |
| **Build System** | MSBuild via command line OR `dotnet` CLI for .NET Framework |
| **Target Framework** | .NET Framework 4.8 (last supported major version for Office add-ins) |
| **Add-in Type** | **COM Add-in** (not VSTO — VSTO requires Visual Studio project templates) |
| **Interop Strategy** | NetOfficeFw (version-agnostic Office interop) |
| **Language** | C# |
| **Package Manager** | NuGet CLI (`nuget install` or `dotnet add package`) |

### Why COM Add-in over VSTO?
- VSTO requires Visual Studio with Office/SharePoint workload for project scaffolding, debugging, and designers.
- COM Add-ins can be built entirely with command-line tools, C# code, and XML — perfect for Antigravity IDE.
- COM Add-ins work on **Office 2000+**, giving maximum backward compatibility.

---

## 3. Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    OFFICE APPLICATION                        │
│  ┌─────────────┐  ┌──────────────────┐  ┌─────────────────┐ │
│  │ Ribbon Tab  │  │ Custom Task Pane │  │ Context Menu    │ │
│  │ (XML + C#)  │  │ (WPF Sidebar)    │  │ (CommandBar/    │ │
│  │             │  │                  │  │  Ribbon)        │ │
│  └──────┬──────┘  └────────┬─────────┘  └─────────────────┘ │
│         │                    │                                │
│  ┌──────┴────────────────────┴─────────────────────────────┐  │
│  │              COM Add-in Shim (.NET 4.8)               │  │
│  │         Implements: IDTExtensibility2, IRibbonExtensibility││
│  └───────────────────────────┬─────────────────────────────┘  │
└──────────────────────────────┼───────────────────────────────┘
                               │
┌──────────────────────────────┼───────────────────────────────┐
│         .NET ASSEMBLY        │                               │
│  ┌───────────────────────────┴─────────────────────────────┐  │
│  │              WPF Chat UI (Sidebar)                     │  │
│  │  • Markdown rendering (Markdig)                        │  │
│  │  • Streaming token display                             │  │
│  │  • Model selector, temperature, max tokens           │  │
│  │  • Conversation history per document                   │  │
│  └───────────────────────────┬─────────────────────────────┘  │
│  ┌───────────────────────────┴─────────────────────────────┐  │
│  │              Mistral API Client                        │  │
│  │  • HttpClient with async/await                         │  │
│  │  • OpenAI-compatible endpoints                         │  │
│  │  • SSE (Server-Sent Events) streaming parser           │  │
│  │  • Retry logic with exponential backoff                │  │
│  └───────────────────────────┬─────────────────────────────┘  │
│  ┌───────────────────────────┴─────────────────────────────┐  │
│  │              Document Controller                       │  │
│  │  • Word: Insert/replace text, Track Changes            │  │
│  │  • Excel: Read cells, write formulas, analyze ranges   │  │
│  │  • PowerPoint: Modify slides, speaker notes            │  │
│  │  • Outlook: Read emails, draft replies                 │  │
│  └─────────────────────────────────────────────────────────┘  │
└───────────────────────────────────────────────────────────────┘
```

---

## 4. Technology Stack

### 4.1 Core NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `NetOfficeFw.Word` | Latest | Word interop (version-agnostic) |
| `NetOfficeFw.Excel` | Latest | Excel interop (version-agnostic) |
| `NetOfficeFw.PowerPoint` | Latest | PowerPoint interop |
| `NetOfficeFw.Outlook` | Latest | Outlook interop |
| `NetOfficeFw.Office` | Latest | Shared Office types |
| `Newtonsoft.Json` | 13.x | JSON serialization for API |
| `Markdig` | Latest | Markdown → WPF rendering |
| `Markdig.Wpf` | Latest | WPF integration for Markdig |
| `Hardcodet.NotifyIcon.Wpf` | Latest | System tray icon |

### 4.2 System Dependencies

| Dependency | Minimum Version | Notes |
|------------|-----------------|-------|
| .NET Framework | 4.5 | 4.8 strongly recommended |
| VSTO 2010 Runtime | 10.0.60828 | Required for Office 2010; bundled with 2013+ |
| Windows | 7 SP1 | Windows 10/11 recommended |
| TLS | 1.2 | Enforced for all API calls |

---

## 5. Application Coverage

| Application | Min Version | Max Version | Integration Points |
|-------------|-------------|-------------|-------------------|
| **Microsoft Word** | 2010 (14.0) | 365 (16.0+) | Ribbon, Task Pane, Context Menu, Track Changes |
| **Microsoft Excel** | 2010 (14.0) | 365 (16.0+) | Ribbon, Task Pane, Context Menu, UDF support |
| **Microsoft PowerPoint** | 2010 (14.0) | 365 (16.0+) | Ribbon, Task Pane, Context Menu, Slide notes |
| **Microsoft Outlook** | 2010 (14.0) | 365 (16.0+) | Ribbon (Explorer/Inspector), Task Pane, Compose/Read |

### Version Detection
```csharp
// Runtime version detection
string version = application.Version; // e.g., "14.0", "15.0", "16.0"
double majorVersion = double.Parse(version.Split('.')[0]);
// 14 = 2010, 15 = 2013, 16 = 2016/2019/2021/365
```

---

## 6. UI/UX Requirements

### 6.1 Custom Task Pane — AI Sidebar

**Technology:** WPF UserControl inside Windows Forms `ElementHost`

**Layout:**
```
┌─────────────────────────────────┐
│  🤖 Mistral AI Assistant    [⚙] │  ← Header with settings button
├─────────────────────────────────┤
│                                 │
│  [System] Welcome! How can I   │  ← Chat history (scrollable)
│           help you today?      │
│                                 │
│  [User] Summarize this doc     │
│                                 │
│  [System] Here's a summary...   │  ← Markdown rendered
│           • Point 1            │
│           • Point 2            │
│                                 │
├─────────────────────────────────┤
│  Model: [mistral-large ▼]      │  ← Model selector
│  Temp: [0.7] Tokens: [4096]    │  ← Sliders/inputs
├─────────────────────────────────┤
│  Type your message...      [➤] │  ← Input + send
└─────────────────────────────────┘
```

**Requirements:**
- [ ] **Streaming display**: Tokens appear in real-time as they arrive from API
- [ ] **Markdown rendering**: Bold, italic, code blocks, bullet lists, tables
- [ ] **Syntax highlighting**: For code snippets (optional, nice-to-have)
- [ ] **Copy button**: Per-message copy to clipboard
- [ ] **Regenerate button**: Retry last prompt
- [ ] **New Chat**: Clear conversation history (per-document isolation)
- [ ] **Resizeable**: User can drag sidebar width
- [ ] **Keyboard shortcuts**: Ctrl+Enter to send, Escape to close sidebar

### 6.2 Ribbon Tab — "Mistral AI"

**Office 2010+ Ribbon XML Structure:**
```xml
<tab id="tabMistralAI" label="Mistral AI">
  <group id="grpChat" label="Chat">
    <button id="btnToggleSidebar" label="Open AI Chat" imageMso="HappyFace"/>
    <button id="btnNewChat" label="New Chat" imageMso="NewAppointment"/>
  </group>
  <group id="grpDraft" label="Draft">
    <button id="btnGenerate" label="Generate" imageMso="CreateReport"/>
    <button id="btnContinue" label="Continue Writing" imageMso="GoToNextEdit"/>
  </group>
  <group id="grpEdit" label="Edit">
    <button id="btnSummarize" label="Summarize" imageMso="AutoSummary"/>
    <button id="btnRewrite" label="Rewrite" imageMso="ReviewTrackChanges"/>
    <button id="btnExpand" label="Expand" imageMso="InkExpand"/>
    <button id="btnShorten" label="Shorten" imageMso="ShrinkOnePage"/>
    <menu id="menuTranslate" label="Translate" imageMso="Translate">
      <button id="btnTransEN" label="English"/>
      <button id="btnTransES" label="Spanish"/>
      <button id="btnTransFR" label="French"/>
      <button id="btnTransDE" label="German"/>
      <button id="btnTransZH" label="Chinese"/>
    </menu>
  </group>
  <group id="grpSettings" label="Settings">
    <button id="btnSettings" label="Configure" imageMso="AdpDiagramTableRelationships"/>
  </group>
</tab>
```

### 6.3 Context Menu Integration

**Office 2010**: Use `CommandBars` API (`MsoControlType.msoControlButton`)
**Office 2013+**: Use Ribbon XML `contextMenu` element

**Menu Items:**
- "Ask Mistral about this"
- "Summarize selection"
- "Rewrite selection"
- "Translate selection"
- "Explain this"

---

## 7. Mistral API Integration

### 7.1 Configuration Schema

```json
{
  "baseUrl": "https://api.mistral.ai/v1",
  "apiKey": "<encrypted>",
  "defaultModel": "mistral-large-latest",
  "fallbackModel": "mistral-small-latest",
  "maxTokens": 4096,
  "temperature": 0.7,
  "topP": 1.0,
  "safeMode": false,
  "streamResponses": true,
  "systemPrompts": {
    "word": "You are a professional document assistant...",
    "excel": "You are a data analyst. For formulas, return ONLY the formula...",
    "powerpoint": "You are a presentation designer...",
    "outlook": "You are an email assistant. Maintain professional tone..."
  }
}
```

### 7.2 API Endpoints

| Endpoint | Method | Purpose |
|----------|--------|---------|
| `POST /chat/completions` | Chat | All conversational AI features |
| `POST /embeddings` | Embeddings | Future: semantic document search |

### 7.3 Request/Response Format

**Request:**
```csharp
var request = new
{
    model = "mistral-large-latest",
    messages = new[]
    {
        new { role = "system", content = systemPrompt },
        new { role = "user", content = userPrompt }
    },
    stream = true,
    temperature = 0.7,
    max_tokens = 4096
};
```

**Streaming Response (SSE):**
```
data: {"id":"...","object":"chat.completion.chunk","choices":[{"delta":{"content":"Hello"}}]}

data: {"id":"...","object":"chat.completion.chunk","choices":[{"delta":{"content":" world"}}]}

data: [DONE]
```

### 7.4 Context Management

- **Per-document isolation**: Each document maintains its own conversation thread
- **Token budget tracking**: Approximate token count using character-based heuristic (`chars / 4`)
- **Auto-trimming**: When approaching context limit, remove oldest messages (keep system prompt)
- **Max history**: Default 10 message pairs per conversation

---

## 8. Feature Matrix (Copilot Parity)

### 8.1 Universal Features (All Applications)

| Feature | Description | Priority | Office 2010 Support |
|---------|-------------|----------|-------------------|
| **Chat** | General Q&A with document context | P0 | ✅ Full |
| **Draft** | Generate content from prompt | P0 | ✅ Full |
| **Summarize** | Summarize selection or document | P0 | ✅ Full |
| **Rewrite** | Paraphrase, improve grammar | P0 | ✅ Full |
| **Tone Change** | Professional, casual, persuasive, concise | P1 | ✅ Full |
| **Translate** | Translate to 20+ languages | P1 | ✅ Full |
| **Explain** | Explain complex concepts | P1 | ✅ Full |
| **Custom Prompts** | User-defined quick prompts | P2 | ✅ Full |

### 8.2 Microsoft Word

| Feature | Description | Priority | Notes |
|---------|-------------|----------|-------|
| **Insert at Cursor** | Insert AI output at cursor position | P0 | |
| **Replace Selection** | Replace selected text | P0 | |
| **Track Changes** | AI edits appear as tracked changes | P1 | Word 2010+ native support |
| **Document Context** | Feed entire document as context | P1 | Token-limited; summarize if too long |
| **Style Preservation** | Match existing font/formatting | P2 | |
| **Table Generation** | Generate tables from description | P2 | |

### 8.3 Microsoft Excel

| Feature | Description | Priority | Notes |
|---------|-------------|----------|-------|
| **Formula Generation** | Natural language → Excel formula | P0 | Return only formula string |
| **Formula Explanation** | Explain selected cell's formula | P0 | |
| **Data Analysis** | Insights from selected range | P1 | |
| **Sample Data** | Generate realistic sample data | P2 | |
| **Chart Recommendation** | Suggest chart type for data | P2 | Text-only suggestion |

### 8.4 Microsoft PowerPoint

| Feature | Description | Priority | Notes |
|---------|-------------|----------|-------|
| **Slide Content** | Generate bullet points from topic | P0 | |
| **Speaker Notes** | Generate notes for current slide | P1 | |
| **Slide Summary** | Summarize slide content | P1 | |
| **Outline Expansion** | Expand outline into full slides | P2 | |

### 8.5 Microsoft Outlook

| Feature | Description | Priority | Notes |
|---------|-------------|----------|-------|
| **Email Draft** | Draft email from prompt | P0 | Compose inspector |
| **Reply Generation** | Generate reply to selected email | P0 | Reading pane + compose |
| **Email Summary** | Summarize long email/thread | P1 | |
| **Tone Adjustment** | Adjust draft tone | P1 | |

---

## 9. Security & Privacy

### 9.1 API Key Protection

| Layer | Implementation |
|-------|----------------|
| **Encryption** | Windows DPAPI (`ProtectedData.Protect`) |
| **Storage** | `%LOCALAPPDATA%\MistralOfficeAddin\config.dat` |
| **Access** | Only decrypt at runtime, never log |
| **UI** | Password-masked input field |

### 9.2 Data Privacy

- **No proxy**: Direct HTTPS from machine to `api.mistral.ai`
- **No telemetry**: Zero analytics or usage tracking
- **No cloud storage**: All data stays local except explicit API calls
- **User consent**: Prompt before sending large documents (>4000 tokens)
- **TLS enforcement**: Reject connections below TLS 1.2

### 9.3 Audit Trail

- Log API calls locally (optional, user-enabled) for debugging
- Log format: timestamp, model, token count, latency
- Stored in `%LOCALAPPDATA%\MistralOfficeAddin\logs\`

---

## 10. Deployment & Installation

### 10.1 Installer Requirements

**Tool:** Inno Setup 6 (free, scriptable, supports both x86/x64)

**Outputs:**
- `MistralOfficeAddin-x86.exe` — for 32-bit Office
- `MistralOfficeAddin-x64.exe` — for 64-bit Office

**Prerequisites (bundled or downloaded):**
- .NET Framework 4.8 Web Installer
- VSTO 2010 Runtime (only for Office 2010 users)

### 10.2 COM Registration

The installer must register the COM Add-in for each Office application:

```
HKCU\Software\Microsoft\Office\Word\Addins\MistralAI.Addin
    FriendlyName = "Mistral AI Assistant"
    Description = "AI-powered assistant using your own Mistral API key"
    LoadBehavior = 3
    Manifest = "C:\Program Files\MistralOfficeAddin\MistralAI.dll"

HKCU\Software\Microsoft\Office\Excel\Addins\MistralAI.Addin
    [same structure]

HKCU\Software\Microsoft\Office\PowerPoint\Addins\MistralAI.Addin
    [same structure]

HKCU\Software\Microsoft\Office\Outlook\Addins\MistralAI.Addin
    [same structure]
```

### 10.3 LoadBehavior Values

| Value | Meaning |
|-------|---------|
| `0` | Unloaded |
| `1` | Loaded |
| `2` | Load on startup, unloaded on exit |
| `3` | **Load on startup** (recommended) |
| `8` | Load on demand |
| `9` | Load on startup + demand |

### 10.4 Uninstallation

- Remove all registry entries
- Delete `%LOCALAPPDATA%\MistralOfficeAddin\`
- Unregister COM component via `regasm /unregister`

---

## 11. Performance & Reliability

### 11.1 Performance Targets

| Metric | Target | Measurement |
|--------|--------|-------------|
| Add-in Load Time | < 2s | From Office startup to ribbon visible |
| Sidebar Open | < 500ms | Click to visible task pane |
| First Token | < 3s | Send click to first streamed token |
| Streaming Latency | < 100ms/token | Between consecutive tokens |
| Memory Footprint | < 150MB | Working set during active use |
| Idle Memory | < 50MB | When sidebar is closed |

### 11.2 Reliability Requirements

- [ ] **Async-only API calls**: Never block UI thread
- [ ] **Cancellation tokens**: Allow user to abort in-flight requests
- [ ] **Timeout handling**: 30s default, configurable
- [ ] **Retry logic**: Exponential backoff (1s, 2s, 4s, 8s) for 5xx errors
- [ ] **Offline detection**: Detect network loss, queue actions, notify user
- [ ] **Graceful degradation**: If API fails, show error message (not crash)
- [ ] **Exception isolation**: One app's error doesn't affect others

### 11.3 Error Handling

| Scenario | Behavior |
|----------|----------|
| Invalid API key | Show settings dialog with red error text |
| Rate limit (429) | Display: "Rate limited. Retrying in X seconds..." |
| Network timeout | "Connection timeout. Check your internet." |
| Token limit exceeded | "Message too long. Try a shorter prompt or clear history." |
| Office version too old | Disable features requiring newer APIs |

---

## 12. Version-Specific Limitations

Be transparent with users. These Copilot 365 features are **technically impossible** on older Office:

| Feature | Office 2010 | Office 2013 | Office 2016 | Office 365 |
|---------|-------------|-------------|-------------|--------------|
| Modern Fluent UI | ❌ | ❌ | ⚠️ Partial | ✅ |
| Real-time co-authoring AI | ❌ | ❌ | ❌ | ✅ |
| Excel Data Types | ❌ | ❌ | ❌ | ✅ |
| PowerPoint Designer | ❌ | ❌ | ⚠️ Text only | ✅ |
| Semantic Index / Graph | ❌ | ❌ | ❌ | ✅ |
| Loop Components | ❌ | ❌ | ❌ | ✅ |
| Web Add-in Support | ❌ | ❌ | ✅ | ✅ |
| Track Changes (AI edits) | ✅ | ✅ | ✅ | ✅ |
| Custom Task Pane | ✅ | ✅ | ✅ | ✅ |
| Ribbon XML | ✅ | ✅ | ✅ | ✅ |

---

## 13. Project Structure

```
MistralOfficeAddin/
├── src/
│   ├── MistralOfficeAddin.csproj          # Project file
│   ├── Properties/
│   │   ├── AssemblyInfo.cs
│   │   └── Resources.resx
│   ├── Addin/
│   │   ├── Connect.cs                     # IDTExtensibility2 implementation
│   │   ├── Ribbon.xml                     # Ribbon definition
│   │   ├── RibbonCallback.cs              # Ribbon button handlers
│   │   └── CustomTaskPaneManager.cs       # Task pane lifecycle
│   ├── UI/
│   │   ├── ChatSidebar.xaml               # WPF chat interface
│   │   ├── ChatSidebar.xaml.cs
│   │   ├── SettingsWindow.xaml            # API key/settings dialog
│   │   ├── SettingsWindow.xaml.cs
│   │   └── Converters/
│   │       └── MarkdownConverter.cs
│   ├── API/
│   │   ├── MistralClient.cs               # HttpClient wrapper
│   │   ├── StreamingParser.cs             # SSE parser
│   │   ├── TokenCounter.cs                # Approximate token counting
│   │   └── Models/
│   │       ├── ChatRequest.cs
│   │       ├── ChatResponse.cs
│   │       └── StreamingChunk.cs
│   ├── Core/
│   │   ├── ConfigManager.cs               # DPAPI-encrypted config
│   │   ├── ConversationStore.cs           # Per-document history
│   │   ├── VersionDetector.cs             # Office version detection
│   │   └── Logger.cs                      # Local debug logging
│   └── Hosts/
│       ├── WordController.cs              # Word-specific operations
│       ├── ExcelController.cs             # Excel-specific operations
│       ├── PowerPointController.cs        # PPT-specific operations
│       └── OutlookController.cs           # Outlook-specific operations
├── installer/
│   ├── setup-x86.iss                      # Inno Setup 32-bit
│   ├── setup-x64.iss                      # Inno Setup 64-bit
│   └── assets/
│       ├── banner.bmp
│       └── icon.ico
├── tests/
│   └── [Unit tests - optional v1]
├── docs/
│   ├── USER_GUIDE.md
│   └── TROUBLESHOOTING.md
├── README.md
├── LICENSE
└── build.bat                              # One-click build script
```

---

## 14. Testing Matrix

### 14.1 Minimum Test Environment

| OS | Office Version | Architecture | Priority |
|----|----------------|--------------|----------|
| Windows 10 | Office 2010 (14.0) | x86 | P0 |
| Windows 10 | Office 2013 (15.0) | x64 | P1 |
| Windows 10 | Office 2016 (16.0) | x64 | P1 |
| Windows 11 | Office 2021 (16.0) | x64 | P1 |
| Windows 11 | Microsoft 365 Current | x64 | P0 |

### 14.2 Test Scenarios

1. **Installation**: Install on clean machine, verify all Office apps show ribbon tab
2. **API Key**: Enter key, verify encryption, test connection
3. **Chat**: Send message, verify streaming, verify markdown rendering
4. **Word**: Select text → Summarize → Verify output inserted
5. **Excel**: Select range → Analyze → Verify insight displayed
6. **PowerPoint**: Generate slide bullets → Verify added to slide
7. **Outlook**: Draft email → Verify text in compose body
8. **Cross-version**: Install on Office 2010, upgrade to 365, verify persistence
9. **Uninstall**: Remove, verify registry clean, verify Office loads normally

---

## Appendix A: Mistral Free Tier Limits

| Limit | Value |
|-------|-------|
| Rate Limit | ~1 request/second |
| RPM (Requests Per Minute) | ~60 |
| TPM (Tokens Per Minute) | Varies by model |
| Cost | $0 (free tier) |
| Best Models | `mistral-small-latest`, `open-mistral-nemo` |

> Note: Free tier is generous for personal use but has rate limits. For heavy usage, users can add a small pay-as-you-go balance.

---

## Appendix B: Glossary

| Term | Definition |
|------|------------|
| **BYOK** | Bring Your Own Key — user provides their own API key |
| **COM Add-in** | Component Object Model add-in; works across all Office versions |
| **DPAPI** | Data Protection API — Windows built-in encryption |
| **PIA** | Primary Interop Assembly — .NET wrapper for Office COM |
| **SSE** | Server-Sent Events — streaming response format |
| **VSTO** | Visual Studio Tools for Office — requires VS |
| **WPF** | Windows Presentation Foundation — modern UI framework |

---

*End of Requirements Specification*
