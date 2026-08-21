# AI Assistant for Microsoft Office 🤖

**Version 0.4.0**  
*Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302*

A cross-version Microsoft Office COM Add-in bringing modern AI chat and document intelligence to **Word, Excel, and PowerPoint** across **Office 2010 through Office 365**.

Powered by a provider-neutral orchestration architecture supporting **Mistral, Groq, Gemini, and Custom OpenAI-compatible endpoints** with a Bring-Your-Own-Key (BYOK) model, zero intermediary servers, and encrypted local storage via Windows DPAPI.

---

## ✨ Features

- **Cross-Version & Universal**: Runs smoothly on Office 2010, 2013, 2016, 2019, 2021, and Microsoft 365 (both 32-bit and 64-bit architectures).
- **Multi-Provider AI Support**: Seamlessly switch between Mistral, Groq, Gemini, and Custom OpenAI-compatible endpoints (such as local Ollama or private LLM gateways).
- **Modern WPF Sidebar UI**: Sleek, responsive chat interface with Markdig Markdown rendering, animated token streaming, code blocks, and instant copy/insert buttons.
- **Deep Office Integration**:
  - **Word**: Document drafting, text continuation, summarization, rewriting, track changes integration, and multilingual translation.
  - **Excel**: Selection range analysis and formula generation (`=XLOOKUP`, `=SUMIFS`, etc.).
  - **PowerPoint**: Slide context reading, bullet point generation, and speaker notes creation.
- **Local Attachments & Vision**: Safe Open XML text extraction (.docx, .xlsx, .pptx), PDF extraction via PdfPig, and image attachment routing to vision-enabled models.
- **Real-Time Streaming**: Low-latency SSE parser streaming tokens live as AI models generate them.
- **Security & Privacy**: Direct HTTPS connections to AI provider APIs. API keys are encrypted with Windows DPAPI. No telemetry, no middleman servers.
- **In-app User Manual**: Click **User Manual** on the AI Assistant ribbon (or the **?** in the chat panel) for a full guide to every button, prompt examples, and railway office workflows — it works even before you add an API key.

---

## 🏗️ Architecture

```
Office Application (Word, Excel, PowerPoint)
    │
    ├── Ribbon Tab (Ribbon.xml + RibbonCallback.cs)
    └── Custom Task Pane (ICustomTaskPaneConsumer)
            │
            └── TaskPaneControl (ActiveX / WinForms + IObjectSafety)
                    │
                    └── ElementHost
                            │
                            └── ChatSidebar (WPF UserControl)
                                    ├── Markdig Markdown Renderer
                                    ├── ChatOrchestrator (IAIProvider)
                                    │     ├── MistralProvider
                                    │     ├── GroqProvider
                                    │     ├── GeminiProvider
                                    │     └── CustomApiProvider
                                    ├── AttachmentExtractor (OpenXML / PdfPig / Vision)
                                    ├── ConfigManager (DPAPI Encryption)
                                    └── Host Controllers (Word/Excel/PowerPoint)
```

---

## 🛠️ Building from Source

### Prerequisites
- Windows 7 SP1, 8.1, 10, or 11
- .NET Framework 4.8
- Microsoft Office 2010, 2013, 2016, 2019, 2021, or Microsoft 365

### Build and install
Run `install.cmd` from the root directory:
```cmd
install.cmd
```
This script will:
1. Locate MSBuild and download `nuget.exe` if not present.
2. Restore all required NuGet packages.
3. Compile both `x86` (32-bit) and `x64` (64-bit) Release assemblies into `bin\x86\Release\` and `bin\x64\Release\`.
4. Register the COM components for the current user (HKCU, `LoadBehavior=3`) — no administrator elevation required.

To uninstall:
```cmd
uninstall.cmd
```

### Build only
`install.cmd` always registers as well. For a build-only loop, invoke MSBuild directly — once per platform:
```cmd
msbuild src\MSOfficeAIAssistant.csproj /p:Configuration=Release /p:Platform=x86
msbuild src\MSOfficeAIAssistant.csproj /p:Configuration=Release /p:Platform=x64
```

### Tests
Build `tests\MSOfficeAIAssistant.Tests.csproj` and run the resulting console runner:
```cmd
bin\x86\Release\MSOfficeAIAssistant.Tests.exe
```

---

## 📦 Distribution

Compile `installer/setup-x86.iss` or `installer/setup-x64.iss` with Inno Setup 6 to produce standalone installer `.exe` files.

---

## 📘 Documentation

- **In-app User Manual** — the **User Manual** button on the AI Assistant ribbon, or **?** in the chat panel.
- **[AI_Assistant_SSOT.md](AI_Assistant_SSOT.md)** — the single source of truth: architecture, COM contracts, roadmap, troubleshooting (Appendix B), and change-control rules.

---

## 📄 License
MIT License. See [LICENSE](LICENSE) for details.
