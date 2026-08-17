# AI Assistant for Microsoft Office — User Guide

**Version 0.0.0**  
*Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302*

Welcome to **AI Assistant**, a cross-version Microsoft Office add-in that brings multi-provider AI capabilities (Mistral, Groq, Gemini, and Custom OpenAI-compatible endpoints) to **Word, Excel, and PowerPoint** from **Office 2010 to Office 365**.

---

## 🚀 Getting Started

### 1. Installation

#### Automated Installer (Recommended)
1. Run `AIAssistant-Setup-x64.exe` (for 64-bit Office) or `AIAssistant-Setup-x86.exe` (for 32-bit Office).
2. Follow the setup wizard instructions.

#### Developer / Local Installation
1. Run `build.bat` in the repository root to compile both x86 and x64 assemblies.
2. Run `register.cmd` to register the COM classes and add-in entries for the current user.

---

### 2. Configuration & API Key

1. Open Microsoft Word, Excel, or PowerPoint.
2. Switch to the **AI Assistant** tab on the Ribbon.
3. Click **Configure** (or the ⚙️ icon in the chat sidebar).
4. Select your preferred provider (Mistral, Groq, Gemini, or Custom API).
5. Paste your API key, choose or enter your model, and click **Test Connection** to verify.
6. Click **Save Settings**.

> **Security Note:** Your credentials are encrypted using Windows DPAPI (`DataProtectionScope.CurrentUser`) and stored securely in `%LOCALAPPDATA%\MistralOfficeAddin\config.dat`. They are never transmitted to third-party intermediary servers—only directly to the selected provider endpoint.

---

## 🎯 Features by Application

### 📄 Microsoft Word
- **Open AI Chat**: Open the WPF sidebar to chat with Mistral AI about your document.
- **Generate**: Draft new paragraphs, outlines, and proposals.
- **Continue Writing**: Seamlessly continue text starting from your current cursor position.
- **Summarize**: Summarize highlighted sections or the entire active document.
- **Rewrite**: Improve clarity, tone, and formatting of selected text.
- **Expand / Shorten**: Adjust verbosity with one click.
- **Translate**: Translate selection to English, Spanish, French, German, or Chinese.
- **Insert Response**: Click the 📝 Insert button beneath any AI message to write the output directly into the document.

### 📊 Microsoft Excel
- **Data & Formula Analysis**: Select a range of cells and ask Mistral AI to analyze trends or suggest formulas.
- **Formula Writing**: Ask the assistant to generate `=XLOOKUP`, `=SUMIFS`, or complex financial formulas.

### 📽️ Microsoft PowerPoint
- **Generate Bullet Points**: Create formatted bullet points and insert them directly into slide text boxes.
- **Speaker Notes**: Draft and apply speaker notes to the active slide.

---

## ⚙️ Advanced Customization
- **System Instructions**: Customize the system prompt in the Settings dialog to tune the AI persona to your industry.
- **Creativity & Tokens**: Adjust the temperature slider (0.0 for deterministic output, 1.0 for creative brainstorming) and max token ceiling.
