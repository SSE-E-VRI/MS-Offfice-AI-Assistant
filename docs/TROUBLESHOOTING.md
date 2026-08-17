# AI Assistant for Microsoft Office — Troubleshooting Guide

**Version 0.0.0**  
*Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302*

This guide covers common issues and resolutions for the AI Assistant Office Add-in.

---

## 🔍 Diagnostic Logs

The add-in writes local execution logs to:
```
%LOCALAPPDATA%\MistralOfficeAddin\addin.log
```
Or fallback location:
```
%TEMP%\MistralAddinLog.txt
```
Check these files to view detailed error stack traces and host interaction events.

---

## ⚠️ Common Issues & Solutions

### 1. Ribbon Tab Does Not Appear

**Cause:** Add-in `LoadBehavior` is disabled or bitness mismatch.

**Solution:**
1. Check COM bitness: Ensure you registered the 32-bit DLL for 32-bit Office or the 64-bit DLL for 64-bit Office.
2. In Office, go to **File → Options → Add-ins**.
3. Select **COM Add-ins** in the Manage dropdown at the bottom and click **Go...**.
4. Ensure **Mistral AI Assistant** is checked. If it is unchecked or lists an error, check the load status at the bottom.
5. In the Registry, ensure `HKCU\Software\Microsoft\Office\<App>\Addins\MistralAI.Addin\LoadBehavior` is set to `3` (DWORD).

---

### 2. "Task pane factory is not available"

**Cause:** Office host has not finished exposing `ICustomTaskPaneConsumer` or the application was started in embedded mode.

**Solution:**
1. Restart the Office application normally (not inside another application preview).
2. Open a new document before clicking Open AI Chat.

---

### 3. API Error 401: Unauthorized / Invalid API Key

**Cause:** The Mistral API key is incorrect, expired, or missing.

**Solution:**
1. Go to [console.mistral.ai](https://console.mistral.ai) and generate a new API key.
2. Open Settings (⚙️ in sidebar or Ribbon Configure).
3. Paste the new key and click **Test Connection**.

---

### 4. API Error 429: Rate Limit Exceeded

**Cause:** Your Mistral tier has hit concurrency or requests-per-minute limits.

**Solution:**
- The add-in includes automatic exponential backoff retry.
- For high-volume usage, consider switching to `mistral-small-latest` or upgrading your Mistral API tier.

---

### 5. Office 2010 Specific Issues

**Prerequisites for Office 2010:**
- .NET Framework 4.8 Runtime installed.
- Ensure Visual Studio 2010 Tools for Office Runtime (VSTO Runtime) is present if custom ribbons fail to render.
