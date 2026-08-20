# AI Assistant for Microsoft Office — User Guide

**Version 0.4.0**
*Designed and developed by D.Manikandan B.E, SSE/E/VRI, Mob No 9444861302*

> **This guide is now the in-app User Manual.**
> Open Word, Excel, or PowerPoint, switch to the **AI Assistant** ribbon tab, and click
> **User Manual** (Help group) — or press the **?** button in the top-right of the chat panel.
> The manual is embedded in the add-in itself, so it ships with every install and opens even
> when no API key has been configured yet.

The manual covers:

1. Features available (BYOK providers, streaming, attachments, context scope, Word/Excel/PowerPoint
   capabilities, DPAPI-encrypted local storage, and what is *not* included).
2. How to configure your API key (Configure / ⚙, Test Connection, Save, encryption path).
3. How to get a free API key (Mistral, Groq, Gemini consoles; Custom/Ollama as a zero-cloud option).
4. Example prompts for Word (drafting, rewriting, summarizing, action items, bilingual notices, Track edits).
5. Example prompts for Excel (always Review &amp; Apply — never silent workbook writes).
6. Example prompts for PowerPoint (deck build, review, reorganising, Insert image = local file).
7. A catalogue of every ribbon control (Chat / Draft / More / Help groups).
8. A catalogue of every chat-panel control (header, toolbar, message actions, input area).
9. Indian Railways office workflows (safety circulars, inspection notes, estimates/BOQ, MOM,
   briefing decks, bilingual notices) with copy-paste prompts and the required safety instructions.
10. Installation and a pointer to troubleshooting.

For the developer-facing feature notes see [PHASE_1_TO_4_FEATURES.md](PHASE_1_TO_4_FEATURES.md),
and for common issues see [TROUBLESHOOTING.md](TROUBLESHOOTING.md).

The offline HTML source of the manual lives at `src/Help/UserManual.html` and is embedded into the
add-in as a manifest resource (same mechanism as `src/Addin/Ribbon.xml`).