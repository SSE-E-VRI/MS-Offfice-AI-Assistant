# Phases 1-4: Implemented Office Workflows

This release intentionally stops before Phase 5. It does not add Microsoft Graph,
tenant search, SharePoint/OneDrive retrieval, or enterprise connector integration.

## Safe AI workflow (Phase 1)

- Select exactly what is sent with a prompt: **Selection**, **Current file**,
  **Selection + file**, or **Attachments only**.
- Review a response before insertion and explicitly confirm every native Office edit.
- Use the **Undo** button for the host application's most recent undoable change.
- Use **Log** to inspect the recent, user-approved AI action trail.
- Conversation history and the action trail are stored locally with Windows DPAPI
  encryption for the current Windows user.

## Word (Phase 2)

- Whole-document, prompt-aware context retrieves relevant passages, current cursor
  context, an outline, and action-item context when useful.
- Enable **Track edits** before inserting or replacing an AI response. Use **Accept**
  or **Reject** to decide selected revisions, or all revisions when there is no
  selection.
- Markdown tables in AI output render as native Word tables. Selected Markdown,
  tab-delimited, and pipe-delimited text can be converted through the Word host APIs.
- Attached text sources are labelled for source-aware responses.

## Excel (Phase 3)

Excel responses can propose a previewable, confirm-before-apply action for a bounded
A1 range. The native action set includes formulas, values, fill-down, table writes,
Excel Table creation, conditional formatting, sorting, filtering, data validation,
charts, PivotTables, named ranges, and duplicate removal.

The assistant returns those changes in an `excel_actions` block. Each card shows the
target, operation, description, and proposed configuration before it can change the
workbook.

## PowerPoint (Phase 4)

- Prompts can use full-deck text, section names, slide titles, and speaker notes,
  rather than only the active slide.
- Review prompts include deterministic flags for untitled slides, duplicate titles,
  and text-light slides.
- Insert a structured outline to update the active slide and create later slides using
  the current presentation layout where possible.
- Supported safe deck operations are moving slides, creating/renaming sections, and
  setting speaker notes.
- Generated visual suggestions are preserved in speaker notes. Use **Insert image**
  to choose and explicitly approve a local image for the active slide; it receives
  accessible alt text derived from its file name.
