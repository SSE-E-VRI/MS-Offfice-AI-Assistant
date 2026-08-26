using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using MSOfficeAIAssistant.Core;
using Word = NetOffice.WordApi;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// Word host operations.  The context helpers deliberately keep their ranking logic
    /// separate from COM so it can be exercised without an Office installation.
    /// </summary>
    public class WordController : IOfficeHostController
    {
        private const int DefaultDocumentContextCharacters = 32000;

        // Store as raw object; wrap lazily only when needed.
        private readonly object _rawAppObj;
        private Word.Application _wordApp;

        public string HostType
        {
            get { return "Word"; }
        }

        public WordController(object appObj)
        {
            // NetOffice wrapping here can trigger COM event subscriptions on Word's UI thread.
            _rawAppObj = appObj;
        }

        private Word.Application GetApp()
        {
            if (_wordApp != null) return _wordApp;
            if (_rawAppObj == null) return null;

            try
            {
                if (_rawAppObj is Word.Application)
                {
                    _wordApp = (Word.Application)_rawAppObj;
                }
                else
                {
                    _wordApp = new Word.Application(null, _rawAppObj);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetApp failed: {0}", ex.Message));
            }
            return _wordApp;
        }

        public string GetSelectedText()
        {
            try
            {
                if (_rawAppObj != null)
                {
                    dynamic app = _rawAppObj;
                    dynamic selection = app.Selection;
                    if (selection != null)
                    {
                        int selType = Convert.ToInt32(selection.Type);
                        // wdSelectionIP = 1 (insertion point / blinking cursor only).
                        if (selType != 1)
                        {
                            string text = CleanWordText(Convert.ToString(selection.Text));
                            if (!string.IsNullOrEmpty(text)) return text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetSelectedText dynamic failed: {0}", ex.Message));
            }

            try
            {
                var app = GetApp();
                if (app != null && app.Selection != null &&
                    app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                {
                    string text = CleanWordText(app.Selection.Text);
                    if (!string.IsNullOrEmpty(text)) return text;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetSelectedText NetOffice failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public void InsertTextAtCursor(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    bool undoRecordStarted = TryStartUndoRecord(app, "AI Assistant insert");
                    try
                    {
                        WordMarkdownRenderer.Render(app, text);
                    }
                    finally
                    {
                        EndUndoRecord(app, undoRecordStarted);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.InsertTextAtCursor failed", ex);
                throw;
            }
        }

        public bool InsertText(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            try
            {
                var app = GetApp();
                if (app == null) return false;
                InsertTextAtCursor(text);
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.InsertText failed: {0}", ex.Message));
                return false;
            }
        }

        public bool Undo()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null && app.ActiveDocument != null)
                {
                    app.ActiveDocument.Undo();
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.Undo failed: {0}", ex.Message));
            }
            return false;
        }

        public string GetDocumentContext(string prompt, int maxCharacters)
        {
            return GetRelevantDocumentContext(prompt, maxCharacters > 0 ? maxCharacters : 24000);
        }

        public void ReplaceSelection(string text)
        {
            if (text == null) text = string.Empty;
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    bool undoRecordStarted = TryStartUndoRecord(app, "AI Assistant replace");
                    try
                    {
                        if (string.IsNullOrWhiteSpace(text))
                        {
                            DeleteCurrentSelection(app);
                        }
                        else
                        {
                            // The renderer clears a non-collapsed selection before it inserts the
                            // formatted Markdown content.  This avoids losing tables and lists.
                            WordMarkdownRenderer.Render(app, text);
                        }
                    }
                    finally
                    {
                        EndUndoRecord(app, undoRecordStarted);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ReplaceSelection failed", ex);
                throw;
            }
        }

        /// <summary>
        /// Inserts text with Word revision tracking enabled.  A non-collapsed selection is
        /// preserved and therefore replaced, which makes this suitable for AI rewrites.
        /// </summary>
        public bool ReplaceSelectionWithTrackChanges(string markdown)
        {
            return RenderWithTrackChanges(markdown, true);
        }

        /// <summary>
        /// Inserts text at the current insertion point with Word revision tracking enabled.
        /// Any existing selection is collapsed to its end instead of being replaced.
        /// </summary>
        public bool InsertTextAtCursorWithTrackChanges(string markdown)
        {
            return RenderWithTrackChanges(markdown, false);
        }

        // Kept for compatibility with existing callers.  The prior implementation inserted
        // plain text; delegating here retains Markdown formatting as tracked revisions.
        public int GetPendingRevisionCount()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null && app.ActiveDocument.Revisions != null)
                    return app.ActiveDocument.Revisions.Count;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetPendingRevisionCount failed: {0}", ex.Message));
            }
            return 0;
        }

        /// <summary>
        /// Accepts revisions in the current selection.  It intentionally does not accept the
        /// whole document when the selection is collapsed.
        /// </summary>
        public bool AcceptRevisionsInSelection()
        {
            return ApplySelectionRevisionDecision(true);
        }

        /// <summary>
        /// Rejects revisions in the current selection.  It intentionally does not reject the
        /// whole document when the selection is collapsed.
        /// </summary>
        public bool RejectRevisionsInSelection()
        {
            return ApplySelectionRevisionDecision(false);
        }

        /// <summary>
        /// Explicit whole-document revision actions for a deliberately chosen user command.
        /// UI callers should prefer the selection methods above whenever possible.
        /// </summary>
        public bool AcceptAllRevisions()
        {
            return ApplyDocumentRevisionDecision(true);
        }

        public bool RejectAllRevisions()
        {
            return ApplyDocumentRevisionDecision(false);
        }

        public HostOperationResult ExecuteInsertText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return HostOperationResult.Failed("Text to insert cannot be null or empty.");

            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                InsertTextAtCursor(text);
                return HostOperationResult.Ok("Text inserted at cursor.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertText failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertText");
            }
        }

        public HostOperationResult ExecuteReplaceSelection(string text)
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                ReplaceSelection(text);
                return HostOperationResult.Ok("Selection replaced.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteReplaceSelection failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteReplaceSelection");
            }
        }

        public HostOperationResult ExecuteInsertWithTrackChanges(string markdown, bool replaceSelection)
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                bool ok = RenderWithTrackChanges(markdown, replaceSelection);
                if (ok)
                    return HostOperationResult.Ok("Inserted text with Track Changes.");
                else
                    return HostOperationResult.Failed("Word Track Changes insertion returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertWithTrackChanges failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertWithTrackChanges");
            }
        }

        public HostOperationResult ExecuteAcceptAllRevisions()
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                bool ok = AcceptAllRevisions();
                return ok ? HostOperationResult.Ok("Accepted all revisions in Word document.") : HostOperationResult.Failed("AcceptAllRevisions returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteAcceptAllRevisions failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteAcceptAllRevisions");
            }
        }

        public HostOperationResult ExecuteRejectAllRevisions()
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                bool ok = RejectAllRevisions();
                return ok ? HostOperationResult.Ok("Rejected all revisions in Word document.") : HostOperationResult.Failed("RejectAllRevisions returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteRejectAllRevisions failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteRejectAllRevisions");
            }
        }

        // D-13 Tier 2: mutation methods that lacked a structured HostOperationResult wrapper.
        public HostOperationResult ExecuteAcceptRevisionsInSelection()
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                bool ok = AcceptRevisionsInSelection();
                return ok ? HostOperationResult.Ok("Accepted revisions in the current selection.") : HostOperationResult.Failed("AcceptRevisionsInSelection returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteAcceptRevisionsInSelection failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteAcceptRevisionsInSelection");
            }
        }

        public HostOperationResult ExecuteRejectRevisionsInSelection()
        {
            try
            {
                var app = GetApp();
                if (app == null)
                    return HostOperationResult.Failed("Word application is not accessible.");

                bool ok = RejectRevisionsInSelection();
                return ok ? HostOperationResult.Ok("Rejected revisions in the current selection.") : HostOperationResult.Failed("RejectRevisionsInSelection returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteRejectRevisionsInSelection failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteRejectRevisionsInSelection");
            }
        }

        public HostOperationResult ExecuteUndoLastChange()
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("Word application or active document is not accessible.");

                bool ok = UndoLastChange();
                return ok ? HostOperationResult.Ok("Undid the last Word change.") : HostOperationResult.Failed("UndoLastChange returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteUndoLastChange failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteUndoLastChange");
            }
        }

        public HostOperationResult ExecuteAddComment(string commentText, string targetText = null)
        {
            if (string.IsNullOrWhiteSpace(commentText))
                return HostOperationResult.Failed("Comment text cannot be empty.", 0, targetText);

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.", 0, targetText);

                dynamic doc = app.ActiveDocument;
                dynamic targetRange = null;

                // 1. If target text is provided, search for it in document
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    try
                    {
                        dynamic contentRange = doc.Content;
                        dynamic find = contentRange.Find;
                        find.ClearFormatting();
                        find.Text = targetText;
                        find.Forward = true;
                        find.Wrap = 0; // wdFindStop
                        if (find.Execute())
                        {
                            targetRange = contentRange;
                        }
                    }
                    catch { }
                }

                // 2. Fall back to current selection
                if (targetRange == null)
                {
                    try { targetRange = app.Selection.Range; } catch { }
                }

                // 3. Fall back to whole document content
                if (targetRange == null)
                {
                    try { targetRange = doc.Content; } catch { }
                }

                if (targetRange == null)
                    return HostOperationResult.Failed("Could not resolve a valid target range for the comment in Word.", 0, targetText);

                doc.Comments.Add(targetRange, commentText);
                return HostOperationResult.Ok("Comment added successfully", targetText);
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteAddComment failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteAddComment", targetText);
            }
        }

        public HostOperationResult ExecuteInsertTable(int rows, int cols, List<List<string>> data = null)
        {
            if (rows <= 0 || cols <= 0)
                return HostOperationResult.Failed(string.Format("Invalid table dimensions {0}x{1}.", rows, cols));

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                try { range = app.Selection != null ? app.Selection.Range : doc.Content; } catch { }
                if (range == null)
                {
                    try { range = doc.Content; } catch { }
                }
                if (range == null)
                    return HostOperationResult.Failed("Could not determine insertion range for Word table.");

                dynamic tables = doc.Tables;
                dynamic table = tables.Add(range, rows, cols, 1, 1);

                if (data != null && data.Count > 0)
                {
                    int rLimit = Math.Min(rows, data.Count);
                    for (int r = 0; r < rLimit; r++)
                    {
                        var rowData = data[r];
                        if (rowData == null) continue;
                        int cLimit = Math.Min(cols, rowData.Count);
                        for (int c = 0; c < cLimit; c++)
                        {
                            try
                            {
                                table.Cell(r + 1, c + 1).Range.Text = rowData[c] ?? string.Empty;
                            }
                            catch { }
                        }
                    }
                }

                return HostOperationResult.Ok(string.Format("Inserted {0}x{1} table into Word document", rows, cols));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertTable failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertTable");
            }
        }

        /// <summary>
        /// Uses Word's normal one-step undo stack.  This is deliberately a single, explicit
        /// undo rather than a broad rollback; it may undo the last document edit regardless of
        /// whether that edit came from the assistant or the user.
        /// </summary>
        public bool UndoLastChange()
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null) return false;
                app.ActiveDocument.Undo();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.UndoLastChange failed: {0}", ex.Message));
                return false;
            }
        }

        /// <summary>
        /// Legacy context entry point.  For long documents it now samples the entire document
        /// rather than always returning only the opening characters.
        /// </summary>
        public string GetDocumentText(int maxCharacters)
        {
            string fullText = GetFullDocumentText();
            return BuildRelevantDocumentContext(fullText, string.Empty, maxCharacters);
        }

        /// <summary>
        /// Produces a bounded, prompt-aware context from all of the active document.  It keeps
        /// document opening/closing context, ranks prompt-relevant passages, and includes text
        /// around the current cursor where possible.
        /// </summary>
        public string GetRelevantDocumentContext(string prompt, int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = DefaultDocumentContextCharacters;

            string fullText = GetFullDocumentText();
            if (string.IsNullOrEmpty(fullText)) return string.Empty;

            string cursorText = GetTextAroundSelection(Math.Min(1800, Math.Max(600, maxCharacters / 8)));
            string liveOutline = TryGetLiveDocumentOutline(Math.Min(2400, Math.Max(500, maxCharacters / 7)));
            string actionItems = PromptRequestsActionItems(prompt)
                ? BuildActionItemContext(fullText, Math.Min(2200, Math.Max(500, maxCharacters / 8)))
                : string.Empty;

            return WordDocumentContextBuilder.BuildRelevantDocumentContext(
                fullText, prompt, maxCharacters, cursorText, liveOutline, actionItems);
        }

        /// <summary>
        /// Returns Word heading structure when styles are available, otherwise a conservative
        /// text-derived outline.  Useful as grounding for outline, summary, and rewrite prompts.
        /// </summary>
        public string GetDocumentOutline(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 4000;

            string liveOutline = TryGetLiveDocumentOutline(maxCharacters);
            if (!string.IsNullOrEmpty(liveOutline)) return liveOutline;

            return BuildDocumentOutline(GetFullDocumentText(), maxCharacters);
        }

        /// <summary>
        /// Extracts plausible TODOs, owners, deadlines, and next steps from the active document.
        /// The output is context for an AI prompt, not a claim that every item is an assignment.
        /// </summary>
        public string GetActionItemContext(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 4000;
            return BuildActionItemContext(GetFullDocumentText(), maxCharacters);
        }

        /// <summary>
        /// Replaces a selected Markdown table with a native Word table.  The generated text is
        /// parsed first, so ordinary prose cannot accidentally be converted into a table.
        /// </summary>
        public string GetActiveDocumentName()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null)
                    return app.ActiveDocument.Name;
            }
            catch { }
            return "WordDocument";
        }

        public string GetContextReadout()
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null) return string.Empty;

                // Try to find the nearest heading above the cursor position
                dynamic selection = null;
                try { selection = app.Selection; } catch { }
                if (selection == null) return "Document: " + GetActiveDocumentName();

                object selStartObj = null;
                try { selStartObj = selection.Start; } catch { }
                int selStart = selStartObj != null ? Convert.ToInt32(selStartObj) : -1;

                if (selStart < 0)
                    return "Document: " + GetActiveDocumentName();

                // Scan paragraphs for headings before the cursor position
                int paragraphCount = 0;
                try { paragraphCount = app.ActiveDocument.Paragraphs.Count; } catch { }

                string lastHeadingFound = null;
                for (int index = 1; index <= Math.Min(200, paragraphCount); index++)
                {
                    try
                    {
                        Word.Paragraph paragraph = app.ActiveDocument.Paragraphs[index];
                        if (paragraph == null || paragraph.Range == null) continue;

                        object rangeStartObj = null;
                        try { rangeStartObj = paragraph.Range.Start; } catch { }
                        int rangeStart = rangeStartObj != null ? Convert.ToInt32(rangeStartObj) : -1;

                        // Only consider paragraphs before the cursor
                        if (rangeStart < 0 || rangeStart >= selStart) continue;

                        int outlineLevel = 0;
                        try { outlineLevel = Convert.ToInt32(paragraph.OutlineLevel, CultureInfo.InvariantCulture); } catch { }

                        string styleName = string.Empty;
                        try { styleName = Convert.ToString(paragraph.Style, CultureInfo.InvariantCulture); }
                        catch { }

                        bool isHeading = (outlineLevel >= 1 && outlineLevel <= 9);
                        if (!isHeading && styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                            isHeading = true;

                        if (isHeading)
                        {
                            string text = CleanWordText(paragraph.Range.Text);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                lastHeadingFound = text.Trim();
                            }
                        }
                    }
                    catch { }
                }

                if (!string.IsNullOrEmpty(lastHeadingFound))
                    return string.Format("Section: {0}", lastHeadingFound);

                return "Document: " + GetActiveDocumentName();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetContextReadout failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        // Pure, unit-testable helpers.  Keeping these public lets a lightweight test project
        // validate contextual retrieval without loading Word or NetOffice.
        public static string BuildRelevantDocumentContext(string documentText, string prompt, int maxCharacters)
        {
            return WordDocumentContextBuilder.BuildRelevantDocumentContext(
                documentText, prompt, maxCharacters, string.Empty, string.Empty, string.Empty);
        }

        public static string BuildDocumentOutline(string documentText, int maxCharacters)
        {
            return WordDocumentContextBuilder.BuildDocumentOutline(documentText, maxCharacters);
        }

        public static string BuildActionItemContext(string documentText, int maxCharacters)
        {
            return WordDocumentContextBuilder.BuildActionItemContext(documentText, maxCharacters);
        }

        private string GetFullDocumentText()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null && app.ActiveDocument.Content != null)
                    return CleanWordText(app.ActiveDocument.Content.Text);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetFullDocumentText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private string GetTextAroundSelection(int maxCharacters)
        {
            if (maxCharacters <= 0) return string.Empty;

            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.Selection == null ||
                    app.Selection.Range == null || app.ActiveDocument.Content == null)
                    return string.Empty;

                int contentStart = app.ActiveDocument.Content.Start;
                int contentEnd = app.ActiveDocument.Content.End;
                int selectionStart = app.Selection.Range.Start;
                int selectionEnd = app.Selection.Range.End;
                int padding = Math.Max(200, maxCharacters / 2);
                int start = Math.Max(contentStart, selectionStart - padding);
                int end = Math.Min(contentEnd, selectionEnd + padding);
                if (end <= start) return string.Empty;

                Word.Range contextRange = app.ActiveDocument.Range(start, end);
                return CleanWordText(contextRange.Text);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetTextAroundSelection failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private string TryGetLiveDocumentOutline(int maxCharacters)
        {
            if (maxCharacters <= 0) return string.Empty;

            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null)
                    return string.Empty;

                int totalParagraphs = 0;
                try { totalParagraphs = app.ActiveDocument.Paragraphs.Count; } catch { }

                // If document is larger than 200 paragraphs, COM iteration is slow in Word 2010.
                // Fall back directly to pure text-based outline generation.
                if (totalParagraphs > 200)
                {
                    try
                    {
                        string docText = CleanWordText(app.ActiveDocument.Content.Text);
                        return WordDocumentContextBuilder.BuildDocumentOutline(docText, maxCharacters);
                    }
                    catch { }
                }

                if (app.ActiveDocument.Paragraphs == null) return string.Empty;

                StringBuilder outline = new StringBuilder();
                int paragraphCount = Math.Min(200, totalParagraphs);
                for (int index = 1; index <= paragraphCount; index++)
                {
                    Word.Paragraph paragraph = app.ActiveDocument.Paragraphs[index];
                    if (paragraph == null || paragraph.Range == null) continue;

                    int outlineLevel = Convert.ToInt32(paragraph.OutlineLevel, CultureInfo.InvariantCulture);
                    string styleName = string.Empty;
                    try { styleName = Convert.ToString(paragraph.Style, CultureInfo.InvariantCulture); }
                    catch { }

                    bool isHeading = outlineLevel >= 1 && outlineLevel <= 9;
                    if (!isHeading && styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                        isHeading = true;
                    if (!isHeading) continue;

                    string text = CleanWordText(paragraph.Range.Text);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    text = WordDocumentContextBuilder.TrimToLength(text, 220);

                    int level = (outlineLevel >= 1 && outlineLevel <= 9)
                        ? outlineLevel
                        : WordDocumentContextBuilder.InferHeadingLevel(styleName);
                    string line = new string(' ', Math.Max(0, Math.Min(6, level - 1) * 2)) + "- " + text;
                    if (outline.Length + line.Length + 1 > maxCharacters) break;

                    if (outline.Length == 0) outline.Append("[Document outline]\n");
                    outline.Append(line).Append('\n');
                }

                if (outline.Length == 0)
                {
                    string docText = CleanWordText(app.ActiveDocument.Content.Text);
                    return WordDocumentContextBuilder.BuildDocumentOutline(docText, maxCharacters);
                }

                return outline.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.TryGetLiveDocumentOutline failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private bool RenderWithTrackChanges(string markdown, bool replaceSelection)
        {
            if (markdown == null) markdown = string.Empty;

            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.Selection == null) return false;

                bool wasTrackRevisions = app.ActiveDocument.TrackRevisions;
                bool undoRecordStarted = TryStartUndoRecord(app, "AI Assistant tracked edit");
                try
                {
                    app.ActiveDocument.TrackRevisions = true;
                    if (!replaceSelection) CollapseSelectionToEnd(app);

                    if (string.IsNullOrWhiteSpace(markdown))
                    {
                        if (replaceSelection) DeleteCurrentSelection(app);
                    }
                    else
                    {
                        // Render performs the selection replacement after TrackRevisions has been
                        // enabled, so both deletion and formatted insertion are reviewable.
                        WordMarkdownRenderer.Render(app, markdown);
                    }
                    return true;
                }
                finally
                {
                    app.ActiveDocument.TrackRevisions = wasTrackRevisions;
                    EndUndoRecord(app, undoRecordStarted);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.RenderWithTrackChanges failed", ex);
                throw;
            }
        }

        private bool ApplySelectionRevisionDecision(bool accept)
        {
            try
            {
                var app = GetApp();
                if (app == null || app.Selection == null ||
                    app.Selection.Type == Word.Enums.WdSelectionType.wdSelectionIP ||
                    app.Selection.Range == null || app.Selection.Range.Revisions == null)
                    return false;

                if (app.Selection.Range.Revisions.Count == 0) return false;
                if (accept) app.Selection.Range.Revisions.AcceptAll();
                else app.Selection.Range.Revisions.RejectAll();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ApplySelectionRevisionDecision failed", ex);
                throw;
            }
        }

        private bool ApplyDocumentRevisionDecision(bool accept)
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.ActiveDocument.Revisions == null)
                    return false;
                if (app.ActiveDocument.Revisions.Count == 0) return false;

                if (accept) app.ActiveDocument.AcceptAllRevisions();
                else app.ActiveDocument.RejectAllRevisions();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ApplyDocumentRevisionDecision failed", ex);
                throw;
            }
        }

        private static void CollapseSelectionToEnd(Word.Application app)
        {
            if (app != null && app.Selection != null &&
                app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                app.Selection.Collapse(Word.Enums.WdCollapseDirection.wdCollapseEnd);
        }

        private static void DeleteCurrentSelection(Word.Application app)
        {
            if (app != null && app.Selection != null &&
                app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                app.Selection.Range.Text = string.Empty;
        }

        private static bool TryStartUndoRecord(Word.Application app, string name)
        {
            return SafeOfficeProbe.Probe(() =>
            {
                if (app != null && app.UndoRecord != null)
                {
                    app.UndoRecord.StartCustomRecord(name);
                    return true;
                }
                return false;
            }, false, "Word.TryStartUndoRecord");
        }

        private static void EndUndoRecord(Word.Application app, bool recordStarted)
        {
            if (!recordStarted) return;
            SafeOfficeProbe.TryExecute(() =>
            {
                if (app != null && app.UndoRecord != null)
                    app.UndoRecord.EndCustomRecord();
            }, "Word.EndUndoRecord");
        }

        private static string CleanWordText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text.TrimEnd('\r', '\n', '\a', ' ', '\t');
        }

        private static bool PromptRequestsActionItems(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return false;
            string normalized = prompt.ToLowerInvariant();
            return normalized.IndexOf("action item", StringComparison.Ordinal) >= 0 ||
                   normalized.IndexOf("next step", StringComparison.Ordinal) >= 0 ||
                   normalized.IndexOf("to-do", StringComparison.Ordinal) >= 0 ||
                   normalized.IndexOf("todo", StringComparison.Ordinal) >= 0 ||
                   normalized.IndexOf("follow up", StringComparison.Ordinal) >= 0 ||
                   normalized.IndexOf("follow-up", StringComparison.Ordinal) >= 0;
        }

        /// <summary>
        /// Navigate to a specific paragraph by 1-based index (matching [¶N] citation numbering).
        /// Defensive: returns false on any error without throwing.
        /// </summary>
        public bool NavigateToParagraph(int paragraphIndex)
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null)
                    return false;

                if (paragraphIndex < 1 || paragraphIndex > app.ActiveDocument.Paragraphs.Count)
                    return false;

                Word.Paragraph paragraph = app.ActiveDocument.Paragraphs[paragraphIndex];
                if (paragraph == null || paragraph.Range == null)
                    return false;

                paragraph.Range.Select();

                // Attempt to scroll into view if the API is available
                try
                {
                    if (app.ActiveWindow != null)
                    {
                        app.ActiveWindow.ScrollIntoView(paragraph.Range, true);
                    }
                }
                catch
                {
                    // If ScrollIntoView is not available or fails, the selection alone is usually enough
                    // for Word to scroll to the current selection on its own.
                }

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.NavigateToParagraph failed: {0}", ex.Message));
                return false;
            }
        }
    }

    /// <summary>
    /// Pure document-context logic used by WordController.  It has no COM dependencies so the
    /// ranking, outline, and action-item behavior can be unit tested in isolation.
    /// </summary>
    internal static class WordDocumentContextBuilder
    {
        private const int DefaultContextCharacters = 32000;
        private const int ChunkTargetCharacters = 1500;

        private static readonly HashSet<string> StopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "about", "after", "again", "also", "am", "an", "and", "are", "as", "at", "be", "been", "before",
            "between", "by", "can", "could", "do", "does", "document", "for", "from", "get", "give", "has",
            "have", "help", "how", "i", "if", "in", "into", "is", "it", "make", "me", "my", "of", "on",
            "or", "please", "provide", "rewrite", "should", "show", "summarize", "summary", "that", "the", "their",
            "them", "then", "this", "to", "up", "use", "what", "when", "with", "would", "you", "your"
        };

        private sealed class ContextChunk
        {
            public int Index;
            public string Text;
            public int Score;
            public int StartLineIndex;
        }

        public static string BuildRelevantDocumentContext(
            string documentText,
            string prompt,
            int maxCharacters,
            string cursorText,
            string suppliedOutline,
            string actionItems)
        {
            if (maxCharacters <= 0) maxCharacters = DefaultContextCharacters;
            string normalizedDocument = NormalizeDocumentText(documentText);
            if (string.IsNullOrEmpty(normalizedDocument)) return string.Empty;

            // Sending a complete small document is both clearer and more accurate than making a
            // synthetic excerpt with duplicated headings.
            if (normalizedDocument.Length <= maxCharacters)
                return normalizedDocument;

            string normalizedCursor = NormalizeDocumentText(cursorText);
            string outline = NormalizeOutline(suppliedOutline);
            if (string.IsNullOrEmpty(outline))
                outline = BuildDocumentOutline(normalizedDocument, Math.Min(2200, Math.Max(450, maxCharacters / 7)));
            string normalizedActions = NormalizeActionItems(actionItems);

            int overheadBudget = Math.Max(0, maxCharacters / 3);
            outline = TrimToLength(outline, Math.Min(2200, overheadBudget));
            normalizedCursor = TrimToLength(normalizedCursor, Math.Min(1800, Math.Max(500, maxCharacters / 8)));
            normalizedActions = TrimToLength(normalizedActions, Math.Min(1800, Math.Max(500, maxCharacters / 8)));

            StringBuilder result = new StringBuilder(maxCharacters + 128);
            AppendContextSection(result, "[Document outline]", RemoveSectionHeader(outline), maxCharacters);
            AppendContextSection(result, "[Text around the cursor]", normalizedCursor, maxCharacters);
            AppendContextSection(result, "[Potential action items]", RemoveSectionHeader(normalizedActions), maxCharacters);

            int remainingForExcerpts = maxCharacters - result.Length;
            if (remainingForExcerpts < 300)
                return TrimToLength(result.ToString().Trim(), maxCharacters);

            List<ContextChunk> chunks = CreateChunks(normalizedDocument, remainingForExcerpts);
            List<string> terms = ExtractPromptTerms(prompt);
            ScoreChunks(chunks, terms, prompt);
            List<ContextChunk> selected = SelectChunks(chunks, terms.Count > 0, remainingForExcerpts);
            selected.Sort(delegate(ContextChunk left, ContextChunk right) { return left.Index.CompareTo(right.Index); });

            if (result.Length > 0) result.Append('\n');
            result.Append("[Relevant document excerpts; ").Append(chunks.Count)
                .Append(" sections scanned]\n");

            for (int i = 0; i < selected.Count && result.Length < maxCharacters; i++)
            {
                ContextChunk chunk = selected[i];
                string label = string.Format("[Excerpt {0} of {1}, ~Paragraph {2}]\n", chunk.Index + 1, chunks.Count, chunk.StartLineIndex + 1);
                int remaining = maxCharacters - result.Length;
                if (remaining <= label.Length + 20) break;

                int remainingItems = selected.Count - i;
                int share = Math.Max(250, (remaining - label.Length) / remainingItems);
                string excerpt = TrimToLength(chunk.Text, share);
                result.Append(label).Append(excerpt).Append('\n');
            }

            return TrimToLength(result.ToString().Trim(), maxCharacters);
        }

        public static string BuildDocumentOutline(string documentText, int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 4000;
            string normalized = NormalizeDocumentText(documentText);
            if (string.IsNullOrEmpty(normalized)) return string.Empty;

            string[] lines = normalized.Split(new[] { '\n' }, StringSplitOptions.None);
            StringBuilder outline = new StringBuilder();
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (string.IsNullOrEmpty(line)) continue;

                int level;
                string heading;
                if (!TryInferHeading(line, i + 1 < lines.Length && string.IsNullOrWhiteSpace(lines[i + 1]), out level, out heading))
                    continue;

                string entry = new string(' ', Math.Max(0, Math.Min(6, level - 1) * 2)) + "- " + heading;
                if (outline.Length + entry.Length + 1 > maxCharacters) break;
                outline.Append(entry).Append('\n');
            }

            if (outline.Length == 0) return string.Empty;
            return outline.ToString().TrimEnd();
        }

        public static string BuildActionItemContext(string documentText, int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 4000;
            string normalized = NormalizeDocumentText(documentText);
            if (string.IsNullOrEmpty(normalized)) return string.Empty;

            string[] candidates = Regex.Split(normalized, @"(?:\n+|(?<=[.!?])\s+)");
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            StringBuilder output = new StringBuilder();
            for (int i = 0; i < candidates.Length; i++)
            {
                string candidate = candidates[i].Trim();
                if (!LooksLikeActionItem(candidate)) continue;
                candidate = TrimToLength(candidate, 420);
                if (!seen.Add(candidate)) continue;

                string entry = "- " + candidate;
                if (output.Length + entry.Length + 1 > maxCharacters) break;
                output.Append(entry).Append('\n');
            }
            return output.ToString().TrimEnd();
        }

        public static int InferHeadingLevel(string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName)) return 1;
            Match match = Regex.Match(styleName, @"(\d+)");
            if (!match.Success) return 1;

            int level;
            if (Int32.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out level))
                return Math.Max(1, Math.Min(9, level));
            return 1;
        }

        public static string TrimToLength(string value, int maxCharacters)
        {
            if (string.IsNullOrEmpty(value) || maxCharacters <= 0) return string.Empty;
            if (value.Length <= maxCharacters) return value;
            if (maxCharacters <= 3) return value.Substring(0, maxCharacters);
            return value.Substring(0, maxCharacters - 3).TrimEnd() + "...";
        }

        private static string NormalizeDocumentText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\a", string.Empty);
            normalized = normalized.Replace('\v', '\n').Replace('\f', '\n');
            normalized = Regex.Replace(normalized, @"[ \t]+\n", "\n");
            normalized = Regex.Replace(normalized, @"\n{3,}", "\n\n");
            return normalized.Trim();
        }

        private static string NormalizeOutline(string outline)
        {
            return NormalizeDocumentText(outline);
        }

        private static string NormalizeActionItems(string actionItems)
        {
            return NormalizeDocumentText(actionItems);
        }

        private static string RemoveSectionHeader(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            const string outlineHeader = "[Document outline]";
            const string actionHeader = "[Potential action items]";
            if (value.StartsWith(outlineHeader, StringComparison.OrdinalIgnoreCase))
                return value.Substring(outlineHeader.Length).Trim();
            if (value.StartsWith(actionHeader, StringComparison.OrdinalIgnoreCase))
                return value.Substring(actionHeader.Length).Trim();
            return value;
        }

        private static void AppendContextSection(StringBuilder result, string heading, string value, int maxCharacters)
        {
            if (string.IsNullOrWhiteSpace(value) || result.Length >= maxCharacters) return;
            int available = maxCharacters - result.Length;
            if (available <= heading.Length + 2) return;
            if (result.Length > 0) result.Append('\n');
            result.Append(heading).Append('\n');
            available = maxCharacters - result.Length;
            result.Append(TrimToLength(value, available)).Append('\n');
        }

        private static List<ContextChunk> CreateChunks(string documentText, int availableCharacters)
        {
            int target = Math.Min(ChunkTargetCharacters, Math.Max(450, availableCharacters / 4));
            string[] lines = documentText.Split(new[] { '\n' }, StringSplitOptions.None);
            List<ContextChunk> chunks = new List<ContextChunk>();
            StringBuilder current = new StringBuilder();
            int chunkStartLineIndex = 0;

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                if (current.Length > 0 && current.Length + line.Length + 1 > target && current.Length >= 300)
                {
                    AddChunk(chunks, current, chunkStartLineIndex);
                    chunkStartLineIndex = i;
                }

                if (line.Length > target && current.Length == 0)
                {
                    int start = 0;
                    while (start < line.Length)
                    {
                        int length = Math.Min(target, line.Length - start);
                        chunks.Add(new ContextChunk { Index = chunks.Count, Text = line.Substring(start, length), StartLineIndex = i });
                        start += length;
                    }
                    chunkStartLineIndex = i + 1;
                    continue;
                }

                if (current.Length > 0) current.Append('\n');
                current.Append(line);
                if (string.IsNullOrWhiteSpace(line) && current.Length >= 300)
                {
                    AddChunk(chunks, current, chunkStartLineIndex);
                    chunkStartLineIndex = i + 1;
                }
            }
            AddChunk(chunks, current, chunkStartLineIndex);

            if (chunks.Count == 0)
                chunks.Add(new ContextChunk { Index = 0, Text = documentText, StartLineIndex = 0 });
            return chunks;
        }

        private static void AddChunk(List<ContextChunk> chunks, StringBuilder current, int startLineIndex)
        {
            string text = current.ToString().Trim();
            if (!string.IsNullOrEmpty(text))
                chunks.Add(new ContextChunk { Index = chunks.Count, Text = text, StartLineIndex = startLineIndex });
            current.Length = 0;
        }

        private static List<string> ExtractPromptTerms(string prompt)
        {
            List<string> terms = new List<string>();
            if (string.IsNullOrWhiteSpace(prompt)) return terms;

            MatchCollection matches = Regex.Matches(prompt.ToLowerInvariant(), @"[\p{L}\p{Nd}_-]{3,}");
            for (int i = 0; i < matches.Count; i++)
            {
                string term = matches[i].Value;
                if (StopWords.Contains(term) || terms.Contains(term)) continue;
                terms.Add(term);
            }
            return terms;
        }

        private static void ScoreChunks(List<ContextChunk> chunks, List<string> terms, string prompt)
        {
            string phrase = NormalizePromptPhrase(prompt);
            for (int i = 0; i < chunks.Count; i++)
            {
                string comparison = chunks[i].Text.ToLowerInvariant();
                int score = 0;
                for (int t = 0; t < terms.Count; t++)
                {
                    int occurrences = CountWholeTermOccurrences(comparison, terms[t]);
                    score += occurrences * (terms[t].Length >= 7 ? 6 : 4);
                }
                if (phrase.Length >= 8 && comparison.IndexOf(phrase, StringComparison.Ordinal) >= 0)
                    score += 12;

                // A heading-like opening is a useful tie breaker because it gives a relevant
                // passage its local section name when chunks are sorted back into document order.
                if (LooksLikeHeadingLine(GetFirstNonBlankLine(chunks[i].Text))) score += 1;
                chunks[i].Score = score;
            }
        }

        private static List<ContextChunk> SelectChunks(List<ContextChunk> chunks, bool hasTerms, int availableCharacters)
        {
            List<ContextChunk> selected = new List<ContextChunk>();
            if (chunks.Count == 0) return selected;

            int maxChunkCount = Math.Max(2, Math.Min(chunks.Count, Math.Max(2, availableCharacters / 900)));
            List<ContextChunk> ranked = new List<ContextChunk>(chunks);
            ranked.Sort(delegate(ContextChunk left, ContextChunk right)
            {
                int scoreComparison = right.Score.CompareTo(left.Score);
                return scoreComparison != 0 ? scoreComparison : left.Index.CompareTo(right.Index);
            });

            if (hasTerms)
            {
                // A narrow context budget must still answer the user's question.  Rank relevant
                // passages before reserving opening/closing document samples; otherwise a two-
                // chunk budget would always become a beginning/end prefix again.
                for (int i = 0; i < ranked.Count && selected.Count < maxChunkCount; i++)
                {
                    if (ranked[i].Score <= 0) break;
                    AddUnique(selected, ranked[i]);
                    if (selected.Count < maxChunkCount && ranked[i].Index > 0)
                        AddUnique(selected, chunks[ranked[i].Index - 1]);
                    if (selected.Count < maxChunkCount && ranked[i].Index + 1 < chunks.Count)
                        AddUnique(selected, chunks[ranked[i].Index + 1]);
                }
            }

            AddUniqueWhenCapacityRemains(selected, chunks[0], maxChunkCount);
            if (chunks.Count > 1)
                AddUniqueWhenCapacityRemains(selected, chunks[chunks.Count - 1], maxChunkCount);

            // When there is no useful query term (or no match), sample evenly across the entire
            // document.  This is what prevents a long document from degenerating into a prefix.
            int sampleOrdinal = 1;
            while (selected.Count < maxChunkCount)
            {
                int denominator = Math.Max(1, maxChunkCount - 1);
                int index = (int)Math.Round((double)sampleOrdinal * (chunks.Count - 1) / denominator);
                AddUnique(selected, chunks[Math.Max(0, Math.Min(chunks.Count - 1, index))]);
                sampleOrdinal++;
                if (sampleOrdinal > chunks.Count + maxChunkCount) break;
            }

            return selected;
        }

        private static void AddUniqueWhenCapacityRemains(List<ContextChunk> selected, ContextChunk candidate, int maximum)
        {
            if (selected.Count >= maximum) return;
            AddUnique(selected, candidate);
        }

        private static void AddUnique(List<ContextChunk> selected, ContextChunk candidate)
        {
            for (int i = 0; i < selected.Count; i++)
            {
                if (selected[i].Index == candidate.Index) return;
            }
            selected.Add(candidate);
        }

        private static int CountWholeTermOccurrences(string value, string term)
        {
            if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(term)) return 0;
            int count = 0;
            int start = 0;
            while (start < value.Length)
            {
                int index = value.IndexOf(term, start, StringComparison.Ordinal);
                if (index < 0) break;
                int end = index + term.Length;
                bool leftBoundary = index == 0 || !IsTermCharacter(value[index - 1]);
                bool rightBoundary = end >= value.Length || !IsTermCharacter(value[end]);
                if (leftBoundary && rightBoundary) count++;
                start = end;
            }
            return count;
        }

        private static bool IsTermCharacter(char value)
        {
            return Char.IsLetterOrDigit(value) || value == '_' || value == '-';
        }

        private static string NormalizePromptPhrase(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt)) return string.Empty;
            string normalized = Regex.Replace(prompt.ToLowerInvariant(), @"\s+", " ").Trim();
            return normalized.Length > 80 ? normalized.Substring(0, 80) : normalized;
        }

        private static bool TryInferHeading(string line, bool followedByBlankLine, out int level, out string heading)
        {
            level = 1;
            heading = string.Empty;
            if (string.IsNullOrWhiteSpace(line) || line.Length > 180) return false;

            Match markdown = Regex.Match(line, @"^(#{1,6})\s+(.+?)\s*#*$");
            if (markdown.Success)
            {
                level = markdown.Groups[1].Value.Length;
                heading = markdown.Groups[2].Value.Trim();
                return true;
            }

            Match numbered = Regex.Match(line, @"^(\d+(?:\.\d+){0,5})[.)]?\s+(.+)$");
            if (numbered.Success)
            {
                level = numbered.Groups[1].Value.Split('.').Length;
                heading = line;
                return true;
            }

            if (followedByBlankLine && LooksLikeHeadingLine(line))
            {
                heading = line;
                return true;
            }
            return false;
        }

        private static bool LooksLikeHeadingLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line) || line.Length > 120) return false;
            string trimmed = line.Trim();
            if (trimmed.EndsWith(".", StringComparison.Ordinal) || trimmed.EndsWith("?", StringComparison.Ordinal)) return false;
            if (Regex.IsMatch(trimmed, @"^#{1,6}\s+")) return true;

            int letterCount = 0;
            int upperCount = 0;
            for (int i = 0; i < trimmed.Length; i++)
            {
                if (!Char.IsLetter(trimmed[i])) continue;
                letterCount++;
                if (Char.IsUpper(trimmed[i])) upperCount++;
            }
            return letterCount >= 3 && upperCount * 100 / letterCount >= 70;
        }

        private static string GetFirstNonBlankLine(string value)
        {
            string[] lines = value.Split(new[] { '\n' }, StringSplitOptions.None);
            for (int i = 0; i < lines.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(lines[i])) return lines[i].Trim();
            }
            return string.Empty;
        }

        private static bool LooksLikeActionItem(string value)
        {
            if (string.IsNullOrWhiteSpace(value) || value.Length > 800) return false;
            string normalized = value.ToLowerInvariant();
            if (Regex.IsMatch(normalized, @"^\s*(?:[-*•]|\[[ x]\])\s*(?:todo|to-do|action|next step|follow[- ]?up)")) return true;

            string[] cues =
            {
                "action item", "todo", "to-do", "next step", "follow up", "follow-up", "due ",
                "deadline", "owner", "responsible for", "needs to", "need to", "must ", "will ", "assign "
            };
            for (int i = 0; i < cues.Length; i++)
            {
                if (normalized.IndexOf(cues[i], StringComparison.Ordinal) >= 0) return true;
            }
            return false;
        }
    }
}
