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

        // Bookmarks used to pin a Selection-scope chat exchange to the exact range the user
        // had selected, so Insert can replace that text later even if the live selection has
        // since moved (the user scrolled, clicked elsewhere, or is just slow to click Insert).
        // Tracked here so a long session doesn't quietly accumulate one invisible bookmark per
        // question the user never inserted.
        private readonly List<string> _sourceSelectionBookmarks = new List<string>();
        private int _sourceSelectionBookmarkCounter;
        private const int MaxTrackedSourceBookmarks = 30;
        private const string SourceBookmarkPrefix = "AIAsstSel";

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

        /// <summary>
        /// Inserts text with Word revision tracking enabled.  A non-collapsed selection is
        /// preserved and therefore replaced, which makes this suitable for AI rewrites.
        /// </summary>
        /// <summary>
        /// Pins the current selection with a hidden Word bookmark so a later Insert can replace
        /// exactly this text, even if the user has since clicked or scrolled elsewhere in the
        /// document while the response was streaming in. Returns null when there is nothing to
        /// pin (no real selection, or the app/document is not accessible) -- callers should treat
        /// null as "fall back to whatever is live-selected when Insert is actually clicked",
        /// which was the previous, only, behavior.
        /// </summary>
        public string CreateSelectionBookmark()
        {
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.Selection == null) return null;
                if (app.Selection.Type == Word.Enums.WdSelectionType.wdSelectionIP) return null;

                string name = SourceBookmarkPrefix + (++_sourceSelectionBookmarkCounter).ToString(CultureInfo.InvariantCulture);
                app.ActiveDocument.Bookmarks.Add(name, app.Selection.Range);
                _sourceSelectionBookmarks.Add(name);
                TrimTrackedSourceBookmarks(app);
                return name;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.CreateSelectionBookmark failed: {0}", ex.Message));
                return null;
            }
        }

        /// <summary>
        /// Selects the range a prior CreateSelectionBookmark call pinned, so the next tracked or
        /// plain insert replaces that exact text instead of whatever happens to be selected now.
        /// Returns false -- leaving the current selection untouched -- if the bookmark is gone,
        /// which happens once the pinned text itself has been edited or deleted; callers should
        /// fall back to the live selection in that case, same as when no bookmark was ever pinned.
        /// </summary>
        public bool TrySelectSourceBookmark(string bookmarkName)
        {
            if (string.IsNullOrEmpty(bookmarkName)) return false;
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.ActiveDocument.Bookmarks == null) return false;
                if (!app.ActiveDocument.Bookmarks.Exists(bookmarkName)) return false;

                app.ActiveDocument.Bookmarks[bookmarkName].Range.Select();
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.TrySelectSourceBookmark failed: {0}", ex.Message));
                return false;
            }
        }

        /// <summary>
        /// Best-effort cleanup once a pinned selection has been consumed by Insert (or abandoned,
        /// e.g. the user asked a follow-up instead). Safe to call even if the bookmark is already
        /// gone -- inserting text over a bookmarked range often removes it implicitly.
        /// </summary>
        public void ForgetSourceBookmark(string bookmarkName)
        {
            if (string.IsNullOrEmpty(bookmarkName)) return;
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null && app.ActiveDocument.Bookmarks != null
                    && app.ActiveDocument.Bookmarks.Exists(bookmarkName))
                {
                    app.ActiveDocument.Bookmarks[bookmarkName].Delete();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.ForgetSourceBookmark failed: {0}", ex.Message));
            }
            _sourceSelectionBookmarks.Remove(bookmarkName);
        }

        /// <summary>Caps how many pinned-but-never-inserted bookmarks a long session can leave behind.</summary>
        private void TrimTrackedSourceBookmarks(Word.Application app)
        {
            while (_sourceSelectionBookmarks.Count > MaxTrackedSourceBookmarks)
            {
                string oldest = _sourceSelectionBookmarks[0];
                _sourceSelectionBookmarks.RemoveAt(0);
                try
                {
                    if (app.ActiveDocument.Bookmarks.Exists(oldest))
                        app.ActiveDocument.Bookmarks[oldest].Delete();
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("WordController.TrimTrackedSourceBookmarks failed: {0}", ex.Message));
                }
            }
        }

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

        public HostOperationResult ExecuteFindReplace(string findText, string replaceText)
        {
            if (string.IsNullOrWhiteSpace(findText))
                return HostOperationResult.Failed("Find text cannot be empty.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic range = doc.Content;
                dynamic find = range.Find;
                find.ClearFormatting();
                try { find.Replacement.ClearFormatting(); } catch { }
                find.Text = findText;
                try { find.Replacement.Text = replaceText ?? string.Empty; } catch { }
                find.Forward = true;
                find.Wrap = 1; // wdFindContinue
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;
                bool found = find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2); // wdReplaceAll
                if (found)
                    return HostOperationResult.Ok(string.Format("Replaced '{0}' with '{1}'.", findText, replaceText ?? string.Empty));
                else
                    return HostOperationResult.Ok(string.Format("Text '{0}' not found; no changes made.", findText));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteFindReplace failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteFindReplace");
            }
        }

        public HostOperationResult ExecuteApplyStyle(int paragraphIndex, string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                return HostOperationResult.Failed("Style name cannot be empty.");
            if (paragraphIndex < 1)
                return HostOperationResult.Failed("Paragraph index must be at least 1.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0;
                try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                if (paragraphIndex > count)
                    return HostOperationResult.Failed(string.Format("Paragraph {0} does not exist (document has {1} paragraphs).", paragraphIndex, count));

                dynamic para = doc.Paragraphs[paragraphIndex];
                string trimmedStyle = styleName.Trim();
                try
                {
                    // Try via Styles collection first for exact match
                    dynamic styleObj = doc.Styles[trimmedStyle];
                    if (styleObj != null) para.Range.Style = styleObj;
                    else para.Range.Style = trimmedStyle;
                }
                catch
                {
                    para.Range.Style = trimmedStyle;
                }
                return HostOperationResult.Ok(string.Format("Applied style '{0}' to paragraph {1}.", trimmedStyle, paragraphIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteApplyStyle failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteApplyStyle");
            }
        }

        public HostOperationResult ExecuteApplyStyleByText(string targetText, string styleName)
        {
            if (string.IsNullOrWhiteSpace(styleName))
                return HostOperationResult.Failed("Style name cannot be empty.");
            if (string.IsNullOrWhiteSpace(targetText))
                return ExecuteApplyStyle(1, styleName);

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0;
                try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                int foundIndex = -1;
                string lowerTarget = targetText.Trim().ToLowerInvariant();
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        string paraText = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                        if (paraText.ToLowerInvariant().IndexOf(lowerTarget, StringComparison.Ordinal) >= 0)
                        {
                            foundIndex = i;
                            break;
                        }
                    }
                    catch { }
                }
                if (foundIndex < 0)
                    return HostOperationResult.Failed(string.Format("Target text '{0}' not found in document.", targetText));

                return ExecuteApplyStyle(foundIndex, styleName);
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteApplyStyleByText failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteApplyStyleByText");
            }
        }

        public HostOperationResult ExecuteSetCase(string target, string caseType)
        {
            if (string.IsNullOrWhiteSpace(caseType))
                return HostOperationResult.Failed("Case type is required (title, sentence, upper, lower).");
            string normalized = caseType.Trim().ToLowerInvariant();
            bool isTitle = normalized == "title" || normalized == "title_case" || normalized == "titlecase";
            bool isSentence = normalized == "sentence" || normalized == "sentence_case";
            bool isUpper = normalized == "upper" || normalized == "upper_case" || normalized == "uppercase";
            bool isLower = normalized == "lower" || normalized == "lower_case" || normalized == "lowercase";
            if (!isTitle && !isSentence && !isUpper && !isLower)
                return HostOperationResult.Failed("Case type must be one of: title, sentence, upper, lower.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                List<int> targetParagraphs = new List<int>();

                if (!string.IsNullOrWhiteSpace(target))
                {
                    string lowerTarget = target.Trim().ToLowerInvariant();
                    int count = 0;
                    try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            string paraText = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                            if (paraText.ToLowerInvariant().IndexOf(lowerTarget, StringComparison.Ordinal) >= 0)
                            {
                                targetParagraphs.Add(i);
                                break; // only first matching paragraph for targeted mode
                            }
                        }
                        catch { }
                    }
                    if (targetParagraphs.Count == 0)
                        return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", target));
                }
                else
                {
                    // No target -> apply to selection if non-empty, else whole document
                    try
                    {
                        if (_rawAppObj != null)
                        {
                            dynamic selApp = _rawAppObj;
                            dynamic sel = selApp.Selection;
                            if (sel != null && Convert.ToInt32(sel.Type) != 1)
                            {
                                dynamic selRange = sel.Range;
                                if (selRange != null)
                                {
                                    string selText = Convert.ToString(selRange.Text) ?? string.Empty;
                                    if (!string.IsNullOrWhiteSpace(selText))
                                    {
                                        string converted = ConvertCase(selText, isTitle, isSentence, isUpper, isLower);
                                        selRange.Text = converted;
                                        return HostOperationResult.Ok(string.Format("Changed case to {0} for selection.", normalized));
                                    }
                                }
                            }
                        }
                    }
                    catch { }

                    // Fall back to whole document
                    int count = 0;
                    try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                    for (int i = 1; i <= count; i++) targetParagraphs.Add(i);
                }

                // Apply to collected paragraph indices (for targeted case, single element)
                foreach (int idx in targetParagraphs)
                {
                    try
                    {
                        dynamic para = doc.Paragraphs[idx];
                        string original = Convert.ToString(para.Range.Text) ?? string.Empty;
                        // Preserve trailing paragraph mark \r
                        bool hasMark = original.EndsWith("\r");
                        string core = hasMark ? original.Substring(0, original.Length - 1) : original;
                        string converted = ConvertCase(core, isTitle, isSentence, isUpper, isLower);
                        if (hasMark) converted = converted + "\r";
                        para.Range.Text = converted;
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("ExecuteSetCase paragraph {0} failed: {1}", idx, ex.Message));
                    }
                }

                if (targetParagraphs.Count == 1)
                    return HostOperationResult.Ok(string.Format("Changed paragraph {0} to {1} case.", targetParagraphs[0], normalized));
                else
                    return HostOperationResult.Ok(string.Format("Changed {0} paragraphs to {1} case.", targetParagraphs.Count, normalized));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteSetCase failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteSetCase");
            }
        }

        private static string ConvertCase(string text, bool isTitle, bool isSentence, bool isUpper, bool isLower)
        {
            if (string.IsNullOrEmpty(text)) return text;
            if (isUpper) return text.ToUpperInvariant();
            if (isLower) return text.ToLowerInvariant();
            if (isTitle)
            {
                System.Globalization.TextInfo ti = System.Globalization.CultureInfo.CurrentCulture.TextInfo;
                return ti.ToTitleCase(text.ToLowerInvariant());
            }
            if (isSentence)
            {
                // Sentence case: first letter upper, rest lower per sentence
                string lower = text.ToLowerInvariant();
                System.Text.StringBuilder sb = new System.Text.StringBuilder(lower.Length);
                bool capNext = true;
                for (int i = 0; i < lower.Length; i++)
                {
                    char c = lower[i];
                    if (capNext && char.IsLetter(c))
                    {
                        sb.Append(char.ToUpperInvariant(c));
                        capNext = false;
                    }
                    else
                    {
                        sb.Append(c);
                    }
                    if (c == '.' || c == '!' || c == '?') capNext = true;
                }
                return sb.ToString();
            }
            return text;
        }

        public HostOperationResult ExecuteReorganizeParagraphs(string orderCsv)
        {
            if (string.IsNullOrWhiteSpace(orderCsv))
                return HostOperationResult.Failed("Order is required (e.g., 3,1,2).");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0;
                try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                if (count < 2)
                    return HostOperationResult.Failed("Document has fewer than 2 paragraphs to reorder.");

                string[] parts = orderCsv.Split(new char[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                List<int> order = new List<int>();
                foreach (string p in parts)
                {
                    int v;
                    if (!int.TryParse(p.Trim(), out v)) return HostOperationResult.Failed(string.Format("Invalid order index '{0}'.", p));
                    if (v < 1 || v > count) return HostOperationResult.Failed(string.Format("Order index {0} out of range 1..{1}.", v, count));
                    if (order.Contains(v)) return HostOperationResult.Failed(string.Format("Duplicate order index {0}.", v));
                    order.Add(v);
                }
                if (order.Count < 2)
                    return HostOperationResult.Failed("Order must contain at least 2 indices.");
                if (order.Count > count)
                    return HostOperationResult.Failed(string.Format("Order count {0} exceeds paragraph count {1}.", order.Count, count));

                // Capture original paragraph texts
                List<string> original = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        string t = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                        original.Add(t);
                    }
                    catch { original.Add(string.Empty); }
                }

                // Build new order: specified indices first, then remaining in original order
                List<string> reordered = new List<string>();
                HashSet<int> used = new HashSet<int>();
                foreach (int idx in order)
                {
                    reordered.Add(original[idx - 1]);
                    used.Add(idx);
                }
                // Append untouched paragraphs in original order if partial reorder
                if (order.Count < count)
                {
                    for (int i = 1; i <= count; i++)
                    {
                        if (!used.Contains(i)) reordered.Add(original[i - 1]);
                    }
                }

                // Apply back
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic para = doc.Paragraphs[i];
                        para.Range.Text = reordered[i - 1];
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("Reorganize paragraph {0} failed: {1}", i, ex.Message));
                    }
                }

                return HostOperationResult.Ok(string.Format("Reordered {0} paragraphs.", count));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteReorganizeParagraphs failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteReorganizeParagraphs");
            }
        }

        public HostOperationResult ExecuteNormalizeWhitespace()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0;
                try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                if (count == 0) return HostOperationResult.Ok("Document is empty.");

                int changed = 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic para = doc.Paragraphs[i];
                        string original = Convert.ToString(para.Range.Text) ?? string.Empty;
                        bool hasMark = original.EndsWith("\r");
                        string core = hasMark ? original.Substring(0, original.Length - 1) : original;
                        string normalized = System.Text.RegularExpressions.Regex.Replace(core, @"[ \t]{2,}", " ");
                        normalized = normalized.Trim();
                        normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"\s+$", string.Empty);
                        if (normalized != core)
                        {
                            if (hasMark) normalized = normalized + "\r";
                            para.Range.Text = normalized;
                            changed++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("Normalize paragraph {0} failed: {1}", i, ex.Message));
                    }
                }

                // Remove consecutive blank paragraphs via Find
                try
                {
                    dynamic range = doc.Content;
                    dynamic find = range.Find;
                    find.ClearFormatting();
                    try { find.Replacement.ClearFormatting(); } catch { }
                    find.Text = "^p^p";
                    find.Replacement.Text = "^p";
                    find.Forward = true;
                    find.Wrap = 1;
                    find.Format = false;
                    // Execute twice to collapse triple blanks
                    find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                    find.Execute(Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 2);
                }
                catch { }

                return HostOperationResult.Ok(string.Format("Normalized whitespace in {0} paragraphs.", changed));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteNormalizeWhitespace failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteNormalizeWhitespace");
            }
        }

        public HostOperationResult ExecuteSetFont(string targetText, string fontName, string fontSizeStr, string boldStr, string italicStr, string underlineStr, string color, string highlight)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    // Locate first occurrence of targetText
                    dynamic doc = app.ActiveDocument;
                    int count = 0;
                    try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                    bool found = false;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            string paraText = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                            if (paraText.IndexOf(targetText, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                range = doc.Paragraphs[i].Range;
                                found = true;
                                break;
                            }
                        }
                        catch { }
                    }
                    if (!found)
                    {
                        // Try document Find
                        try
                        {
                            range = doc.Content;
                            dynamic find = range.Find;
                            find.ClearFormatting(); try { find.Replacement.ClearFormatting(); } catch { }
                            find.Text = targetText;
                            find.Forward = true; find.Wrap = 0; find.Format = false;
                            bool ok = find.Execute();
                            if (!ok) return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText));
                            // after Execute, range is at found location
                        }
                        catch { return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText)); }
                    }
                }
                else
                {
                    try
                    {
                        dynamic selApp = _rawAppObj;
                        dynamic sel = selApp.Selection;
                        if (sel != null && Convert.ToInt32(sel.Type) != 1 && sel.Range != null) range = sel.Range;
                        else range = app.ActiveDocument.Content;
                    }
                    catch { range = app.ActiveDocument.Content; }
                }
                if (range == null) return HostOperationResult.Failed("Could not resolve target range for font change.");

                dynamic font = range.Font;
                if (!string.IsNullOrWhiteSpace(fontName)) try { font.Name = fontName.Trim(); } catch { }
                if (!string.IsNullOrWhiteSpace(fontSizeStr))
                {
                    float sz; if (float.TryParse(fontSizeStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out sz) && sz >= 6 && sz <= 72) try { font.Size = sz; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(boldStr))
                {
                    bool b; if (bool.TryParse(boldStr, out b)) try { font.Bold = b ? 1 : 0; } catch { }
                    else if (boldStr.Trim() == "1" || boldStr.Trim().ToLowerInvariant() == "true") try { font.Bold = 1; } catch { }
                    else if (boldStr.Trim() == "0" || boldStr.Trim().ToLowerInvariant() == "false") try { font.Bold = 0; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(italicStr))
                {
                    bool b; if (bool.TryParse(italicStr, out b)) try { font.Italic = b ? 1 : 0; } catch { }
                    else if (italicStr.Trim().ToLowerInvariant() == "true") try { font.Italic = 1; } catch { }
                    else if (italicStr.Trim().ToLowerInvariant() == "false") try { font.Italic = 0; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(underlineStr))
                {
                    string u = underlineStr.Trim().ToLowerInvariant();
                    try
                    {
                        if (u == "true" || u == "1" || u == "single") font.Underline = 1; // wdUnderlineSingle
                        else if (u == "false" || u == "0" || u == "none") font.Underline = 0;
                    }
                    catch { }
                }
                if (!string.IsNullOrWhiteSpace(color))
                {
                    try { font.Color = HexToWordColor(color.Trim()); } catch { }
                }
                if (!string.IsNullOrWhiteSpace(highlight))
                {
                    try
                    {
                        int hi = ParseHighlightIndex(highlight.Trim());
                        if (hi >= 0) range.HighlightColorIndex = hi;
                    }
                    catch { }
                }
                return HostOperationResult.Ok("Font formatting applied.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteSetFont failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteSetFont");
            }
        }

        private static int HexToWordColor(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return 0;
            string h = hex.Trim().TrimStart('#');
            if (h.Length != 6) return 0;
            int r = Convert.ToInt32(h.Substring(0, 2), 16);
            int g = Convert.ToInt32(h.Substring(2, 2), 16);
            int b = Convert.ToInt32(h.Substring(4, 2), 16);
            // Word VBA RGB: R + G*256 + B*65536 (BGR)
            return r + (g * 256) + (b * 65536);
        }

        private static int ParseHighlightIndex(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return -1;
            string n = name.Trim().ToLowerInvariant();
            if (n == "none" || n == "0") return 0;
            if (n == "yellow") return 7;
            if (n == "green") return 4;
            if (n == "cyan") return 8;
            if (n == "magenta") return 5;
            if (n == "blue") return 2;
            if (n == "red") return 6;
            if (n == "darkblue") return 9;
            if (n == "darkyellow") return 14;
            int v; if (int.TryParse(n, out v) && v >= 0 && v <= 16) return v;
            return -1;
        }

        public HostOperationResult ExecuteSetParagraphFormat(string targetText, string alignment, string lineSpacingStr, string spaceBeforeStr, string spaceAfterStr, string leftIndentStr, string firstLineIndentStr)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    dynamic doc = app.ActiveDocument;
                    int count = 0; try { count = Convert.ToInt32(doc.Paragraphs.Count); } catch { }
                    bool found = false;
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            string paraText = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                            if (paraText.IndexOf(targetText, StringComparison.OrdinalIgnoreCase) >= 0) { range = doc.Paragraphs[i].Range; found = true; break; }
                        }
                        catch { }
                    }
                    if (!found) return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText));
                }
                else
                {
                    try
                    {
                        dynamic selApp = _rawAppObj;
                        dynamic sel = selApp.Selection;
                        if (sel != null && Convert.ToInt32(sel.Type) != 1 && sel.Range != null) range = sel.Range;
                        else range = app.ActiveDocument.Content;
                    }
                    catch { range = app.ActiveDocument.Content; }
                }
                if (range == null) return HostOperationResult.Failed("Could not resolve target range.");

                dynamic pf = range.ParagraphFormat;
                if (!string.IsNullOrWhiteSpace(alignment))
                {
                    string a = alignment.Trim().ToLowerInvariant();
                    int wdAlign = -1;
                    if (a == "left") wdAlign = 0;
                    else if (a == "center") wdAlign = 1;
                    else if (a == "right") wdAlign = 2;
                    else if (a == "justify") wdAlign = 3;
                    if (wdAlign >= 0) try { pf.Alignment = wdAlign; } catch { }
                }
                float f;
                if (!string.IsNullOrWhiteSpace(lineSpacingStr) && float.TryParse(lineSpacingStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { pf.LineSpacing = f * 12f; pf.LineSpacingRule = 0; } catch { } // wdLineSpaceMultiple=5? Use 0 for AtLeast approximated
                if (!string.IsNullOrWhiteSpace(spaceBeforeStr) && float.TryParse(spaceBeforeStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { pf.SpaceBefore = f; } catch { }
                if (!string.IsNullOrWhiteSpace(spaceAfterStr) && float.TryParse(spaceAfterStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { pf.SpaceAfter = f; } catch { }
                if (!string.IsNullOrWhiteSpace(leftIndentStr) && float.TryParse(leftIndentStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { pf.LeftIndent = f; } catch { }
                if (!string.IsNullOrWhiteSpace(firstLineIndentStr) && float.TryParse(firstLineIndentStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { pf.FirstLineIndent = f; } catch { }

                return HostOperationResult.Ok("Paragraph formatting applied.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteSetParagraphFormat failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteSetParagraphFormat");
            }
        }

        public HostOperationResult ExecuteInsertBreak(string breakType, string targetText = null)
        {
            if (string.IsNullOrWhiteSpace(breakType))
                return HostOperationResult.Failed("Break type is required (page, column, section_next_page, section_continuous).");
            string bt = breakType.Trim().ToLowerInvariant();
            int wdBreak = -1;
            if (bt == "page" || bt == "page_break") wdBreak = 7; // wdPageBreak
            else if (bt == "column" || bt == "column_break") wdBreak = 8;
            else if (bt == "section_next_page" || bt == "section_next") wdBreak = 2;
            else if (bt == "section_continuous") wdBreak = 3;
            else return HostOperationResult.Failed("Break type must be page, column, section_next_page, or section_continuous.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    // Find target then collapse to end
                    dynamic doc = app.ActiveDocument;
                    range = doc.Content;
                    dynamic find = range.Find;
                    find.ClearFormatting(); try { find.Replacement.ClearFormatting(); } catch { }
                    find.Text = targetText; find.Forward = true; find.Wrap = 0; find.Format = false;
                    bool ok = find.Execute();
                    if (!ok) return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText));
                    range.Collapse(0); // wdCollapseEnd =0
                    try { range.MoveStart(1, 1); } catch { } // move past found text? Instead collapse end already
                    range.Collapse(0);
                }
                else
                {
                    try { range = app.Selection.Range; } catch { range = app.ActiveDocument.Content; }
                    try { range.Collapse(0); } catch { } // wdCollapseEnd
                }
                if (range == null) return HostOperationResult.Failed("Could not resolve insertion range.");
                range.InsertBreak(wdBreak);
                return HostOperationResult.Ok(string.Format("Inserted {0} break.", bt));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertBreak failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertBreak");
            }
        }

        public HostOperationResult ExecuteSetPageSetup(string orientation, string topMargin, string bottomMargin, string leftMargin, string rightMargin)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic section = null;
                try { section = doc.Sections[1]; } catch { }
                if (section == null) return HostOperationResult.Failed("Could not access document sections.");
                dynamic ps = section.PageSetup;

                if (!string.IsNullOrWhiteSpace(orientation))
                {
                    string o = orientation.Trim().ToLowerInvariant();
                    if (o == "portrait") try { ps.Orientation = 0; } catch { }
                    else if (o == "landscape") try { ps.Orientation = 1; } catch { }
                }
                float f;
                if (!string.IsNullOrWhiteSpace(topMargin) && float.TryParse(topMargin, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { ps.TopMargin = app.InchesToPoints(f); } catch { try { ps.TopMargin = f * 72; } catch { } }
                if (!string.IsNullOrWhiteSpace(bottomMargin) && float.TryParse(bottomMargin, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { ps.BottomMargin = app.InchesToPoints(f); } catch { try { ps.BottomMargin = f * 72; } catch { } }
                if (!string.IsNullOrWhiteSpace(leftMargin) && float.TryParse(leftMargin, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { ps.LeftMargin = app.InchesToPoints(f); } catch { try { ps.LeftMargin = f * 72; } catch { } }
                if (!string.IsNullOrWhiteSpace(rightMargin) && float.TryParse(rightMargin, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out f))
                    try { ps.RightMargin = app.InchesToPoints(f); } catch { try { ps.RightMargin = f * 72; } catch { } }

                return HostOperationResult.Ok("Page setup updated.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteSetPageSetup failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteSetPageSetup");
            }
        }

        public HostOperationResult ExecuteSetHeaderFooter(string headerText, string footerText)
        {
            if (string.IsNullOrWhiteSpace(headerText) && string.IsNullOrWhiteSpace(footerText))
                return HostOperationResult.Failed("Header or footer text must be provided.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic section = null; try { section = doc.Sections[1]; } catch { }
                if (section == null) return HostOperationResult.Failed("Could not access sections.");

                if (!string.IsNullOrWhiteSpace(headerText))
                {
                    try
                    {
                        dynamic header = section.Headers[1]; // wdHeaderFooterPrimary
                        header.Range.Text = headerText;
                    }
                    catch (Exception ex) { Logger.Warn(string.Format("Set header failed: {0}", ex.Message)); return HostOperationResult.Failed("Could not set header: " + ex.Message); }
                }
                if (!string.IsNullOrWhiteSpace(footerText))
                {
                    try
                    {
                        dynamic footer = section.Footers[1];
                        footer.Range.Text = footerText;
                    }
                    catch (Exception ex) { Logger.Warn(string.Format("Set footer failed: {0}", ex.Message)); return HostOperationResult.Failed("Could not set footer: " + ex.Message); }
                }
                return HostOperationResult.Ok("Header/footer updated.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteSetHeaderFooter failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteSetHeaderFooter");
            }
        }

        public HostOperationResult ExecuteInsertPageNumber(string alignment, string headerFooter)
        {
            string align = (alignment ?? "center").Trim().ToLowerInvariant();
            string hf = (headerFooter ?? "footer").Trim().ToLowerInvariant();
            int wdAlign = 1; // center
            if (align == "left") wdAlign = 0;
            else if (align == "right") wdAlign = 2;
            else if (align == "center") wdAlign = 1;
            else if (align == "justify") wdAlign = 3;

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic section = null; try { section = doc.Sections[1]; } catch { }
                if (section == null) return HostOperationResult.Failed("Could not access sections.");

                dynamic targetRange = null;
                if (hf == "header") try { targetRange = section.Headers[1].Range; } catch { }
                else try { targetRange = section.Footers[1].Range; } catch { }
                if (targetRange == null) return HostOperationResult.Failed("Could not access header/footer range.");

                try { targetRange.Collapse(0); } catch { }
                // Alignment
                try { targetRange.ParagraphFormat.Alignment = wdAlign; } catch { }
                targetRange.Fields.Add(targetRange, 33); // wdFieldPage =33
                return HostOperationResult.Ok(string.Format("Inserted page number in {0} ({1}).", hf, align));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertPageNumber failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertPageNumber");
            }
        }

        public HostOperationResult ExecuteInsertHyperlink(string displayText, string address, string targetText = null)
        {
            if (string.IsNullOrWhiteSpace(displayText)) return HostOperationResult.Failed("Display text is required.");
            if (string.IsNullOrWhiteSpace(address)) return HostOperationResult.Failed("Address is required.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    // Find targetText and use its range
                    range = doc.Content;
                    dynamic find = range.Find;
                    find.ClearFormatting(); try { find.Replacement.ClearFormatting(); } catch { }
                    find.Text = targetText; find.Forward = true; find.Wrap = 0; find.Format = false;
                    bool ok = find.Execute();
                    if (!ok) return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText));
                    range.Text = string.Empty; // clear target for replacement
                }
                else
                {
                    try { range = app.Selection.Range; } catch { range = doc.Content; }
                    try { if (app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP) range.Text = string.Empty; } catch { }
                    range.Collapse(0);
                }
                if (range == null) return HostOperationResult.Failed("Could not resolve range.");
                // Insert hyperlink via Hyperlinks.Add
                dynamic hl = doc.Hyperlinks.Add(range, address, Type.Missing, Type.Missing, displayText, Type.Missing);
                if (hl == null) return HostOperationResult.Failed("Could not create hyperlink.");
                return HostOperationResult.Ok(string.Format("Inserted hyperlink '{0}' -> {1}.", displayText, address));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertHyperlink failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertHyperlink");
            }
        }

        public HostOperationResult ExecuteInsertBookmark(string name, string targetText = null)
        {
            if (string.IsNullOrWhiteSpace(name)) return HostOperationResult.Failed("Bookmark name is required.");
            string trimmed = name.Trim();
            if (!System.Text.RegularExpressions.Regex.IsMatch(trimmed, @"^[A-Za-z_][A-Za-z0-9_]*$"))
                return HostOperationResult.Failed("Bookmark name must start with a letter or underscore and contain only letters, digits, underscore.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                if (doc.Bookmarks.Exists(trimmed)) return HostOperationResult.Failed(string.Format("Bookmark '{0}' already exists.", trimmed));

                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(targetText))
                {
                    range = doc.Content;
                    dynamic find = range.Find;
                    find.ClearFormatting(); try { find.Replacement.ClearFormatting(); } catch { }
                    find.Text = targetText; find.Forward = true; find.Wrap = 0; find.Format = false;
                    bool ok = find.Execute();
                    if (!ok) return HostOperationResult.Failed(string.Format("Target text '{0}' not found.", targetText));
                }
                else
                {
                    try { range = app.Selection.Range; } catch { range = doc.Content; }
                }
                if (range == null) return HostOperationResult.Failed("Could not resolve range.");
                doc.Bookmarks.Add(trimmed, range);
                return HostOperationResult.Ok(string.Format("Inserted bookmark '{0}'.", trimmed));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertBookmark failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertBookmark");
            }
        }

        public HostOperationResult ExecuteInsertImage(string imagePath, string widthStr, string heightStr)
        {
            if (string.IsNullOrWhiteSpace(imagePath)) return HostOperationResult.Failed("Image path is required.");
            if (!System.IO.File.Exists(imagePath)) return HostOperationResult.Failed(string.Format("Image file not found: {0}", imagePath));

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic range = null; try { range = app.Selection.Range; } catch { range = app.ActiveDocument.Content; }
                try { range.Collapse(0); } catch { }
                if (range == null) return HostOperationResult.Failed("Could not resolve range.");

                float w = 0, h = 0;
                float.TryParse(widthStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out w);
                float.TryParse(heightStr, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out h);

                dynamic shape = null;
                if (w > 0 && h > 0)
                {
                    try { shape = app.ActiveDocument.InlineShapes.AddPicture(imagePath, false, true, range); }
                    catch { }
                    if (shape != null)
                    {
                        try { if (w > 0) shape.Width = w; } catch { }
                        try { if (h > 0) shape.Height = h; } catch { }
                    }
                }
                else
                {
                    shape = app.ActiveDocument.InlineShapes.AddPicture(imagePath, false, true, range);
                }
                if (shape == null) return HostOperationResult.Failed("Could not insert image.");
                return HostOperationResult.Ok(string.Format("Inserted image {0}.", System.IO.Path.GetFileName(imagePath)));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteInsertImage failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteInsertImage");
            }
        }

        public HostOperationResult ExecuteFormatTable(int tableIndex, string styleName, string bordersStr, string shading)
        {
            if (tableIndex < 1) return HostOperationResult.Failed("Table index must be at least 1.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Tables.Count); } catch { }
                if (tableIndex > count) return HostOperationResult.Failed(string.Format("Table {0} does not exist (document has {1} tables).", tableIndex, count));

                dynamic table = doc.Tables[tableIndex];
                if (!string.IsNullOrWhiteSpace(styleName))
                {
                    try { table.Style = styleName.Trim(); } catch { try { table.set_Style(styleName.Trim()); } catch { } }
                }
                if (!string.IsNullOrWhiteSpace(bordersStr))
                {
                    string b = bordersStr.Trim().ToLowerInvariant();
                    bool enable = b == "true" || b == "1" || b == "yes";
                    try { table.Borders.Enable = enable ? 1 : 0; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(shading))
                {
                    try
                    {
                        int col = HexToWordColor(shading.Trim());
                        // Apply to header row
                        try { table.Rows[1].Shading.BackgroundPatternColor = col; } catch { }
                    }
                    catch { }
                }
                return HostOperationResult.Ok(string.Format("Formatted table {0}.", tableIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteFormatTable failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteFormatTable");
            }
        }

        public HostOperationResult ExecuteListComments()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Comments.Count); } catch { }
                if (count == 0) return HostOperationResult.Ok("No comments in document.");

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(string.Format("{0} comment(s):", count));
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic c = doc.Comments[i];
                        string author = string.Empty; try { author = Convert.ToString(c.Author); } catch { }
                        string date = string.Empty; try { date = Convert.ToString(c.Date); } catch { }
                        string text = string.Empty;
                        try { text = Convert.ToString(c.Range.Text); } catch { }
                        if (string.IsNullOrWhiteSpace(text)) try { text = Convert.ToString(c.Reference); } catch { }
                        string anchor = string.Empty; try { anchor = Convert.ToString(c.Scope.Text); } catch { try { anchor = Convert.ToString(c.Reference); } catch { } }
                        sb.AppendLine(string.Format("{0}. {1} {2}: \"{3}\" {4}", i, author, date, text.Trim(), !string.IsNullOrWhiteSpace(anchor) ? string.Format("[Anchor: {0}]", anchor.Trim()) : string.Empty));
                    }
                    catch (Exception ex) { Logger.Warn(string.Format("List comment {0} failed: {1}", i, ex.Message)); }
                }
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteListComments failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteListComments");
            }
        }

        public HostOperationResult ExecuteDeleteComment(int commentIndex)
        {
            if (commentIndex < 1) return HostOperationResult.Failed("Comment index must be at least 1.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Comments.Count); } catch { }
                if (commentIndex > count) return HostOperationResult.Failed(string.Format("Comment {0} does not exist (has {1}).", commentIndex, count));
                dynamic c = doc.Comments[commentIndex];
                c.Delete();
                return HostOperationResult.Ok(string.Format("Deleted comment {0}.", commentIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteDeleteComment failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteDeleteComment");
            }
        }

        public HostOperationResult ExecuteDeleteCommentByText(string targetText)
        {
            if (string.IsNullOrWhiteSpace(targetText)) return HostOperationResult.Failed("Target text is required.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Comments.Count); } catch { }
                string lower = targetText.Trim().ToLowerInvariant();
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic c = doc.Comments[i];
                        string txt = string.Empty; try { txt = Convert.ToString(c.Range.Text); } catch { }
                        if (!string.IsNullOrWhiteSpace(txt) && txt.ToLowerInvariant().IndexOf(lower, StringComparison.Ordinal) >= 0)
                        {
                            c.Delete();
                            return HostOperationResult.Ok(string.Format("Deleted comment {0} matching '{1}'.", i, targetText));
                        }
                    }
                    catch { }
                }
                return HostOperationResult.Failed(string.Format("No comment containing '{0}' found.", targetText));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteDeleteCommentByText failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteDeleteCommentByText");
            }
        }

        public HostOperationResult ExecuteEditComment(int commentIndex, string newText)
        {
            if (commentIndex < 1) return HostOperationResult.Failed("Comment index must be at least 1.");
            if (string.IsNullOrWhiteSpace(newText)) return HostOperationResult.Failed("New text is required.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Comments.Count); } catch { }
                if (commentIndex > count) return HostOperationResult.Failed(string.Format("Comment {0} does not exist.", commentIndex));
                dynamic c = doc.Comments[commentIndex];
                // Capture anchor range then recreate
                dynamic anchorRange = null; try { anchorRange = c.Scope; } catch { try { anchorRange = c.Range; } catch { } }
                if (anchorRange == null) try { anchorRange = c.Reference; } catch { }
                // Word Comment object has no direct text setter; recreation is safest
                string anchorText = string.Empty; try { anchorText = Convert.ToString(anchorRange.Text); } catch { }
                c.Delete();
                // Re-add at same anchor if possible, else at selection
                dynamic targetRange = anchorRange;
                if (targetRange == null) try { targetRange = app.Selection.Range; } catch { targetRange = doc.Content; }
                doc.Comments.Add(targetRange, newText);
                return HostOperationResult.Ok(string.Format("Edited comment {0}.", commentIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteEditComment failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteEditComment");
            }
        }

        public HostOperationResult ExecuteListRevisions()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Revisions.Count); } catch { }
                if (count == 0) return HostOperationResult.Ok("No tracked revisions.");

                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine(string.Format("{0} revision(s):", count));
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic r = doc.Revisions[i];
                        string author = string.Empty; try { author = Convert.ToString(r.Author); } catch { }
                        string date = string.Empty; try { date = Convert.ToString(r.Date); } catch { }
                        string type = string.Empty; try { type = Convert.ToString(r.Type); } catch { }
                        string text = string.Empty; try { text = Convert.ToString(r.Range.Text); } catch { }
                        if (text != null && text.Length > 80) text = text.Substring(0, 77) + "...";
                        sb.AppendLine(string.Format("{0}. {1} {2} [{3}] \"{4}\"", i, author, date, type, text != null ? text.Trim() : string.Empty));
                    }
                    catch (Exception ex) { Logger.Warn(string.Format("List revision {0} failed: {1}", i, ex.Message)); }
                }
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteListRevisions failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteListRevisions");
            }
        }

        public HostOperationResult ExecuteAcceptRevision(int revisionIndex)
        {
            if (revisionIndex < 1) return HostOperationResult.Failed("Revision index must be at least 1.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Revisions.Count); } catch { }
                if (revisionIndex > count) return HostOperationResult.Failed(string.Format("Revision {0} does not exist (has {1}).", revisionIndex, count));
                dynamic r = doc.Revisions[revisionIndex];
                r.Accept();
                return HostOperationResult.Ok(string.Format("Accepted revision {0}.", revisionIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteAcceptRevision failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteAcceptRevision");
            }
        }

        public HostOperationResult ExecuteRejectRevision(int revisionIndex)
        {
            if (revisionIndex < 1) return HostOperationResult.Failed("Revision index must be at least 1.");
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                int count = 0; try { count = Convert.ToInt32(doc.Revisions.Count); } catch { }
                if (revisionIndex > count) return HostOperationResult.Failed(string.Format("Revision {0} does not exist (has {1}).", revisionIndex, count));
                dynamic r = doc.Revisions[revisionIndex];
                r.Reject();
                return HostOperationResult.Ok(string.Format("Rejected revision {0}.", revisionIndex));
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteRejectRevision failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteRejectRevision");
            }
        }

        public HostOperationResult ExecuteCompareDocuments(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return HostOperationResult.Failed("File path is required.");
            if (!System.IO.File.Exists(filePath)) return HostOperationResult.Failed(string.Format("File not found: {0}", filePath));
            string ext = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
            if (ext != ".docx" && ext != ".doc" && ext != ".pdf" && ext != ".txt")
                return HostOperationResult.Failed("Compare supports .docx/.doc/.pdf/.txt (OpenXML).");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                // Try native Compare - creates a new document with tracked changes
                try
                {
                    dynamic compareDoc = doc.Compare(filePath, "Comparison", 0, true, true, true, true, true, true, true, true, true, true, true, true);
                    if (compareDoc != null)
                    {
                        try { compareDoc.Windows[1].Visible = true; } catch { }
                        return HostOperationResult.Ok(string.Format("Compared with '{0}' - review the new document with tracked changes.", System.IO.Path.GetFileName(filePath)));
                    }
                }
                catch (Exception ex1)
                {
                    Logger.Warn(string.Format("Document.Compare failed, trying Application.CompareDocuments: {0}", ex1.Message));
                    try
                    {
                        dynamic targetDoc = app.Documents.Open(filePath, false, true);
                        if (targetDoc != null)
                        {
                            dynamic result = app.CompareDocuments(doc, targetDoc, 0, 1, true, true, true, true, true, true, true, true, true, true, true);
                            try { targetDoc.Close(false); } catch { }
                            if (result != null) return HostOperationResult.Ok(string.Format("Compared with '{0}'.", System.IO.Path.GetFileName(filePath)));
                        }
                    }
                    catch (Exception ex2)
                    {
                        Logger.Warn(string.Format("CompareDocuments failed: {0}", ex2.Message));
                        return HostOperationResult.FromException(ex2, "WordController.ExecuteCompareDocuments");
                    }
                }
                return HostOperationResult.Failed("Could not create comparison document.");
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteCompareDocuments failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteCompareDocuments");
            }
        }

        public HostOperationResult ExecuteTranslate(string targetLanguage, string sourceLanguage, string paragraphTextOrIndex, string translatedText)
        {
            if (string.IsNullOrWhiteSpace(targetLanguage))
                return HostOperationResult.Failed("Target language is required (e.g. en, fr, de, es, zh).");
            if (string.IsNullOrWhiteSpace(translatedText))
                return HostOperationResult.Failed("Translated text is required. The AI should supply the translation content.");

            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null)
                    return HostOperationResult.Failed("No active Word document available.");

                dynamic doc = app.ActiveDocument;
                // Automatic language detection if sourceLanguage not supplied
                string detectedSource = sourceLanguage;
                if (string.IsNullOrWhiteSpace(detectedSource))
                {
                    try
                    {
                        // Word's LanguageDetected or LanguageID
                        dynamic selRange = null;
                        int paraIdx = 0;
                        if (!string.IsNullOrWhiteSpace(paragraphTextOrIndex) && int.TryParse(paragraphTextOrIndex.Trim(), out paraIdx) && paraIdx >= 1)
                        {
                            try { selRange = doc.Paragraphs[paraIdx].Range; } catch { selRange = app.Selection.Range; }
                        }
                        else if (!string.IsNullOrWhiteSpace(paragraphTextOrIndex))
                        {
                            // find paragraph containing text
                            int pCount = Convert.ToInt32(doc.Paragraphs.Count);
                            string lowerTarget = paragraphTextOrIndex.Trim().ToLowerInvariant();
                            for (int i = 1; i <= pCount; i++)
                            {
                                try
                                {
                                    string pt = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                                    if (pt.ToLowerInvariant().IndexOf(lowerTarget, StringComparison.Ordinal) >= 0) { selRange = doc.Paragraphs[i].Range; break; }
                                }
                                catch { }
                            }
                            if (selRange == null) selRange = app.Selection.Range;
                        }
                        else
                        {
                            try { selRange = app.Selection.Range; } catch { selRange = doc.Content; }
                        }
                        try { detectedSource = Convert.ToString(selRange.LanguageID); } catch { detectedSource = "auto"; }
                        if (string.IsNullOrWhiteSpace(detectedSource)) detectedSource = "auto";
                    }
                    catch { detectedSource = "auto"; }
                }

                // Resolve target range: paragraph index, or text snippet, or selection, or whole doc fallback
                dynamic targetRange = null;
                string resolvedInfo = string.Empty;
                int pIdx = 0;
                if (!string.IsNullOrWhiteSpace(paragraphTextOrIndex) && int.TryParse(paragraphTextOrIndex.Trim(), out pIdx) && pIdx >= 1)
                {
                    int pCount = Convert.ToInt32(doc.Paragraphs.Count);
                    if (pIdx > pCount) return HostOperationResult.Failed(string.Format("Paragraph {0} does not exist (document has {1} paragraphs).", pIdx, pCount));
                    try { targetRange = doc.Paragraphs[pIdx].Range; resolvedInfo = string.Format("paragraph {0}", pIdx); } catch { }
                }
                else if (!string.IsNullOrWhiteSpace(paragraphTextOrIndex))
                {
                    string lowerTarget = paragraphTextOrIndex.Trim().ToLowerInvariant();
                    int pCount = Convert.ToInt32(doc.Paragraphs.Count);
                    for (int i = 1; i <= pCount; i++)
                    {
                        try
                        {
                            string pt = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? string.Empty;
                            if (pt.ToLowerInvariant().IndexOf(lowerTarget, StringComparison.Ordinal) >= 0) { targetRange = doc.Paragraphs[i].Range; resolvedInfo = string.Format("paragraph {0} matching \"{1}\"", i, paragraphTextOrIndex.Trim()); break; }
                        }
                        catch { }
                    }
                    if (targetRange == null)
                    {
                        // Fallback to Find
                        try
                        {
                            dynamic findRange = doc.Content;
                            dynamic find = findRange.Find;
                            find.ClearFormatting(); try { find.Replacement.ClearFormatting(); } catch { }
                            find.Text = paragraphTextOrIndex; find.Forward = true; find.Wrap = 0; find.Format = false;
                            if (find.Execute()) { targetRange = findRange; resolvedInfo = string.Format("range matching \"{0}\"", paragraphTextOrIndex.Trim()); }
                        }
                        catch { }
                    }
                }
                if (targetRange == null)
                {
                    try
                    {
                        dynamic sel = app.Selection;
                        if (sel != null && Convert.ToInt32(sel.Type) != 1 && sel.Range != null) { targetRange = sel.Range; resolvedInfo = "current selection"; }
                        else { targetRange = app.Selection.Range; resolvedInfo = "cursor"; }
                    }
                    catch { targetRange = doc.Content; resolvedInfo = "document"; }
                }
                if (targetRange == null) return HostOperationResult.Failed("Could not resolve target range for translation.");

                string originalText = CleanWordText(Convert.ToString(targetRange.Text) ?? string.Empty);
                if (string.IsNullOrWhiteSpace(originalText)) originalText = "(empty)";

                // Side-by-side preview: add comment with original text anchored to target
                try
                {
                    string commentText = string.Format("[Original ({0}) → {1}] {2}", detectedSource, targetLanguage.Trim(), originalText.Length > 300 ? originalText.Substring(0, 297) + "..." : originalText);
                    doc.Comments.Add(targetRange, commentText);
                }
                catch (Exception cex) { Logger.Warn(string.Format("Translate side-by-side comment failed: {0}", cex.Message)); }

                // Translation-specific Track Changes: enable TrackRevisions and replace text with translated content, preserving paragraph formatting
                var wordApp = GetApp();
                bool wasTrackRevisions = false;
                bool wasTrackFormatting = true;
                bool undoRecordStarted = false;
                try
                {
                    if (wordApp != null && wordApp.ActiveDocument != null)
                    {
                        wasTrackRevisions = wordApp.ActiveDocument.TrackRevisions;
                        wasTrackFormatting = TryGetTrackFormatting(wordApp);
                        undoRecordStarted = TryStartUndoRecord(wordApp, string.Format("Translate {0}→{1}", detectedSource, targetLanguage.Trim()));
                        wordApp.ActiveDocument.TrackRevisions = true;
                        TrySetTrackFormatting(wordApp, false);
                    }

                    // Preserve paragraph style and formatting: save style name
                    string origStyle = string.Empty;
                    try { origStyle = Convert.ToString(targetRange.ParagraphStyle); } catch { try { origStyle = Convert.ToString(targetRange.Style); } catch { } }

                    // Replace text - paragraph Range includes trailing \r, so preserve it
                    string paraMark = string.Empty;
                    string current = Convert.ToString(targetRange.Text) ?? string.Empty;
                    if (current.EndsWith("\r")) paraMark = "\r";
                    // Use translatedText trimmed; Word will handle formatting inheritance
                    string newText = translatedText.Trim();
                    if (!string.IsNullOrEmpty(paraMark) && !newText.EndsWith("\r")) newText = newText + paraMark;

                    // Collapse preserving formatting: setting Range.Text retains paragraph formatting but resets character formatting inside.
                    // We restore style after.
                    targetRange.Text = newText;
                    if (!string.IsNullOrWhiteSpace(origStyle))
                    {
                        try { targetRange.Style = origStyle; } catch { }
                    }

                    return HostOperationResult.Ok(string.Format("Translated {0} ({1}→{2}) with Track Changes and side-by-side comment for review.", resolvedInfo, detectedSource, targetLanguage.Trim()), targetRange.Text);
                }
                finally
                {
                    try
                    {
                        if (wordApp != null && wordApp.ActiveDocument != null)
                        {
                            wordApp.ActiveDocument.TrackRevisions = wasTrackRevisions;
                            TrySetTrackFormatting(wordApp, wasTrackFormatting);
                            EndUndoRecord(wordApp, undoRecordStarted);
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ExecuteTranslate failed", ex);
                return HostOperationResult.FromException(ex, "WordController.ExecuteTranslate");
            }
        }

        public HostOperationResult ExecuteInsertToc(string headingLevels)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                int upper = 1, lower = 3;
                if (!string.IsNullOrWhiteSpace(headingLevels))
                {
                    string hl = headingLevels.Trim();
                    string[] parts = hl.Split(new char[] { '-', ':' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 2) { int.TryParse(parts[0].Trim(), out upper); int.TryParse(parts[1].Trim(), out lower); }
                    else if (parts.Length == 1) { int.TryParse(parts[0].Trim(), out lower); upper = 1; }
                    if (upper < 1) upper = 1;
                    if (lower < upper) lower = upper;
                    if (lower > 9) lower = 9;
                }
                int tocCount = 0;
                try { tocCount = Convert.ToInt32(doc.TablesOfContents.Count); } catch { tocCount = 0; }
                if (tocCount > 0)
                {
                    try { doc.TablesOfContents[1].Update(); return HostOperationResult.Ok("Updated existing Table of Contents."); } catch (Exception ex) { return HostOperationResult.FromException(ex, "WordController.ExecuteInsertToc"); }
                }
                dynamic range = null;
                try { range = app.Selection.Range; } catch { range = doc.Content; }
                try { range.Collapse(0); } catch { }
                doc.TablesOfContents.Add(range, true, upper, lower, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return HostOperationResult.Ok(string.Format("Inserted Table of Contents (levels {0}-{1}).", upper, lower));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteInsertToc failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteInsertToc"); }
        }

        public HostOperationResult ExecuteUpdateToc()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                int tocCount = 0;
                try { tocCount = Convert.ToInt32(doc.TablesOfContents.Count); } catch { tocCount = 0; }
                if (tocCount == 0) return HostOperationResult.Failed("No Table of Contents found to update.");
                int updated = 0;
                for (int i = 1; i <= tocCount; i++) { try { doc.TablesOfContents[i].Update(); updated++; } catch (Exception ex) { Logger.Warn(string.Format("TOC {0} update failed: {1}", i, ex.Message)); } }
                return HostOperationResult.Ok(string.Format("Updated {0} Table(s) of Contents.", updated));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteUpdateToc failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteUpdateToc"); }
        }

        public HostOperationResult ExecuteExportPdf(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return HostOperationResult.Failed("Export path is required.");
            string p = path.Trim().Trim('"').Trim('\'');
            if (!p.ToLowerInvariant().EndsWith(".pdf")) p = p + ".pdf";
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                string dir = System.IO.Path.GetDirectoryName(p);
                if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
                doc.ExportAsFixedFormat(p, 17, false, 0, 0, 0, 0, true, true, 0, true, true, false);
                return HostOperationResult.Ok(string.Format("Exported PDF to {0}.", p));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteExportPdf failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteExportPdf"); }
        }

        public HostOperationResult ExecuteSaveAs(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return HostOperationResult.Failed("Save path is required.");
            string p = path.Trim().Trim('"').Trim('\'');
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                string dir = System.IO.Path.GetDirectoryName(p);
                if (!string.IsNullOrWhiteSpace(dir) && !System.IO.Directory.Exists(dir)) System.IO.Directory.CreateDirectory(dir);
                doc.SaveAs2(p, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                return HostOperationResult.Ok(string.Format("Saved document to {0}.", p));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteSaveAs failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteSaveAs"); }
        }

        public HostOperationResult ExecuteToggleTrackChanges(string enabled)
        {
            if (string.IsNullOrWhiteSpace(enabled)) return HostOperationResult.Failed("Enabled flag required (true/false/on/off).");
            string v = enabled.Trim().ToLowerInvariant();
            bool on = v == "true" || v == "1" || v == "on" || v == "yes" || v == "enable" || v == "enabled";
            bool off = v == "false" || v == "0" || v == "off" || v == "no" || v == "disable" || v == "disabled";
            if (!on && !off) return HostOperationResult.Failed("Enabled must be true/false/on/off.");
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                app.ActiveDocument.TrackRevisions = on;
                return HostOperationResult.Ok(on ? "Track Changes enabled." : "Track Changes disabled.");
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteToggleTrackChanges failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteToggleTrackChanges"); }
        }

        public HostOperationResult ExecuteListStyles()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                int count = 0;
                try { count = Convert.ToInt32(doc.Styles.Count); } catch { count = 0; }
                if (count == 0) return HostOperationResult.Ok("No styles found.");
                var sb = new StringBuilder();
                sb.AppendLine(string.Format("{0} style(s):", count));
                for (int i = 1; i <= Math.Min(count, 100); i++) { try { string n = Convert.ToString(doc.Styles[i].NameLocal); if (!string.IsNullOrWhiteSpace(n)) sb.AppendLine(string.Format("{0}. {1}", i, n.Trim())); } catch { } }
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteListStyles failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteListStyles"); }
        }

        public HostOperationResult ExecuteSetProofingLanguage(string language, string target)
        {
            if (string.IsNullOrWhiteSpace(language)) return HostOperationResult.Failed("Language is required (e.g. en-US, fr-FR).");
            string lang = language.Trim();
            // Map common codes to wdLanguageID constants (e.g. 1033 for en-US) but allow string via LanguageID
            var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) { { "en-US", 1033 }, { "en-GB", 2057 }, { "fr-FR", 1036 }, { "de-DE", 1031 }, { "es-ES", 3082 }, { "it-IT", 1040 }, { "ta-IN", 1097 }, { "hi-IN", 1081 }, { "te-IN", 1098 }, { "kn-IN", 1099 }, { "ml-IN", 1100 }, { "bn-IN", 1093 }, { "mr-IN", 1102 }, { "gu-IN", 1095 } };
            int wdLang = 0;
            if (!map.TryGetValue(lang, out wdLang)) { int.TryParse(lang, out wdLang); }
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(target))
                {
                    string lower = target.Trim().ToLowerInvariant();
                    int pCount = Convert.ToInt32(doc.Paragraphs.Count);
                    for (int i = 1; i <= pCount; i++) { try { string pt = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? ""; if (pt.ToLowerInvariant().IndexOf(lower, StringComparison.Ordinal) >= 0) { range = doc.Paragraphs[i].Range; break; } } catch { } }
                    if (range == null) { try { range = doc.Content; dynamic f = range.Find; f.ClearFormatting(); f.Text = target; if (!f.Execute()) return HostOperationResult.Failed(string.Format("Target '{0}' not found.", target)); } catch { } }
                }
                else { try { dynamic sel = app.Selection; if (sel != null && Convert.ToInt32(sel.Type) != 1 && sel.Range != null) range = sel.Range; else range = doc.Content; } catch { range = doc.Content; } }
                if (range == null) return HostOperationResult.Failed("Could not resolve range for proofing language.");
                if (wdLang != 0) try { range.LanguageID = wdLang; } catch { range.LanguageID = lang; }
                else try { range.LanguageID = lang; } catch (Exception ex) { return HostOperationResult.FromException(ex, "WordController.ExecuteSetProofingLanguage"); }
                return HostOperationResult.Ok(string.Format("Set proofing language to {0}.", lang));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteSetProofingLanguage failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteSetProofingLanguage"); }
        }

        public HostOperationResult ExecuteMergeDocument(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return HostOperationResult.Failed("Path is required.");
            string p = path.Trim().Trim('"').Trim('\'');
            if (!System.IO.File.Exists(p)) return HostOperationResult.Failed(string.Format("File not found: {0}", p));
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                try { range = app.Selection.Range; } catch { range = doc.Content; }
                try { range.Collapse(0); } catch { }
                range.InsertFile(p, Type.Missing, false, false, false);
                return HostOperationResult.Ok(string.Format("Inserted file {0}.", System.IO.Path.GetFileName(p)));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteMergeDocument failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteMergeDocument"); }
        }

        public HostOperationResult ExecuteSetWatermark(string text, string color)
        {
            if (string.IsNullOrWhiteSpace(text)) return HostOperationResult.Failed("Watermark text is required.");
            string t = text.Trim();
            int col = 12632256; // gray
            if (!string.IsNullOrWhiteSpace(color)) { try { col = HexToWordColor(color.Trim()); } catch { } if (col == 0) col = 12632256; }
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                for (int si = 1; si <= Convert.ToInt32(doc.Sections.Count); si++)
                {
                    try
                    {
                        dynamic sec = doc.Sections[si];
                        dynamic hdr = sec.Headers[1];
                        dynamic shape = hdr.Shapes.AddTextEffect(0, t, "Calibri", 36, -1, -1, 150, 300);
                        try { shape.Rotation = 315; } catch { }
                        try { shape.Fill.ForeColor.RGB = col; } catch { }
                        try { shape.Line.Visible = 0; } catch { }
                        try { shape.WrapFormat.Type = 3; } catch { }
                    }
                    catch (Exception ex) { Logger.Warn(string.Format("Watermark section {0} failed: {1}", si, ex.Message)); }
                }
                return HostOperationResult.Ok(string.Format("Watermark '{0}' added.", t));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteSetWatermark failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteSetWatermark"); }
        }

        public HostOperationResult ExecuteInsertCaption(string label, string title)
        {
            if (string.IsNullOrWhiteSpace(title)) return HostOperationResult.Failed("Caption title is required.");
            string lbl = string.IsNullOrWhiteSpace(label) ? "Figure" : label.Trim();
            // Normalize label
            string norm = lbl.ToLowerInvariant();
            if (norm == "table") lbl = "Table";
            else if (norm == "equation") lbl = "Equation";
            else lbl = "Figure";
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                try { range = app.Selection.Range; } catch { range = doc.Content; }
                // InsertCaption requires a label that exists in CaptionLabels collection
                try { doc.CaptionLabels.Add(lbl); } catch { }
                range.InsertCaption(lbl, title.Trim(), Type.Missing, 1, Type.Missing);
                return HostOperationResult.Ok(string.Format("Inserted {0} caption: {1}", lbl, title.Trim()));
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteInsertCaption failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteInsertCaption"); }
        }

        public HostOperationResult ExecuteDelete(string target)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                if (string.IsNullOrWhiteSpace(target))
                {
                    try
                    {
                        dynamic sel = app.Selection;
                        if (sel != null && Convert.ToInt32(sel.Type) != 1 && sel.Range != null) { sel.Range.Delete(); return HostOperationResult.Ok("Deleted selection."); }
                        return HostOperationResult.Failed("Nothing selected to delete.");
                    }
                    catch (Exception ex) { return HostOperationResult.FromException(ex, "WordController.ExecuteDelete"); }
                }
                string t = target.Trim();
                if (t.ToLowerInvariant().StartsWith("table:"))
                {
                    string numStr = t.Substring(6).Trim();
                    int idx;
                    if (!int.TryParse(numStr, out idx) || idx < 1) return HostOperationResult.Failed("Table index must be >=1 (e.g. table:2).");
                    int cnt = Convert.ToInt32(doc.Tables.Count);
                    if (idx > cnt) return HostOperationResult.Failed(string.Format("Table {0} not found (has {1}).", idx, cnt));
                    doc.Tables[idx].Delete();
                    return HostOperationResult.Ok(string.Format("Deleted table {0}.", idx));
                }
                int pIdx;
                if (int.TryParse(t, out pIdx) && pIdx >= 1)
                {
                    int cnt = Convert.ToInt32(doc.Paragraphs.Count);
                    if (pIdx > cnt) return HostOperationResult.Failed(string.Format("Paragraph {0} not found (has {1}).", pIdx, cnt));
                    doc.Paragraphs[pIdx].Range.Delete();
                    return HostOperationResult.Ok(string.Format("Deleted paragraph {0}.", pIdx));
                }
                // Text snippet: find and delete
                try
                {
                    dynamic rng = doc.Content;
                    dynamic f = rng.Find;
                    f.ClearFormatting();
                    f.Text = t;
                    f.Forward = true; f.Wrap = 0; f.Format = false;
                    if (f.Execute()) { rng.Delete(); return HostOperationResult.Ok(string.Format("Deleted text matching '{0}'.", t)); }
                    return HostOperationResult.Failed(string.Format("Text '{0}' not found.", t));
                }
                catch (Exception ex) { return HostOperationResult.FromException(ex, "WordController.ExecuteDelete"); }
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteDelete failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteDelete"); }
        }

        public HostOperationResult ExecuteApplyList(string target, string listType)
        {
            string lt = (listType ?? "bullet").Trim().ToLowerInvariant();
            bool isNumber = lt == "number" || lt == "numbered" || lt == "ordered" || lt == "1" || lt == "decimal";
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                dynamic range = null;
                if (!string.IsNullOrWhiteSpace(target))
                {
                    string lower = target.Trim().ToLowerInvariant();
                    int pCount = Convert.ToInt32(doc.Paragraphs.Count);
                    for (int i = 1; i <= pCount; i++) { try { string pt = Convert.ToString(doc.Paragraphs[i].Range.Text) ?? ""; if (pt.ToLowerInvariant().IndexOf(lower, StringComparison.Ordinal) >= 0) { range = doc.Paragraphs[i].Range; break; } } catch { } }
                    if (range == null) { try { range = doc.Content; dynamic f = range.Find; f.ClearFormatting(); f.Text = target; if (!f.Execute()) return HostOperationResult.Failed(string.Format("Target '{0}' not found.", target)); } catch { } }
                }
                else { try { dynamic sel = app.Selection; if (sel != null && sel.Range != null) range = sel.Range; else range = doc.Content; } catch { range = doc.Content; } }
                if (range == null) return HostOperationResult.Failed("Could not resolve range for list.");
                if (isNumber) range.ListFormat.ApplyNumberDefault();
                else range.ListFormat.ApplyBulletDefault();
                return HostOperationResult.Ok(isNumber ? "Applied numbered list." : "Applied bullet list.");
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteApplyList failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteApplyList"); }
        }

        public HostOperationResult ExecuteReadabilityStats()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.ActiveDocument == null) return HostOperationResult.Failed("No active Word document available.");
                dynamic doc = app.ActiveDocument;
                var sb = new StringBuilder();
                sb.AppendLine(string.Format("Words: {0}", Convert.ToString(doc.Words.Count)));
                sb.AppendLine(string.Format("Characters: {0}", Convert.ToString(doc.Characters.Count)));
                sb.AppendLine(string.Format("Paragraphs: {0}", Convert.ToString(doc.Paragraphs.Count)));
                sb.AppendLine(string.Format("Sentences: {0}", Convert.ToString(doc.Sentences.Count)));
                sb.AppendLine(string.Format("Pages: {0}", Convert.ToString(doc.ComputeStatistics(2)))); // wdStatisticPages =2
                sb.AppendLine(string.Format("Lines: {0}", Convert.ToString(doc.ComputeStatistics(1)))); // wdStatisticLines =1
                try
                {
                    dynamic rss = doc.ReadabilityStatistics;
                    int rc = Convert.ToInt32(rss.Count);
                    for (int i = 1; i <= Math.Min(rc, 10); i++) { try { dynamic rs = rss[i]; string n = Convert.ToString(rs.Name); string v = Convert.ToString(rs.Value); sb.AppendLine(string.Format("{0}: {1}", n, v)); } catch { } }
                }
                catch { }
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex) { Logger.Error("WordController.ExecuteReadabilityStats failed", ex); return HostOperationResult.FromException(ex, "WordController.ExecuteReadabilityStats"); }
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
        /// Returns the active document's file name, or a generic fallback if unavailable.
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

        public string GetActiveDocumentPath()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null)
                {
                    string path = null;
                    try { path = Convert.ToString(app.ActiveDocument.FullName); } catch { }
                    if (!string.IsNullOrWhiteSpace(path)) return path;
                    try { path = Convert.ToString(app.ActiveDocument.Path); } catch { }
                    return path ?? string.Empty;
                }
            }
            catch { }
            return string.Empty;
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

                // Find nearest heading above cursor, efficiently even in long documents.
                // Approach: determine paragraph index containing the selection, then walk backwards
                int paragraphCount = 0;
                try { paragraphCount = app.ActiveDocument.Paragraphs.Count; } catch { paragraphCount = 0; }
                if (paragraphCount <= 0) return "Document: " + GetActiveDocumentName();

                // Locate paragraph index that contains selStart by forward scan until range.End >= selStart
                int selParaIndex = -1;
                // For very large docs, use text-based outline fallback to avoid O(n) COM over thousands of paragraphs
                const int LargeDocThreshold = 800;
                if (paragraphCount > LargeDocThreshold)
                {
                    try
                    {
                        string docText = CleanWordText(app.ActiveDocument.Content.Text);
                        string outline = WordDocumentContextBuilder.BuildDocumentOutline(docText, 2400);
                        if (!string.IsNullOrWhiteSpace(outline))
                        {
                            // Fallback: return last heading line from text-based outline that occurs before selection's text offset ratio
                            string selText = string.Empty;
                            try { selText = CleanWordText(app.Selection.Range.Text); } catch { }
                            // Heuristic: use document outline's first heading as section if we can't map precisely
                            var outlineLines = outline.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                            if (outlineLines.Length > 0)
                            {
                                string firstHeading = outlineLines[0].Trim().TrimStart('-', ' ');
                                if (!string.IsNullOrWhiteSpace(firstHeading))
                                    return string.Format("Section: {0}", firstHeading);
                            }
                        }
                    }
                    catch { }
                }

                // Find selection's paragraph index (1-based)
                try
                {
                    for (int i = 1; i <= paragraphCount; i++)
                    {
                        try
                        {
                            Word.Paragraph para = app.ActiveDocument.Paragraphs[i];
                            if (para == null || para.Range == null) continue;
                            int pStart = Convert.ToInt32(para.Range.Start);
                            int pEnd = Convert.ToInt32(para.Range.End);
                            if (selStart >= pStart && selStart <= pEnd)
                            {
                                selParaIndex = i;
                                break;
                            }
                            if (pStart > selStart)
                            {
                                selParaIndex = i - 1;
                                break;
                            }
                        }
                        catch { }
                    }
                    if (selParaIndex == -1 && paragraphCount > 0) selParaIndex = paragraphCount;
                }
                catch { selParaIndex = paragraphCount; }

                // Walk backwards from selParaIndex to find nearest heading (nearest = most recent before cursor)
                string nearestHeading = null;
                int walkStart = selParaIndex > 0 ? selParaIndex - 1 : paragraphCount;
                // Limit backward walk to 400 paragraphs to avoid COM stall on huge docs, but that's enough for typical sections
                int walkLimit = Math.Max(1, walkStart - 400);
                for (int index = walkStart; index >= walkLimit; index--)
                {
                    try
                    {
                        Word.Paragraph paragraph = app.ActiveDocument.Paragraphs[index];
                        if (paragraph == null || paragraph.Range == null) continue;
                        int outlineLevel = 0;
                        try { outlineLevel = Convert.ToInt32(paragraph.OutlineLevel, CultureInfo.InvariantCulture); } catch { }
                        string styleName = string.Empty;
                        try { styleName = Convert.ToString(paragraph.Style, CultureInfo.InvariantCulture); } catch { }
                        bool isHeading = (outlineLevel >= 1 && outlineLevel <= 9);
                        if (!isHeading && styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                            isHeading = true;
                        if (isHeading)
                        {
                            string text = CleanWordText(paragraph.Range.Text);
                            if (!string.IsNullOrWhiteSpace(text))
                            {
                                nearestHeading = text.Trim();
                                break;
                            }
                        }
                    }
                    catch { }
                }
                // Fallback: if not found in immediate 400, scan forward from start up to selParaIndex (covers early headings in huge gaps)
                if (string.IsNullOrEmpty(nearestHeading) && walkLimit > 1)
                {
                    string lastHeadingFound = null;
                    for (int index = 1; index < selParaIndex; index++)
                    {
                        try
                        {
                            Word.Paragraph paragraph = app.ActiveDocument.Paragraphs[index];
                            if (paragraph == null || paragraph.Range == null) continue;
                            object rangeStartObj = null;
                            try { rangeStartObj = paragraph.Range.Start; } catch { }
                            int rangeStart = rangeStartObj != null ? Convert.ToInt32(rangeStartObj) : -1;
                            if (rangeStart < 0 || rangeStart >= selStart) continue;
                            int outlineLevel = 0;
                            try { outlineLevel = Convert.ToInt32(paragraph.OutlineLevel, CultureInfo.InvariantCulture); } catch { }
                            string styleName = string.Empty;
                            try { styleName = Convert.ToString(paragraph.Style, CultureInfo.InvariantCulture); } catch { }
                            bool isHeading = (outlineLevel >= 1 && outlineLevel <= 9);
                            if (!isHeading && styleName.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0)
                                isHeading = true;
                            if (isHeading)
                            {
                                string text = CleanWordText(paragraph.Range.Text);
                                if (!string.IsNullOrWhiteSpace(text)) lastHeadingFound = text.Trim();
                            }
                        }
                        catch { }
                    }
                    nearestHeading = lastHeadingFound;
                }

                if (!string.IsNullOrEmpty(nearestHeading))
                    return string.Format("Section: {0}", nearestHeading);

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
                if (app != null && app.ActiveDocument != null)
                {
                    // Try rich structure-aware extraction first; fallback to plain Content.Text if COM fails
                    string rich = TryGetLiveRichDocumentText(48000);
                    if (!string.IsNullOrWhiteSpace(rich) && rich.Length > 100)
                        return rich;
                    if (app.ActiveDocument.Content != null)
                        return CleanWordText(app.ActiveDocument.Content.Text);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetFullDocumentText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        private string TryGetLiveRichDocumentText(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 32000;
            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null) return string.Empty;
                dynamic doc = null;
                try { doc = _rawAppObj != null ? ((dynamic)_rawAppObj).ActiveDocument : null; } catch { }
                if (doc == null)
                {
                    try { doc = app.ActiveDocument; } catch { return string.Empty; }
                }
                var sb = new StringBuilder();
                // Paragraphs with semantic structure (style, outline level)
                try
                {
                    int pCount = 0;
                    try { pCount = Convert.ToInt32(doc.Paragraphs.Count); } catch { pCount = 0; }
                    if (pCount > 1200)
                    {
                        // Long document: COM iteration beyond 1200 would be slow and would truncate content after 1200.
                        // Fall back to complete Content.Text enumeration so content after paragraph 1200 still reaches the AI context.
                        // This preserves completeness at the cost of per-paragraph style detail, which is acceptable for long docs.
                        try
                        {
                            string contentText = CleanWordText(Convert.ToString(doc.Content.Text) ?? string.Empty);
                            if (!string.IsNullOrWhiteSpace(contentText))
                            {
                                string[] rawParas = contentText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                                for (int pi = 0; pi < rawParas.Length; pi++)
                                {
                                    string txt = rawParas[pi].Trim();
                                    if (string.IsNullOrWhiteSpace(txt)) continue;
                                    if (txt.Length > 400) txt = txt.Substring(0, 397) + "...";
                                    sb.AppendLine(string.Format("[¶{0}] {1}", pi + 1, txt));
                                    if (sb.Length >= maxCharacters) break;
                                }
                            }
                        }
                        catch { }
                    }
                    else
                    {
                        int cap = pCount;
                        for (int i = 1; i <= cap; i++)
                        {
                            try
                            {
                                dynamic para = doc.Paragraphs[i];
                                if (para == null || para.Range == null) continue;
                                string txt = CleanWordText(Convert.ToString(para.Range.Text) ?? string.Empty);
                                if (string.IsNullOrWhiteSpace(txt)) continue;
                                if (txt.Length > 400) txt = txt.Substring(0, 397) + "...";
                                string style = string.Empty;
                                int outlineLevel = 0;
                                try { outlineLevel = Convert.ToInt32(para.OutlineLevel); } catch { }
                                try { style = Convert.ToString(para.Style); } catch { }
                                bool isHeading = (outlineLevel >= 1 && outlineLevel <= 9) || (!string.IsNullOrWhiteSpace(style) && style.IndexOf("heading", StringComparison.OrdinalIgnoreCase) >= 0);
                                if (isHeading)
                                    sb.AppendLine(string.Format("[¶{0} {1}] {2}", i, !string.IsNullOrWhiteSpace(style) ? style.Trim() : "Heading" + outlineLevel, txt));
                                else
                                    sb.AppendLine(string.Format("[¶{0}] {1}", i, txt));
                                if (sb.Length >= maxCharacters) break;
                            }
                            catch { }
                        }
                    }
                }
                catch { }
                // Tables
                try
                {
                    int tCount = 0;
                    try { tCount = Convert.ToInt32(doc.Tables.Count); } catch { tCount = 0; }
                    for (int ti = 1; ti <= Math.Min(tCount, 30); ti++)
                    {
                        try
                        {
                            dynamic tbl = doc.Tables[ti];
                            if (tbl == null) continue;
                            sb.AppendLine(string.Format("[Table {0}]", ti));
                            int rows = 0, cols = 0;
                            try { rows = Convert.ToInt32(tbl.Rows.Count); } catch { }
                            try { cols = Convert.ToInt32(tbl.Columns.Count); } catch { }
                            for (int r = 1; r <= Math.Min(rows, 20); r++)
                            {
                                var rowCells = new List<string>();
                                for (int c = 1; c <= Math.Min(cols, 12); c++)
                                {
                                    try
                                    {
                                        dynamic cell = tbl.Cell(r, c);
                                        string cellTxt = CleanWordText(Convert.ToString(cell.Range.Text) ?? string.Empty);
                                        cellTxt = cellTxt.Replace("|", "/").Replace("\r", " ").Replace("\n", " ").Trim();
                                        rowCells.Add(cellTxt);
                                    }
                                    catch { rowCells.Add(string.Empty); }
                                }
                                sb.AppendLine("| " + string.Join(" | ", rowCells) + " |");
                                if (sb.Length >= maxCharacters) break;
                            }
                            if (sb.Length >= maxCharacters) break;
                        }
                        catch { }
                    }
                }
                catch { }
                // Headers and Footers per section
                try
                {
                    int secCount = 0;
                    try { secCount = Convert.ToInt32(doc.Sections.Count); } catch { secCount = 0; }
                    for (int si = 1; si <= Math.Min(secCount, 20); si++)
                    {
                        try
                        {
                            dynamic sec = doc.Sections[si];
                            if (sec == null) continue;
                            // Headers
                            try
                            {
                                dynamic headers = sec.Headers;
                                int hCount = Convert.ToInt32(headers.Count);
                                for (int hi = 1; hi <= hCount; hi++)
                                {
                                    try
                                    {
                                        dynamic hdr = headers[hi];
                                        if (hdr == null || hdr.Range == null) continue;
                                        string hTxt = CleanWordText(Convert.ToString(hdr.Range.Text) ?? string.Empty);
                                        if (!string.IsNullOrWhiteSpace(hTxt)) sb.AppendLine(string.Format("[Header Sec{0} {1}] {2}", si, hi, hTxt));
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                            try
                            {
                                dynamic footers = sec.Footers;
                                int fCount = Convert.ToInt32(footers.Count);
                                for (int fi = 1; fi <= fCount; fi++)
                                {
                                    try
                                    {
                                        dynamic ftr = footers[fi];
                                        if (ftr == null || ftr.Range == null) continue;
                                        string fTxt = CleanWordText(Convert.ToString(ftr.Range.Text) ?? string.Empty);
                                        if (!string.IsNullOrWhiteSpace(fTxt)) sb.AppendLine(string.Format("[Footer Sec{0} {1}] {2}", si, fi, fTxt));
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                        }
                        catch { }
                        if (sb.Length >= maxCharacters) break;
                    }
                }
                catch { }
                // Footnotes & Endnotes
                try
                {
                    int fnCount = 0;
                    try { fnCount = Convert.ToInt32(doc.Footnotes.Count); } catch { fnCount = 0; }
                    for (int fi = 1; fi <= Math.Min(fnCount, 50); fi++)
                    {
                        try
                        {
                            dynamic fn = doc.Footnotes[fi];
                            string fnTxt = CleanWordText(Convert.ToString(fn.Range.Text) ?? string.Empty);
                            if (!string.IsNullOrWhiteSpace(fnTxt)) sb.AppendLine(string.Format("[Footnote {0}] {1}", fi, fnTxt));
                        }
                        catch { }
                        if (sb.Length >= maxCharacters) break;
                    }
                }
                catch { }
                try
                {
                    int enCount = 0;
                    try { enCount = Convert.ToInt32(doc.Endnotes.Count); } catch { enCount = 0; }
                    for (int ei = 1; ei <= Math.Min(enCount, 50); ei++)
                    {
                        try
                        {
                            dynamic en = doc.Endnotes[ei];
                            string enTxt = CleanWordText(Convert.ToString(en.Range.Text) ?? string.Empty);
                            if (!string.IsNullOrWhiteSpace(enTxt)) sb.AppendLine(string.Format("[Endnote {0}] {1}", ei, enTxt));
                        }
                        catch { }
                        if (sb.Length >= maxCharacters) break;
                    }
                }
                catch { }
                // Comments
                try
                {
                    int cCount = 0;
                    try { cCount = Convert.ToInt32(doc.Comments.Count); } catch { cCount = 0; }
                    for (int ci = 1; ci <= Math.Min(cCount, 100); ci++)
                    {
                        try
                        {
                            dynamic c = doc.Comments[ci];
                            string author = string.Empty;
                            try { author = Convert.ToString(c.Author) ?? string.Empty; } catch { }
                            string cTxt = CleanWordText(Convert.ToString(c.Range.Text) ?? string.Empty);
                            if (string.IsNullOrWhiteSpace(cTxt)) try { cTxt = CleanWordText(Convert.ToString(c.Reference) ?? string.Empty); } catch { }
                            string anchor = string.Empty;
                            try { anchor = CleanWordText(Convert.ToString(c.Scope.Text) ?? string.Empty); } catch { }
                            if (!string.IsNullOrWhiteSpace(cTxt))
                            {
                                if (!string.IsNullOrWhiteSpace(anchor) && anchor.Length <= 120)
                                    sb.AppendLine(string.Format("[Comment {0} by {1} on \"{2}\"] {3}", ci, author, anchor, cTxt));
                                else
                                    sb.AppendLine(string.Format("[Comment {0} by {1}] {2}", ci, author, cTxt));
                            }
                        }
                        catch { }
                        if (sb.Length >= maxCharacters) break;
                    }
                }
                catch { }
                // Revisions (tracked changes)
                try
                {
                    int rCount = 0;
                    try { rCount = Convert.ToInt32(doc.Revisions.Count); } catch { rCount = 0; }
                    for (int ri = 1; ri <= Math.Min(rCount, 100); ri++)
                    {
                        try
                        {
                            dynamic rev = doc.Revisions[ri];
                            string author = string.Empty;
                            try { author = Convert.ToString(rev.Author) ?? string.Empty; } catch { }
                            string revType = string.Empty;
                            try { revType = Convert.ToString(rev.Type); } catch { }
                            string revTxt = CleanWordText(Convert.ToString(rev.Range.Text) ?? string.Empty);
                            if (revTxt != null && revTxt.Length > 80) revTxt = revTxt.Substring(0, 77) + "...";
                            sb.AppendLine(string.Format("[Revision {0} {1} by {2}] {3}", ri, revType, author, revTxt));
                        }
                        catch { }
                        if (sb.Length >= maxCharacters) break;
                    }
                }
                catch { }
                // Document sections outline (OutlineLevel)
                if (sb.Length > 0 && sb.Length < maxCharacters)
                {
                    // Trim to maxCharacters
                    string result = sb.ToString();
                    if (result.Length > maxCharacters) result = result.Substring(0, maxCharacters) + "\n...[truncated]";
                    return result;
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.TryGetLiveRichDocumentText failed: {0}", ex.Message));
                return string.Empty;
            }
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

        private static bool TryGetTrackFormatting(Word.Application app)
        {
            try
            {
                if (app != null && app.ActiveDocument != null) return app.ActiveDocument.TrackFormatting;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.TryGetTrackFormatting failed: {0}", ex.Message));
            }
            return true;
        }

        private static void TrySetTrackFormatting(Word.Application app, bool value)
        {
            try
            {
                if (app != null && app.ActiveDocument != null) app.ActiveDocument.TrackFormatting = value;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.TrySetTrackFormatting failed: {0}", ex.Message));
            }
        }

        private bool RenderWithTrackChanges(string markdown, bool replaceSelection)
        {
            if (markdown == null) markdown = string.Empty;

            try
            {
                var app = GetApp();
                if (app == null || app.ActiveDocument == null || app.Selection == null) return false;

                bool wasTrackRevisions = app.ActiveDocument.TrackRevisions;
                // Formatting revisions are suppressed while the renderer runs: it sets font,
                // size and style on every emitted range, and Word would log each of those as a
                // separate "Formatted" revision, burying the single insertion the user cares about.
                bool wasTrackFormatting = TryGetTrackFormatting(app);
                bool undoRecordStarted = TryStartUndoRecord(app, "AI Assistant tracked edit");
                try
                {
                    app.ActiveDocument.TrackRevisions = true;
                    TrySetTrackFormatting(app, false);
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
                    TrySetTrackFormatting(app, wasTrackFormatting);
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
