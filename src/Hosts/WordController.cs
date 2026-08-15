using System;
using System.Runtime.InteropServices;
using System.Text;
using MistralOfficeAddin.Core;
using Word = NetOffice.WordApi;

namespace MistralOfficeAddin.Hosts
{
    public class WordController
    {
        // Store as raw object; wrap lazily only when needed
        private readonly object _rawAppObj;
        private Word.Application _wordApp;

        public WordController(object appObj)
        {
            // Store raw reference — do NOT wrap with NetOffice here.
            // NetOffice wrapping triggers COM event subscriptions which block on the COM thread.
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
                var app = GetApp();
                if (app != null && app.Selection != null)
                {
                    // If selection is just the blinking cursor (insertion point), return empty to read full document
                    if (app.Selection.Type == Word.Enums.WdSelectionType.wdSelectionIP)
                    {
                        return string.Empty;
                    }

                    string text = app.Selection.Text;
                    if (!string.IsNullOrEmpty(text))
                    {
                        text = text.TrimEnd('\r', '\n', '\a', ' ');
                        // Ignore 1-character selections (usually accidental cursor clicks)
                        if (text.Trim().Length > 1)
                        {
                            return text;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetSelectedText failed: {0}", ex.Message));
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
                    WordMarkdownRenderer.Render(app, text);
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.InsertTextAtCursor failed", ex);
                throw;
            }
        }

        public void ReplaceSelection(string text)
        {
            if (text == null) text = string.Empty;
            try
            {
                var app = GetApp();
                if (app != null)
                {
                    if (app.Selection != null && app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                    {
                        app.Selection.Text = string.Empty;
                    }
                    WordMarkdownRenderer.Render(app, text);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ReplaceSelection failed", ex);
                throw;
            }
        }

        public string GetDocumentText(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 32000;
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null && app.ActiveDocument.Content != null)
                {
                    string fullText = app.ActiveDocument.Content.Text ?? string.Empty;
                    fullText = fullText.TrimEnd('\r', '\n', '\a', ' ');
                    if (fullText.Length > maxCharacters)
                        return fullText.Substring(0, maxCharacters) + "\n...[document truncated for length]";
                    return fullText;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WordController.GetDocumentText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

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

        public void ApplyTrackChanges(string suggestedText)
        {
            if (string.IsNullOrEmpty(suggestedText)) return;
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveDocument != null)
                {
                    bool wasTrackRevisions = app.ActiveDocument.TrackRevisions;
                    try
                    {
                        app.ActiveDocument.TrackRevisions = true;
                        if (app.Selection != null)
                        {
                            if (!string.IsNullOrEmpty(app.Selection.Text) &&
                                app.Selection.Type != Word.Enums.WdSelectionType.wdSelectionIP)
                                app.Selection.Text = suggestedText;
                            else
                                app.Selection.TypeText(suggestedText);
                        }
                    }
                    finally
                    {
                        app.ActiveDocument.TrackRevisions = wasTrackRevisions;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("WordController.ApplyTrackChanges failed", ex);
                throw;
            }
        }
    }
}
