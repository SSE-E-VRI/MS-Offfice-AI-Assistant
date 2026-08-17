using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.Hosts
{
    public class PowerPointController
    {
        private readonly object _rawAppObj;

        public PowerPointController(object appObj)
        {
            _rawAppObj = appObj;
        }

        private class SlideData
        {
            public string Title { get; set; }
            public List<string> Bullets { get; set; }
            public string SpeakerNotes { get; set; }

            public SlideData()
            {
                Bullets = new List<string>();
            }
        }

        private dynamic GetOrCreateActiveSlide(bool createIfNone = false)
        {
            if (_rawAppObj == null) return null;
            try
            {
                dynamic app = _rawAppObj;
                dynamic activeWin = app.ActiveWindow;
                if (activeWin != null)
                {
                    // 1. Try View.Slide
                    try
                    {
                        dynamic slide = activeWin.View.Slide;
                        if (slide != null) return slide;
                    }
                    catch { }

                    // 2. Try Selection.SlideRange
                    try
                    {
                        dynamic selection = activeWin.Selection;
                        if (selection != null && selection.SlideRange != null && selection.SlideRange.Count > 0)
                        {
                            return selection.SlideRange[1];
                        }
                    }
                    catch { }
                }

                // 3. Fallback: Check ActivePresentation
                dynamic pres = null;
                try { pres = app.ActivePresentation; } catch { }
                if (pres == null && createIfNone)
                {
                    try { pres = app.Presentations.Add(1); } catch { }
                }

                if (pres != null)
                {
                    dynamic slides = pres.Slides;
                    if (slides != null)
                    {
                        int count = 0;
                        try { count = Convert.ToInt32(slides.Count); } catch { }
                        if (count > 0)
                        {
                            return slides[1];
                        }
                        else if (createIfNone)
                        {
                            // Empty presentation (0 slides): create Slide 1
                            try
                            {
                                return slides.Add(1, 2); // ppLayoutText = 2
                            }
                            catch
                            {
                                try { return slides.Add(1, 1); } catch { }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetOrCreateActiveSlide failed: {0}", ex.Message));
            }
            return null;
        }

        public string GetSlideText()
        {
            try
            {
                dynamic slide = GetOrCreateActiveSlide(false);
                if (slide != null)
                {
                    var sb = new StringBuilder();
                    try
                    {
                        sb.AppendLine(string.Format("[Slide #{0}: {1}]", slide.SlideNumber, slide.Name));
                    }
                    catch { }

                    dynamic shapes = slide.Shapes;
                    if (shapes != null)
                    {
                        int count = Convert.ToInt32(shapes.Count);
                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                dynamic shape = shapes[i];
                                if (shape != null && Convert.ToInt32(shape.HasTextFrame) != 0)
                                {
                                    dynamic textFrame = shape.TextFrame;
                                    if (textFrame != null && Convert.ToInt32(textFrame.HasText) != 0)
                                    {
                                        string text = Convert.ToString(textFrame.TextRange.Text);
                                        if (!string.IsNullOrWhiteSpace(text))
                                            sb.AppendLine(text.Trim());
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    return sb.ToString().TrimEnd();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetSlideText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public void InsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return;

            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return;

                // 1. If user has actively selected text or a specific shape, insert directly
                try
                {
                    dynamic selection = app.ActiveWindow.Selection;
                    if (selection != null)
                    {
                        int selType = Convert.ToInt32(selection.Type);
                        // ppSelectionText = 3
                        if (selType == 3)
                        {
                            selection.TextRange.InsertAfter(CleanMarkdown(text));
                            return;
                        }
                    }
                }
                catch { }

                // 2. Parse text into structured slides
                var slides = ParseSlideData(text);
                if (slides.Count == 0)
                {
                    // Fallback to simple bullet insertion
                    AddBulletPoints(new List<string>(text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)));
                    return;
                }

                dynamic pres = null;
                try { pres = app.ActivePresentation; } catch { }
                if (pres == null)
                {
                    try { pres = app.Presentations.Add(1); } catch { }
                }

                dynamic activeSlide = GetOrCreateActiveSlide(true);
                bool isFirst = true;

                foreach (var slideData in slides)
                {
                    dynamic targetSlide;
                    if (isFirst && activeSlide != null)
                    {
                        targetSlide = activeSlide;
                        isFirst = false;
                    }
                    else
                    {
                        // Add subsequent slides at the end
                        int slideCount = Convert.ToInt32(pres.Slides.Count);
                        targetSlide = pres.Slides.Add(slideCount + 1, 2); // ppLayoutText = 2
                    }

                    PopulateSlide(targetSlide, slideData);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.InsertText failed", ex);
                throw;
            }
        }

        private void PopulateSlide(dynamic slide, SlideData data)
        {
            if (slide == null) return;

            dynamic shapes = slide.Shapes;
            dynamic titleShape = null;
            dynamic bodyShape = null;

            try
            {
                int count = Convert.ToInt32(shapes.Count);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        int type = Convert.ToInt32(shape.Type);
                        // msoPlaceholder = 14
                        if (type == 14 && shape.PlaceholderFormat != null)
                        {
                            int phType = Convert.ToInt32(shape.PlaceholderFormat.Type);
                            // ppPlaceholderTitle = 1, ppPlaceholderCenterTitle = 3
                            if ((phType == 1 || phType == 3) && titleShape == null)
                            {
                                titleShape = shape;
                            }
                            // ppPlaceholderBody = 2, ppPlaceholderObject = 7
                            else if ((phType == 2 || phType == 7) && bodyShape == null)
                            {
                                bodyShape = shape;
                            }
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // Set Title
            if (!string.IsNullOrWhiteSpace(data.Title))
            {
                string cleanTitle = CleanMarkdown(data.Title);
                if (titleShape != null && Convert.ToInt32(titleShape.HasTextFrame) != 0)
                {
                    titleShape.TextFrame.TextRange.Text = cleanTitle;
                }
                else
                {
                    try
                    {
                        if (Convert.ToInt32(shapes.HasTitle) != 0)
                        {
                            shapes.Title.TextFrame.TextRange.Text = cleanTitle;
                        }
                    }
                    catch { }
                }
            }

            // Set Body / Bullets
            if (data.Bullets != null && data.Bullets.Count > 0)
            {
                if (bodyShape != null && Convert.ToInt32(bodyShape.HasTextFrame) != 0)
                {
                    dynamic textRange = bodyShape.TextFrame.TextRange;
                    textRange.Text = string.Empty; // clear default placeholder text

                    for (int b = 0; b < data.Bullets.Count; b++)
                    {
                        string bullet = CleanMarkdown(data.Bullets[b]);
                        if (string.IsNullOrWhiteSpace(bullet)) continue;

                        try
                        {
                            dynamic para = textRange.InsertAfter(b == 0 ? bullet : ("\r" + bullet));
                            // ppBulletUnnumbered = 1
                            para.ParagraphFormat.Bullet.Type = 1;
                        }
                        catch
                        {
                            textRange.InsertAfter((b == 0 ? "• " : "\r• ") + bullet);
                        }
                    }
                }
                else
                {
                    // If no body placeholder exists, add clean textbox
                    try
                    {
                        dynamic tb = shapes.AddTextbox(1, 80, 140, 560, 320);
                        dynamic textRange = tb.TextFrame.TextRange;
                        for (int b = 0; b < data.Bullets.Count; b++)
                        {
                            string bullet = CleanMarkdown(data.Bullets[b]);
                            if (string.IsNullOrWhiteSpace(bullet)) continue;
                            textRange.InsertAfter((b == 0 ? "• " : "\r• ") + bullet);
                        }
                    }
                    catch { }
                }
            }

            // Set Speaker Notes if present
            if (!string.IsNullOrWhiteSpace(data.SpeakerNotes))
            {
                try
                {
                    if (slide.NotesPage != null && slide.NotesPage.Shapes != null)
                    {
                        dynamic noteShapes = slide.NotesPage.Shapes;
                        int nCount = Convert.ToInt32(noteShapes.Count);
                        for (int n = 1; n <= nCount; n++)
                        {
                            dynamic nShape = noteShapes[n];
                            if (nShape != null && nShape.PlaceholderFormat != null)
                            {
                                if (Convert.ToInt32(nShape.PlaceholderFormat.Type) == 2)
                                {
                                    nShape.TextFrame.TextRange.Text = CleanMarkdown(data.SpeakerNotes);
                                    break;
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private List<SlideData> ParseSlideData(string rawText)
        {
            var result = new List<SlideData>();
            if (string.IsNullOrWhiteSpace(rawText)) return result;

            // Split by slide boundaries: "Slide 1:", "### Slide 2:", "## Slide 3", "---", etc.
            string[] rawBlocks = Regex.Split(rawText, @"(?m)(?:^|\n)(?:---|\*\*\*|___|(?:#{1,4}\s*)?(?:\*\*)?Slide\s+\d+[:.]?.*?\n)", RegexOptions.IgnoreCase);

            var blocks = new List<string>();
            foreach (var b in rawBlocks)
            {
                if (!string.IsNullOrWhiteSpace(b)) blocks.Add(b.Trim());
            }

            if (blocks.Count == 0)
            {
                blocks.Add(rawText.Trim());
            }

            foreach (var block in blocks)
            {
                var slide = new SlideData();
                var lines = block.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                bool inNotes = false;
                var notesSb = new StringBuilder();

                foreach (var rawLine in lines)
                {
                    string line = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    // Detect notes section
                    if (line.StartsWith("Speaker Notes:", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("**Speaker Notes:**", StringComparison.OrdinalIgnoreCase) ||
                        line.StartsWith("Notes:", StringComparison.OrdinalIgnoreCase))
                    {
                        inNotes = true;
                        string noteContent = Regex.Replace(line, @"(?i)^\*?\*?(?:Speaker\s+)?Notes:\*?\*?\s*", "");
                        if (!string.IsNullOrWhiteSpace(noteContent)) notesSb.AppendLine(noteContent);
                        continue;
                    }

                    if (inNotes)
                    {
                        notesSb.AppendLine(line);
                        continue;
                    }

                    // Detect title
                    if (string.IsNullOrEmpty(slide.Title) &&
                        (line.StartsWith("#") ||
                         line.StartsWith("Title:", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("**Title:**", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("**Slide", StringComparison.OrdinalIgnoreCase) ||
                         line.StartsWith("Slide", StringComparison.OrdinalIgnoreCase)))
                    {
                        string titleText = Regex.Replace(line, @"^(?:#{1,4}\s*|\*?\*?Title:\*?\*?\s*|\*?\*?Slide\s*\d*[:.]?\s*)", "", RegexOptions.IgnoreCase);
                        slide.Title = CleanMarkdown(titleText);
                        continue;
                    }

                    // Skip section headings like "**Content:**" or "### Bullet Points"
                    if (Regex.IsMatch(line, @"^\*?\*?Content(?:\s*\(.*?\))?:\*?\*?$", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(line, @"^\*?\*?Bullet Points:\*?\*?$", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(line, @"^\*?\*?Design Tip:.*?\*?\*?$", RegexOptions.IgnoreCase) ||
                        Regex.IsMatch(line, @"^\*?\*?Visual:.*?\*?\*?$", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    // Bullets / content line
                    string bulletText = Regex.Replace(line, @"^[-*•+]\s*|^\d+[\.)]\s*", "");
                    slide.Bullets.Add(CleanMarkdown(bulletText));
                }

                if (notesSb.Length > 0)
                {
                    slide.SpeakerNotes = notesSb.ToString().Trim();
                }

                if (!string.IsNullOrWhiteSpace(slide.Title) || slide.Bullets.Count > 0)
                {
                    result.Add(slide);
                }
            }

            return result;
        }

        private static string CleanMarkdown(string input)
        {
            if (string.IsNullOrEmpty(input)) return string.Empty;

            string s = input;
            // Remove markdown bold/italic (**text**, *text*, __text__, _text_)
            s = Regex.Replace(s, @"\*\*([^*]+)\*\*", "$1");
            s = Regex.Replace(s, @"\*([^*]+)\*", "$1");
            s = Regex.Replace(s, @"__([^_]+)__", "$1");
            s = Regex.Replace(s, @"_([^_]+)_", "$1");
            // Remove inline code ticks (`code`)
            s = Regex.Replace(s, @"`([^`]+)`", "$1");
            // Remove markdown headers (### Header)
            s = Regex.Replace(s, @"^#{1,6}\s*", "");
            // Remove leading bullet characters
            s = Regex.Replace(s, @"^[-*•+]\s*", "");
            // Remove stray markdown artifacts
            s = s.Replace("**", "").Replace("##", "").Replace("###", "");

            return s.Trim();
        }

        public void AddBulletPoints(List<string> bullets)
        {
            if (bullets == null || bullets.Count == 0) return;

            try
            {
                var slideData = new SlideData();
                foreach (var b in bullets)
                {
                    string clean = CleanMarkdown(b);
                    if (!string.IsNullOrWhiteSpace(clean))
                        slideData.Bullets.Add(clean);
                }

                dynamic slide = GetOrCreateActiveSlide(true);
                PopulateSlide(slide, slideData);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.AddBulletPoints failed", ex);
                throw;
            }
        }

        public void SetSpeakerNotes(string notes)
        {
            if (notes == null) notes = string.Empty;
            try
            {
                dynamic slide = GetOrCreateActiveSlide(true);
                if (slide != null && slide.NotesPage != null && slide.NotesPage.Shapes != null)
                {
                    dynamic shapes = slide.NotesPage.Shapes;
                    int count = Convert.ToInt32(shapes.Count);
                    for (int i = 1; i <= count; i++)
                    {
                        dynamic shape = shapes[i];
                        if (shape != null && shape.PlaceholderFormat != null)
                        {
                            // ppPlaceholderBody = 2
                            if (Convert.ToInt32(shape.PlaceholderFormat.Type) == 2)
                            {
                                shape.TextFrame.TextRange.Text = CleanMarkdown(notes);
                                return;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.SetSpeakerNotes failed", ex);
                throw;
            }
        }

        public string GetActivePresentationName()
        {
            try
            {
                if (_rawAppObj != null)
                {
                    dynamic app = _rawAppObj;
                    dynamic pres = app.ActivePresentation;
                    if (pres != null)
                        return Convert.ToString(pres.Name);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetActivePresentationName failed: {0}", ex.Message));
            }
            return "Presentation";
        }
    }
}
