using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.Hosts
{
    public class PowerPointController : IOfficeHostController
    {
        private readonly object _rawAppObj;

        public string HostType
        {
            get { return "PowerPoint"; }
        }

        public PowerPointController(object appObj)
        {
            _rawAppObj = appObj;
        }

        public string GetActiveDocumentName()
        {
            return GetActivePresentationName();
        }

        public string GetActivePresentationPath()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null && app.ActivePresentation != null)
                {
                    string path = null;
                    try { path = Convert.ToString(app.ActivePresentation.FullName); } catch { }
                    if (!string.IsNullOrWhiteSpace(path)) return path;
                    try { path = Convert.ToString(app.ActivePresentation.Path); } catch { }
                    return path ?? string.Empty;
                }
            }
            catch { }
            return string.Empty;
        }

        public string GetSelectedText()
        {
            try
            {
                dynamic app = _rawAppObj;
                dynamic activeWin = null;
                try { activeWin = app != null ? app.ActiveWindow : null; } catch { }
                if (activeWin != null && IsNormalOrSlideView(activeWin))
                {
                    dynamic selection = null;
                    try { selection = activeWin.Selection; } catch { }
                    if (selection != null)
                    {
                        int selType = 0;
                        try { selType = Convert.ToInt32(selection.Type); } catch { }
                        // ppSelectionText = 3
                        if (selType == 3 && selection.TextRange != null)
                        {
                            string txt = Convert.ToString(selection.TextRange.Text);
                            if (!string.IsNullOrWhiteSpace(txt)) return txt;
                        }
                    }
                }
            }
            catch { }
            return GetSlideText();
        }

        public string GetDocumentContext(string prompt, int maxCharacters)
        {
            return GetPresentationReviewContext(maxCharacters > 0 ? maxCharacters : 7000);
        }

        public bool Undo()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app != null)
                {
                    app.CommandBars.ExecuteMso("Undo");
                    return true;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.Undo failed: {0}", ex.Message));
            }
            return false;
        }

        private static bool IsNormalOrSlideView(dynamic activeWin)
        {
            if (activeWin == null) return false;
            try
            {
                int viewType = Convert.ToInt32(activeWin.ViewType);
                // 1 = ppViewSlide, 9 = ppViewNormal
                return viewType == 1 || viewType == 9;
            }
            catch
            {
                return false;
            }
        }

        // D-2: GetActivePresentation(createIfNone:true) and GetOrCreateActiveSlide(createIfNone:true)
        // were "Get"-shaped names that silently created a presentation/slide as a side effect when
        // the boolean flag was true — a caller (or a future risk-classification pass) could easily
        // read past the flag and treat the call as read-only. Fixed by removing the boolean entirely:
        // *Core holds the shared implementation, and every caller now uses one of two thin, honestly
        // named wrappers below so the method name alone states whether it can mutate.
        private dynamic GetActivePresentationCore(bool createIfNone)
        {
            if (_rawAppObj == null) return null;
            try
            {
                dynamic app = _rawAppObj;
                dynamic presentation = null;
                try { presentation = app.ActivePresentation; } catch { }
                if (presentation == null && createIfNone)
                {
                    // msoTrue is -1.  Passing the explicit value avoids relying on
                    // the COM coercion of msoCTrue (1) when a presentation window is created.
                    try { presentation = app.Presentations.Add(-1); }
                    catch (Exception addEx) { Logger.Warn(string.Format("PowerPointController could not create presentation: {0}", addEx.Message)); }
                }
                return presentation;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetActivePresentation failed: {0}", ex.Message));
                return null;
            }
        }

        /// <summary>Read-only. Returns null if no presentation is open — never creates one.</summary>
        private dynamic GetActivePresentation()
        {
            return GetActivePresentationCore(false);
        }

        /// <summary>Mutating. Creates a new presentation if none is currently open.</summary>
        private dynamic GetOrCreateActivePresentation()
        {
            return GetActivePresentationCore(true);
        }

        private dynamic GetActiveSlideCore(bool createIfNone)
        {
            if (_rawAppObj == null) return null;
            try
            {
                dynamic app = _rawAppObj;
                dynamic activeWin = null;
                // ActiveWindow can throw while PowerPoint is between presentation windows or in Protected View.
                try { activeWin = app.ActiveWindow; } catch { }
                if (activeWin != null)
                {
                    // 1. Try View.Slide if in normal or single slide view
                    if (IsNormalOrSlideView(activeWin))
                    {
                        try
                        {
                            dynamic slide = activeWin.View.Slide;
                            if (slide != null) return slide;
                        }
                        catch { }
                    }

                    // 2. Try Selection.SlideRange (works in Slide Sorter view or Normal view)
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
                dynamic pres = GetActivePresentationCore(createIfNone);

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
                            catch (Exception layoutEx)
                            {
                                Logger.Warn(string.Format("PowerPointController could not add slide with ppLayoutText: {0}", layoutEx.Message));
                                try { return slides.Add(1, 1); } // ppLayoutBlank = 1
                                catch (Exception blankEx) { Logger.Warn(string.Format("PowerPointController could not create slide: {0}", blankEx.Message)); }
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

        /// <summary>Read-only. Returns null if there is no resolvable active slide — never creates one.</summary>
        private dynamic GetActiveSlide()
        {
            return GetActiveSlideCore(false);
        }

        /// <summary>Mutating. Creates a new slide (and/or presentation) if none currently exists.</summary>
        private dynamic GetOrCreateActiveSlide()
        {
            return GetActiveSlideCore(true);
        }

        public string GetSlideText()
        {
            try
            {
                dynamic slide = GetActiveSlide();
                if (slide != null)
                {
                    return GetSlideTextInternal(slide, true).TrimEnd();
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetSlideText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        /// <summary>
        /// Returns an all-slide, bounded context including speaker notes. This is used for
        /// deck-level Q&amp;A and review prompts instead of limiting the model to the active slide.
        /// </summary>
        public string GetPresentationText(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 48000;
            try
            {
                dynamic presentation = GetActivePresentation();
                if (presentation == null || presentation.Slides == null) return string.Empty;

                var sb = new StringBuilder();
                int count = Convert.ToInt32(presentation.Slides.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = null;
                    try { slide = presentation.Slides[i]; } catch { }
                    if (slide == null) continue;

                    string sectionName = GetSectionName(presentation, i);
                    if (!string.IsNullOrWhiteSpace(sectionName))
                        AppendBounded(sb, string.Format("[Section: {0}]\n", sectionName), maxCharacters);

                    AppendBounded(sb, GetSlideTextInternal(slide, true), maxCharacters);
                    AppendBounded(sb, "\n\n", maxCharacters);
                    if (sb.Length >= maxCharacters) break;
                }

                if (sb.Length >= maxCharacters)
                    return sb.ToString(0, maxCharacters) + "\n...[presentation truncated for length]";
                return sb.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetPresentationText failed: {0}", ex.Message));
                return string.Empty;
            }
        }

        /// <summary>
        /// Supplies a compact, deterministic review brief before the LLM adds recommendations.
        /// </summary>
        public string GetPresentationReviewContext(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 28000;
            try
            {
                dynamic presentation = GetActivePresentation();
                if (presentation == null || presentation.Slides == null) return string.Empty;

                int slideCount = Convert.ToInt32(presentation.Slides.Count);
                int untitled = 0;
                int noBody = 0;
                var titles = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var sb = new StringBuilder();
                sb.AppendLine(string.Format("[Presentation Review Context | Slides: {0}]", slideCount));

                for (int i = 1; i <= slideCount; i++)
                {
                    dynamic slide = presentation.Slides[i];
                    string title = GetSlideTitle(slide);
                    // Speaker notes are useful context, but they must not make a slide with no
                    // visible body appear complete in the deterministic review summary.
                    string slideText = GetSlideTextInternal(slide, false);
                    if (string.IsNullOrWhiteSpace(title)) untitled++;
                    else
                    {
                        int occurrences;
                        titles.TryGetValue(title.Trim(), out occurrences);
                        titles[title.Trim()] = occurrences + 1;
                    }
                    if (CountBodyLines(slideText) == 0) noBody++;

                    AppendBounded(sb, string.Format("Slide {0}: {1}\n", i, string.IsNullOrWhiteSpace(title) ? "(untitled)" : title.Trim()), maxCharacters);
                    if (sb.Length >= maxCharacters) break;
                }

                var duplicateTitles = new List<string>();
                foreach (var pair in titles)
                {
                    if (pair.Value > 1) duplicateTitles.Add(pair.Key);
                }
                sb.Insert(0, string.Format("Untitled slides: {0}; slides without body text: {1}; duplicate titles: {2}.\n",
                    untitled, noBody, duplicateTitles.Count == 0 ? "none" : string.Join(", ", duplicateTitles.ToArray())));
                return sb.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetPresentationReviewContext failed: {0}", ex.Message));
                return string.Empty;
            }
        }

        private static bool SlideHasSubstantiveContent(dynamic slide)
        {
            if (slide == null) return false;
            try
            {
                dynamic shapes = slide.Shapes;
                int count = shapes != null ? Convert.ToInt32(shapes.Count) : 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        if (shape == null || Convert.ToInt32(shape.HasTextFrame) == 0) continue;
                        dynamic textFrame = shape.TextFrame;
                        if (textFrame == null || Convert.ToInt32(textFrame.HasText) == 0) continue;
                        string text = Convert.ToString(textFrame.TextRange.Text);
                        if (string.IsNullOrWhiteSpace(text)) continue;

                        string trimmed = text.Trim();
                        // Ignore default PowerPoint placeholder prompts
                        if (trimmed.Equals("Click to add title", StringComparison.OrdinalIgnoreCase) ||
                            trimmed.Equals("Click to add text", StringComparison.OrdinalIgnoreCase) ||
                            trimmed.Equals("Click to add subtitle", StringComparison.OrdinalIgnoreCase) ||
                            trimmed.Equals("Click to add content", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        return true;
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        public bool InsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;

            try
            {
                string cleanedText;
                var legacyActions = PowerPointActionParser.ParseStructuredActions(text, out cleanedText);
                bool legacyExecuted = false;
                bool legacyAnySucceeded = false;
                if (legacyActions != null && legacyActions.Count > 0)
                {
                    foreach (var act in legacyActions)
                    {
                        try
                        {
                            bool ok = ExecuteLegacyPowerPointAction(act);
                            legacyExecuted = true;
                            if (ok) legacyAnySucceeded = true;
                            act.Status = ok ? PowerPointActionStatus.Applied : PowerPointActionStatus.Error;
                            if (ok) act.ResultText = "Applied via legacy InsertText";
                            else act.ErrorMessage = "Failed via legacy InsertText";
                        }
                        catch (Exception lex)
                        {
                            Logger.Warn(string.Format("PowerPointController.InsertText legacy action {0} failed: {1}", act.Type, lex.Message));
                            act.Status = PowerPointActionStatus.Error;
                            act.ErrorMessage = lex.Message;
                        }
                    }
                }
                text = cleanedText;
                if (string.IsNullOrWhiteSpace(text))
                {
                    // If we executed legacy actions, return success based on them (prevents silent discard)
                    if (legacyExecuted) return legacyAnySucceeded;
                    return false;
                }

                dynamic app = _rawAppObj;
                if (app == null) return false;

                // 1. If in normal/slide view and user has actively selected text, replace it (true replacement semantics)
                try
                {
                    dynamic activeWin = null;
                    try { activeWin = app.ActiveWindow; } catch { }
                    if (activeWin != null && IsNormalOrSlideView(activeWin))
                    {
                        dynamic selection = null;
                        try { selection = activeWin.Selection; } catch { }
                        if (selection != null)
                        {
                            int selType = 0;
                            try { selType = Convert.ToInt32(selection.Type); } catch { }
                            // ppSelectionText = 3
                            if (selType == 3 && selection.TextRange != null)
                            {
                                string cleaned = CleanMarkdown(text);
                                try
                                {
                                    // True replacement: set Text directly to preserve formatting context where possible
                                    selection.TextRange.Text = cleaned;
                                }
                                catch
                                {
                                    try { selection.TextRange.Delete(); } catch { }
                                    try { selection.TextRange.InsertAfter(cleaned); } catch { selection.TextRange.Text = cleaned; }
                                }
                                return true;
                            }
                        }
                    }
                }
                catch { }

                // 2. Parse text into structured slides
                var slides = PowerPointActionParser.ParseSlideData(text);
                if (slides.Count == 0)
                {
                    // Fallback to simple bullet insertion
                    AddBulletPoints(new List<string>(text.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)));
                    return true;
                }

                dynamic pres = GetOrCreateActivePresentation();
                if (pres == null) return false;

                dynamic activeSlide = GetOrCreateActiveSlide();
                // If active slide already has user content, do not overwrite it — append new slides instead.
                bool canReuseActiveSlide = activeSlide != null && !SlideHasSubstantiveContent(activeSlide);
                bool isFirst = true;

                foreach (var slideData in slides)
                {
                    dynamic targetSlide;
                    if (isFirst && canReuseActiveSlide)
                    {
                        targetSlide = activeSlide;
                        isFirst = false;
                    }
                    else
                    {
                        // Add subsequent slides using the current template's text layout where possible.
                        int slideCount = Convert.ToInt32(pres.Slides.Count);
                        targetSlide = AddSlideUsingPresentationLayout(pres, slideCount + 1, activeSlide);
                        isFirst = false;
                    }

                    PopulateSlide(targetSlide, slideData);
                }
                return true;
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.InsertText failed", ex);
                return false;
            }
        }

        private bool ExecuteLegacyPowerPointAction(PowerPointAction act)
        {
            if (act == null || string.IsNullOrWhiteSpace(act.Type)) return false;
            string t = act.Type.Trim().ToLowerInvariant();
            try
            {
                switch (t)
                {
                    case "move_slide":
                        return ExecuteMoveSlide(act.Source > 0 ? act.Source : act.Slide, act.Target > 0 ? act.Target : act.Slide).Success;
                    case "create_section":
                        {
                            int s = act.Slide > 0 ? act.Slide : (act.Target > 0 ? act.Target : 1);
                            string n = !string.IsNullOrWhiteSpace(act.Name) ? act.Name : (!string.IsNullOrWhiteSpace(act.Title) ? act.Title : (!string.IsNullOrWhiteSpace(act.Text) ? act.Text : "Section"));
                            return ExecuteCreateSectionBeforeSlide(n, s).Success;
                        }
                    case "rename_section":
                        return ExecuteRenameSectionInPlace(act.Section > 0 ? act.Section : act.Slide, !string.IsNullOrWhiteSpace(act.Name) ? act.Name : act.Text).Success;
                    case "set_notes":
                        return ExecuteSetSpeakerNotesInPlace(act.Slide > 0 ? act.Slide : (act.Target > 0 ? act.Target : 1), !string.IsNullOrWhiteSpace(act.Notes) ? act.Notes : act.Text).Success;
                    case "create_slide":
                        {
                            string title = !string.IsNullOrWhiteSpace(act.Title) ? act.Title : (!string.IsNullOrWhiteSpace(act.Name) ? act.Name : (!string.IsNullOrWhiteSpace(act.Text) ? act.Text : null));
                            string layout = !string.IsNullOrWhiteSpace(act.Layout) ? act.Layout : null;
                            int idx = act.Slide > 0 ? act.Slide : act.Target;
                            // Try to parse bullets from Data if present
                            List<string> bullets = null;
                            if (!string.IsNullOrWhiteSpace(act.Data))
                            {
                                try { bullets = new List<string>(act.Data.Split(new[] { '\r', '\n', ';', '|' }, StringSplitOptions.RemoveEmptyEntries)); } catch { bullets = null; }
                                // clean
                                if (bullets != null) for (int i = 0; i < bullets.Count; i++) bullets[i] = CleanMarkdown(bullets[i]);
                            }
                            else if (!string.IsNullOrWhiteSpace(act.Text) && title != act.Text)
                            {
                                bullets = new List<string> { CleanMarkdown(act.Text) };
                            }
                            string notes = act.Notes;
                            return ExecuteCreateSlide(title, bullets, layout, idx, notes).Success;
                        }
                    case "insert_image":
                    case "add_image":
                        {
                            string path = !string.IsNullOrWhiteSpace(act.ImagePath) ? act.ImagePath : (!string.IsNullOrWhiteSpace(act.Name) ? act.Name : act.Text);
                            int s = act.Slide > 0 ? act.Slide : act.Target;
                            string alt = !string.IsNullOrWhiteSpace(act.AltText) ? act.AltText : act.Title;
                            return ExecuteInsertImage(path, alt, s).Success;
                        }
                    case "delete_slide":
                        return ExecuteDeleteSlide(act.Slide > 0 ? act.Slide : (act.Target > 0 ? act.Target : act.Source)).Success;
                    case "duplicate_slide":
                        return ExecuteDuplicateSlide(act.Slide > 0 ? act.Slide : (act.Target > 0 ? act.Target : act.Source)).Success;
                    case "hide_slide":
                        return ExecuteHideSlide(act.Slide > 0 ? act.Slide : act.Target).Success;
                    case "unhide_slide":
                        return ExecuteUnhideSlide(act.Slide > 0 ? act.Slide : act.Target).Success;
                    case "apply_layout":
                    case "set_layout":
                        return ExecuteApplyLayout(act.Slide > 0 ? act.Slide : act.Target, !string.IsNullOrWhiteSpace(act.Layout) ? act.Layout : (!string.IsNullOrWhiteSpace(act.Name) ? act.Name : act.Text)).Success;
                    case "set_shape_text":
                        return ExecuteSetShapeText(act.Slide > 0 ? act.Slide : act.Target, !string.IsNullOrWhiteSpace(act.ShapeType) ? act.ShapeType : act.Name, !string.IsNullOrWhiteSpace(act.Text) ? act.Text : (act.Data ?? string.Empty)).Success;
                    case "replace_text":
                        return ExecuteReplaceSelectedText(!string.IsNullOrWhiteSpace(act.Text) ? act.Text : (act.Data ?? act.Name ?? string.Empty)).Success;
                    case "add_table":
                        {
                            int s = act.Slide > 0 ? act.Slide : act.Target;
                            int r = act.Rows > 0 ? act.Rows : 2;
                            int c = act.Cols > 0 ? act.Cols : 2;
                            List<List<string>> tableData = null;
                            if (!string.IsNullOrWhiteSpace(act.Data))
                            {
                                try
                                {
                                    // Try JSON 2D array, else pipe-split
                                    if (act.Data.Trim().StartsWith("["))
                                    {
                                        var parsed = Newtonsoft.Json.JsonConvert.DeserializeObject<List<List<string>>>(act.Data);
                                        if (parsed != null) tableData = parsed;
                                    }
                                }
                                catch { }
                                if (tableData == null)
                                {
                                    tableData = new List<List<string>>();
                                    var rowsArr = act.Data.Split(new[] { '\n', ';' }, StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var row in rowsArr)
                                    {
                                        var colsArr = row.Split(new[] { '|', ',', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                                        var cleanRow = new List<string>();
                                        foreach (var cell in colsArr) cleanRow.Add(CleanMarkdown(cell));
                                        if (cleanRow.Count > 0) tableData.Add(cleanRow);
                                    }
                                    if (tableData.Count > 0)
                                    {
                                        if (r < tableData.Count) r = tableData.Count;
                                        foreach (var row in tableData) if (c < row.Count) c = row.Count;
                                    }
                                }
                            }
                            return ExecuteAddTable(s, r, c, tableData).Success;
                        }
                    case "add_chart":
                        {
                            int s = act.Slide > 0 ? act.Slide : act.Target;
                            string ct = !string.IsNullOrWhiteSpace(act.ChartType) ? act.ChartType : (!string.IsNullOrWhiteSpace(act.ShapeType) ? act.ShapeType : "column");
                            string ttl = !string.IsNullOrWhiteSpace(act.Title) ? act.Title : act.Text;
                            return ExecuteAddChart(s, ct, ttl).Success;
                        }
                    case "add_shape":
                        {
                            int s = act.Slide > 0 ? act.Slide : act.Target;
                            string st = !string.IsNullOrWhiteSpace(act.ShapeType) ? act.ShapeType : (!string.IsNullOrWhiteSpace(act.Text) && act.Text.IndexOf("textbox", StringComparison.OrdinalIgnoreCase) >= 0 ? "textbox" : "rectangle");
                            string txt = !string.IsNullOrWhiteSpace(act.Text) ? act.Text : (act.Data ?? string.Empty);
                            // If Data contains shape description, use it
                            if (string.IsNullOrWhiteSpace(txt) && !string.IsNullOrWhiteSpace(act.Title)) txt = act.Title;
                            return ExecuteAddShape(s, st, txt).Success;
                        }
                    case "set_font":
                        return ExecuteSetFont(act.FontName, act.FontSize, act.Bold, act.Italic, act.Color).Success;
                    case "fit_content":
                        return ExecuteFitContent(act.Slide > 0 ? act.Slide : act.Target).Success;
                    default:
                        Logger.Warn(string.Format("ExecuteLegacyPowerPointAction: unhandled type {0}", t));
                        return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ExecuteLegacyPowerPointAction {0} exception: {1}", t, ex.Message));
                return false;
            }
        }

        /// <summary>
        /// Applies a single structured deck action returned by the model, setting status and result/error on the action object.
        /// </summary>
        public HostOperationResult ExecuteCreateDeckFromOutline(string outline)
        {
            if (string.IsNullOrWhiteSpace(outline))
                return HostOperationResult.Failed("Outline content cannot be empty.");

            try
            {
                bool ok = InsertText(outline);
                if (ok)
                    return HostOperationResult.Ok("Deck created or updated from outline.");
                else
                    return HostOperationResult.Failed("Failed to create or update deck from outline.");
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteCreateDeckFromOutline failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteCreateDeckFromOutline");
            }
        }

        public HostOperationResult ExecuteCreateSlide(string title, List<string> bullets, string layout, int slideIndex, string speakerNotes)
        {
            bool hasTitle = !string.IsNullOrWhiteSpace(title);
            bool hasBullets = bullets != null && bullets.Count > 0;
            bool hasNotes = !string.IsNullOrWhiteSpace(speakerNotes);
            if (!hasTitle && !hasBullets && !hasNotes)
                return HostOperationResult.Failed("Slide title, bullets, or speaker notes must be provided.");

            try
            {
                dynamic pres = GetOrCreateActivePresentation();
                if (pres == null) return HostOperationResult.Failed("No active presentation.");

                int count = 0;
                try { count = Convert.ToInt32(pres.Slides.Count); } catch { }

                int insertIndex = slideIndex > 0 ? slideIndex : count + 1;
                if (insertIndex < 1) insertIndex = 1;
                if (insertIndex > count + 1) insertIndex = count + 1;

                dynamic sourceSlide = null;
                try { sourceSlide = GetActiveSlide(); } catch { }

                dynamic newSlide = null;
                if (!string.IsNullOrWhiteSpace(layout))
                {
                    newSlide = AddSlideWithLayout(pres, insertIndex, layout, sourceSlide);
                }
                if (newSlide == null)
                {
                    newSlide = AddSlideUsingPresentationLayout(pres, insertIndex, sourceSlide);
                }
                if (newSlide == null) return HostOperationResult.Failed("Failed to create slide at position " + insertIndex, 0, "Slide " + insertIndex);

                var data = new SlideData();
                data.Title = hasTitle ? CleanMarkdown(title) : null;
                if (hasBullets)
                {
                    foreach (var b in bullets)
                    {
                        string cb = CleanMarkdown(b);
                        if (!string.IsNullOrWhiteSpace(cb)) data.Bullets.Add(cb);
                    }
                }
                data.SpeakerNotes = hasNotes ? CleanMarkdown(speakerNotes) : null;

                PopulateSlide(newSlide, data);
                int actualNum = 0;
                try { actualNum = Convert.ToInt32(newSlide.SlideNumber); } catch { actualNum = insertIndex; }
                return HostOperationResult.Ok(string.Format("Created slide {0}: {1}", actualNum, hasTitle ? title.Trim() : "(untitled)"), "Slide " + actualNum);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteCreateSlide failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteCreateSlide");
            }
        }

        public HostOperationResult ExecuteInsertImage(string filePath, string altText = null)
        {
            return ExecuteInsertImage(filePath, altText, 0);
        }

        public HostOperationResult ExecuteInsertImage(string filePath, string altText, int slideNumber)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return HostOperationResult.Failed("Image file path cannot be empty.");

            // Handle case where entire parameter string was passed (e.g., "image_path=C:\\a.png, slide=2")
            string cleanPath = filePath.Trim().Trim('"').Trim('\'');
            if (cleanPath.IndexOf("image_path", StringComparison.OrdinalIgnoreCase) >= 0 && cleanPath.Contains("="))
            {
                var m = Regex.Match(cleanPath, @"image_path\s*=\s*""?([^"",\s]+)""?", RegexOptions.IgnoreCase);
                if (m.Success && !string.IsNullOrWhiteSpace(m.Groups[1].Value))
                    cleanPath = m.Groups[1].Value.Trim().Trim('"').Trim('\'');
            }
            cleanPath = cleanPath.Trim();

            if (!System.IO.File.Exists(cleanPath))
                return HostOperationResult.Failed(string.Format("Image file not found: {0}", cleanPath));

            try
            {
                bool ok;
                if (slideNumber > 0)
                    ok = InsertImageFromFileOnSlide(cleanPath, altText, slideNumber);
                else
                    ok = InsertImageFromFile(cleanPath, altText);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Inserted image from {0}{1}", System.IO.Path.GetFileName(cleanPath), slideNumber > 0 ? " on slide " + slideNumber : ""), slideNumber > 0 ? "Slide " + slideNumber : null);
                else
                    return HostOperationResult.Failed(string.Format("Failed to insert image from {0}", cleanPath), 0, slideNumber > 0 ? "Slide " + slideNumber : null);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteInsertImage failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteInsertImage", slideNumber > 0 ? "Slide " + slideNumber : null);
            }
        }

        public HostOperationResult ExecuteMoveSlide(int sourceSlideNumber, int destinationSlideNumber)
        {
            if (sourceSlideNumber < 1)
                return HostOperationResult.Failed("Source slide number must be at least 1.", 0, "Slide " + sourceSlideNumber);

            if (destinationSlideNumber < 1)
                return HostOperationResult.Failed("Destination slide number must be at least 1.", 0, "Slide " + destinationSlideNumber);

            try
            {
                bool ok = MoveSlide(sourceSlideNumber, destinationSlideNumber);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Moved slide {0} to position {1}", sourceSlideNumber, destinationSlideNumber), "Slide " + destinationSlideNumber);
                else
                    return HostOperationResult.Failed(string.Format("Failed to move slide {0} to position {1}", sourceSlideNumber, destinationSlideNumber), 0, "Slide " + destinationSlideNumber);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteMoveSlide failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteMoveSlide", "Slide " + destinationSlideNumber);
            }
        }

        public HostOperationResult ExecuteCreateSectionBeforeSlide(string sectionName, int slideNumber)
        {
            if (string.IsNullOrWhiteSpace(sectionName))
                return HostOperationResult.Failed("Section name cannot be empty.", 0, "Slide " + slideNumber);

            if (slideNumber < 1)
                return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);

            try
            {
                bool ok = CreateSectionBeforeSlide(sectionName, slideNumber);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Created section '{0}' before slide {1}", sectionName, slideNumber), "Slide " + slideNumber);
                else
                    return HostOperationResult.Failed(string.Format("Failed to create section '{0}' before slide {1}", sectionName, slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteCreateSectionBeforeSlide failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteCreateSectionBeforeSlide", "Slide " + slideNumber);
            }
        }

        public HostOperationResult ExecuteRenameSectionInPlace(int sectionIndex, string sectionName)
        {
            if (sectionIndex < 1)
                return HostOperationResult.Failed("Section index must be at least 1.", 0, "Section " + sectionIndex);

            if (string.IsNullOrWhiteSpace(sectionName))
                return HostOperationResult.Failed("New section name cannot be empty.", 0, "Section " + sectionIndex);

            try
            {
                bool ok = RenameSection(sectionIndex, sectionName);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Renamed section {0} to '{1}'", sectionIndex, sectionName), "Section " + sectionIndex);
                else
                    return HostOperationResult.Failed(string.Format("Failed to rename section {0} to '{1}'", sectionIndex, sectionName), 0, "Section " + sectionIndex);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteRenameSectionInPlace failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteRenameSectionInPlace", "Section " + sectionIndex);
            }
        }

        public HostOperationResult ExecuteSetSpeakerNotesInPlace(int slideNumber, string notes)
        {
            if (slideNumber < 1)
                return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);

            if (string.IsNullOrWhiteSpace(notes))
                return HostOperationResult.Failed("Speaker notes cannot be empty.", 0, "Slide " + slideNumber);

            try
            {
                bool ok = SetSpeakerNotesForSlide(slideNumber, notes);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Set speaker notes on slide {0}", slideNumber), "Slide " + slideNumber);
                else
                    return HostOperationResult.Failed(string.Format("Failed to set speaker notes on slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteSetSpeakerNotesInPlace failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteSetSpeakerNotesInPlace", "Slide " + slideNumber);
            }
        }

        public HostOperationResult ExecuteDeleteSlide(int slideNumber)
        {
            if (slideNumber < 1) return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = DeleteSlide(slideNumber);
                return ok ? HostOperationResult.Ok(string.Format("Deleted slide {0}", slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to delete slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteDeleteSlide failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteDeleteSlide", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteDuplicateSlide(int slideNumber)
        {
            if (slideNumber < 1) return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = DuplicateSlide(slideNumber);
                return ok ? HostOperationResult.Ok(string.Format("Duplicated slide {0}", slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to duplicate slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteDuplicateSlide failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteDuplicateSlide", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteHideSlide(int slideNumber)
        {
            if (slideNumber < 1) return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = SetSlideHidden(slideNumber, true);
                return ok ? HostOperationResult.Ok(string.Format("Hid slide {0}", slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to hide slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteHideSlide failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteHideSlide", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteUnhideSlide(int slideNumber)
        {
            if (slideNumber < 1) return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = SetSlideHidden(slideNumber, false);
                return ok ? HostOperationResult.Ok(string.Format("Unhid slide {0}", slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to unhide slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteUnhideSlide failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteUnhideSlide", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteApplyLayout(int slideNumber, string layoutName)
        {
            if (slideNumber < 1) return HostOperationResult.Failed("Slide number must be at least 1.", 0, "Slide " + slideNumber);
            if (string.IsNullOrWhiteSpace(layoutName)) return HostOperationResult.Failed("Layout name cannot be empty.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = ApplyLayoutToSlide(slideNumber, layoutName);
                return ok ? HostOperationResult.Ok(string.Format("Applied layout '{0}' to slide {1}", layoutName, slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to apply layout '{0}' to slide {1}", layoutName, slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteApplyLayout failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteApplyLayout", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteSetShapeText(int slideNumber, string shapeNameOrIndex, string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return HostOperationResult.Failed("Text cannot be empty.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = SetShapeText(slideNumber, shapeNameOrIndex, text);
                return ok ? HostOperationResult.Ok(string.Format("Updated shape text on slide {0}", slideNumber), "Slide " + slideNumber) : HostOperationResult.Failed(string.Format("Failed to update shape text on slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteSetShapeText failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteSetShapeText", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteReplaceSelectedText(string newText)
        {
            if (string.IsNullOrWhiteSpace(newText)) return HostOperationResult.Failed("Replacement text cannot be empty.");
            try
            {
                bool ok = ReplaceSelectedText(newText);
                return ok ? HostOperationResult.Ok("Replaced selected text.") : HostOperationResult.Failed("Failed to replace selected text.");
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteReplaceSelectedText failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteReplaceSelectedText"); }
        }

        public HostOperationResult ExecuteAddTable(int slideNumber, int rows, int cols, List<List<string>> data)
        {
            if (rows < 1 || cols < 1) return HostOperationResult.Failed("Rows and cols must be >= 1.", 0, "Slide " + slideNumber);
            try
            {
                bool ok = AddTableToSlide(slideNumber, rows, cols, data);
                return ok ? HostOperationResult.Ok(string.Format("Added {0}x{1} table to slide {2}", rows, cols, slideNumber > 0 ? slideNumber.ToString() : "active"), slideNumber > 0 ? "Slide " + slideNumber : null) : HostOperationResult.Failed(string.Format("Failed to add table to slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteAddTable failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteAddTable", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteAddChart(int slideNumber, string chartType, string title)
        {
            try
            {
                bool ok = AddChartToSlide(slideNumber, chartType, title);
                return ok ? HostOperationResult.Ok(string.Format("Added {0} chart to slide {1}", chartType ?? "column", slideNumber > 0 ? slideNumber.ToString() : "active"), slideNumber > 0 ? "Slide " + slideNumber : null) : HostOperationResult.Failed(string.Format("Failed to add chart to slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteAddChart failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteAddChart", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteAddShape(int slideNumber, string shapeType, string text)
        {
            try
            {
                bool ok = AddShapeToSlide(slideNumber, shapeType, text, 100, 100, 200, 80);
                return ok ? HostOperationResult.Ok(string.Format("Added {0} shape to slide {1}", shapeType ?? "rectangle", slideNumber > 0 ? slideNumber.ToString() : "active"), slideNumber > 0 ? "Slide " + slideNumber : null) : HostOperationResult.Failed(string.Format("Failed to add shape to slide {0}", slideNumber), 0, "Slide " + slideNumber);
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteAddShape failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteAddShape", "Slide " + slideNumber); }
        }

        public HostOperationResult ExecuteSetFont(string fontName, string fontSize, string bold, string italic, string color)
        {
            try
            {
                bool ok = SetFontForSelection(fontName, fontSize, bold, italic, color);
                return ok ? HostOperationResult.Ok("Applied font formatting.") : HostOperationResult.Failed("Failed to apply font formatting - no text selection found.");
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteSetFont failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteSetFont"); }
        }

        public HostOperationResult ExecuteFitContent(int slideNumber)
        {
            try
            {
                bool ok = FitContentToSlide(slideNumber);
                return ok ? HostOperationResult.Ok(string.Format("Fitted content on slide {0}", slideNumber > 0 ? slideNumber.ToString() : "active")) : HostOperationResult.Failed("Fit content had no effect.");
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteFitContent failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteFitContent"); }
        }

        // D-13 Tier 2: mutation methods that lacked a structured HostOperationResult wrapper.
        public HostOperationResult ExecuteUndo()
        {
            try
            {
                bool ok = Undo();
                return ok ? HostOperationResult.Ok("Undid the last PowerPoint change.") : HostOperationResult.Failed("Undo returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteUndo failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteUndo");
            }
        }

        public HostOperationResult ExecuteInsertText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return HostOperationResult.Failed("Text cannot be empty.");

            try
            {
                bool ok = InsertText(text);
                return ok ? HostOperationResult.Ok("Inserted text into the presentation.") : HostOperationResult.Failed("InsertText returned false.");
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteInsertText failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteInsertText");
            }
        }

        public virtual bool MoveSlide(int sourceSlideNumber, int destinationSlideNumber)
        {
            dynamic presentation = GetActivePresentation();
            if (presentation == null || presentation.Slides == null) return false;
            int count = Convert.ToInt32(presentation.Slides.Count);
            if (sourceSlideNumber < 1 || sourceSlideNumber > count || destinationSlideNumber < 1 || destinationSlideNumber > count)
                return false;
            dynamic slide = presentation.Slides[sourceSlideNumber];
            slide.MoveTo(destinationSlideNumber);
            return true;
        }

        public virtual bool CreateSectionBeforeSlide(string sectionName, int slideNumber)
        {
            if (string.IsNullOrWhiteSpace(sectionName)) return false;
            dynamic presentation = GetActivePresentation();
            if (presentation == null || presentation.SectionProperties == null) return false;
            int count = Convert.ToInt32(presentation.Slides.Count);
            if (slideNumber < 1 || slideNumber > count + 1) return false;

            if (slideNumber <= count)
            {
                presentation.SectionProperties.AddBeforeSlide(slideNumber, sectionName.Trim());
            }
            else
            {
                // AddBeforeSlide requires an existing slide.  AddSection is the COM API
                // that creates an empty section at the end of a presentation.
                int sectionCount = Convert.ToInt32(presentation.SectionProperties.Count);
                presentation.SectionProperties.AddSection(sectionCount + 1, sectionName.Trim());
            }
            return true;
        }

        public virtual bool RenameSection(int sectionIndex, string sectionName)
        {
            if (sectionIndex < 1 || string.IsNullOrWhiteSpace(sectionName)) return false;
            dynamic presentation = GetActivePresentation();
            if (presentation == null || presentation.SectionProperties == null) return false;
            presentation.SectionProperties.Rename(sectionIndex, sectionName.Trim());
            return true;
        }

        public virtual bool SetSpeakerNotesForSlide(int slideNumber, string notes)
        {
            if (slideNumber < 1 || string.IsNullOrWhiteSpace(notes)) return false;
            dynamic presentation = GetActivePresentation();
            if (presentation == null || presentation.Slides == null) return false;
            int count = Convert.ToInt32(presentation.Slides.Count);
            if (slideNumber > count) return false;
            dynamic slide = presentation.Slides[slideNumber];
            dynamic shapes = slide.NotesPage != null ? slide.NotesPage.Shapes : null;
            int shapeCount = shapes != null ? Convert.ToInt32(shapes.Count) : 0;
            for (int i = 1; i <= shapeCount; i++)
            {
                dynamic shape = shapes[i];
                if (shape == null || shape.PlaceholderFormat == null) continue;
                if (Convert.ToInt32(shape.PlaceholderFormat.Type) == 2 && Convert.ToInt32(shape.HasTextFrame) != 0)
                {
                    shape.TextFrame.TextRange.Text = CleanMarkdown(notes);
                    return true;
                }
            }
            return false;
        }

        public string GetSpeakerNotesForSlide(int slideNumber)
        {
            if (slideNumber < 1) return string.Empty;
            try
            {
                dynamic presentation = GetActivePresentation();
                if (presentation == null || presentation.Slides == null) return string.Empty;
                int count = Convert.ToInt32(presentation.Slides.Count);
                if (slideNumber > count) return string.Empty;
                dynamic slide = presentation.Slides[slideNumber];
                dynamic shapes = slide.NotesPage != null ? slide.NotesPage.Shapes : null;
                int shapeCount = shapes != null ? Convert.ToInt32(shapes.Count) : 0;
                for (int i = 1; i <= shapeCount; i++)
                {
                    dynamic shape = shapes[i];
                    if (shape == null || shape.PlaceholderFormat == null) continue;
                    if (Convert.ToInt32(shape.PlaceholderFormat.Type) == 2 && Convert.ToInt32(shape.HasTextFrame) != 0)
                    {
                        return shape.TextFrame.TextRange.Text ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetSpeakerNotesForSlide failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        /// <summary>
        /// Supports safe, local visual insertion. Generation/search is deliberately provider-neutral;
        /// callers can create a local image and pass its verified path here.
        /// </summary>
        public bool InsertImageFromFile(string filePath, string altText)
        {
            return InsertImageFromFileOnSlide(filePath, altText, 0);
        }

        public bool InsertImageFromFileOnSlide(string filePath, string altText, int slideNumber)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return false;
            try
            {
                dynamic slide = null;
                if (slideNumber > 0)
                {
                    dynamic pres = GetActivePresentation();
                    if (pres == null || pres.Slides == null) return false;
                    int count = 0;
                    try { count = Convert.ToInt32(pres.Slides.Count); } catch { }
                    if (slideNumber < 1 || slideNumber > count) return false;
                    try { slide = pres.Slides[slideNumber]; } catch { return false; }
                }
                else
                {
                    slide = GetOrCreateActiveSlide();
                }
                if (slide == null || slide.Shapes == null) return false;
                dynamic picture = slide.Shapes.AddPicture(filePath, 0, -1, 70, 115, -1, -1);
                if (picture != null && !string.IsNullOrWhiteSpace(altText))
                {
                    try { picture.AlternativeText = altText.Trim(); } catch { }
                }
                return picture != null;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.InsertImageFromFileOnSlide failed: {0}", ex.Message));
                return false;
            }
        }

        // ========== Slide Lifecycle ==========

        public virtual bool DeleteSlide(int slideNumber)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                slide.Delete();
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.DeleteSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool DuplicateSlide(int slideNumber)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                dynamic dup = slide.Duplicate();
                // Duplicate returns SlideRange; move duplicated slide to right after original for predictable ordering
                if (dup != null && dup.Count > 0)
                {
                    try
                    {
                        dynamic newSlide = dup[1];
                        int newIdx = Convert.ToInt32(newSlide.SlideIndex);
                        if (newIdx != slideNumber + 1)
                        {
                            newSlide.MoveTo(slideNumber + 1);
                        }
                    }
                    catch { }
                }
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.DuplicateSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool SetSlideHidden(int slideNumber, bool hidden)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                // SlideShowTransition.Hidden = msoTrue (-1) / msoFalse (0)
                slide.SlideShowTransition.Hidden = hidden ? -1 : 0;
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.SetSlideHidden failed: {0}", ex.Message)); return false; }
        }

        public bool IsSlideHidden(int slideNumber)
        {
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return false;
                int count = Convert.ToInt32(pres.Slides.Count);
                if (slideNumber < 1 || slideNumber > count) return false;
                dynamic slide = pres.Slides[slideNumber];
                int hidden = Convert.ToInt32(slide.SlideShowTransition.Hidden);
                return hidden != 0;
            }
            catch { return false; }
        }

        // ========== Layout / Formatting / Content ==========

        public virtual bool ApplyLayoutToSlide(int slideNumber, string layoutName)
        {
            if (string.IsNullOrWhiteSpace(layoutName)) return false;
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                dynamic customLayout = FindCustomLayoutByName(pres, layoutName);
                if (customLayout != null)
                {
                    try { slide.CustomLayout = customLayout; return true; } catch (Exception ex) { Logger.Warn(string.Format("ApplyLayout CustomLayout failed: {0}", ex.Message)); }
                }
                // Fallback to PpSlideLayout enum
                int layoutEnum = MapLayoutNameToEnum(layoutName);
                if (layoutEnum > 0)
                {
                    try { slide.Layout = layoutEnum; return true; } catch (Exception ex) { Logger.Warn(string.Format("ApplyLayout enum failed: {0}", ex.Message)); }
                }
                return false;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.ApplyLayoutToSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool SetShapeText(int slideNumber, string shapeNameOrIndex, string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                dynamic shapes = slide.Shapes;
                int sc = Convert.ToInt32(shapes.Count);
                dynamic targetShape = null;
                // Try by index if numeric
                int idx;
                if (!string.IsNullOrWhiteSpace(shapeNameOrIndex) && int.TryParse(shapeNameOrIndex.Trim(), out idx) && idx >= 1 && idx <= sc)
                {
                    try { targetShape = shapes[idx]; } catch { }
                }
                // Try by name
                if (targetShape == null && !string.IsNullOrWhiteSpace(shapeNameOrIndex))
                {
                    string key = shapeNameOrIndex.Trim();
                    for (int i = 1; i <= sc; i++)
                    {
                        try
                        {
                            dynamic s = shapes[i];
                            string n = Convert.ToString(s.Name) ?? string.Empty;
                            if (n.Equals(key, StringComparison.OrdinalIgnoreCase)) { targetShape = s; break; }
                        }
                        catch { }
                    }
                }
                // Fallback to first text-containing shape
                if (targetShape == null)
                {
                    for (int i = 1; i <= sc; i++)
                    {
                        try
                        {
                            dynamic s = shapes[i];
                            if (Convert.ToInt32(s.HasTextFrame) != 0) { targetShape = s; break; }
                        }
                        catch { }
                    }
                }
                if (targetShape == null || Convert.ToInt32(targetShape.HasTextFrame) == 0) return false;
                targetShape.TextFrame.TextRange.Text = CleanMarkdown(text);
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.SetShapeText failed: {0}", ex.Message)); return false; }
        }

        public virtual bool SetSlideText(int slideNumber, string title, string bodyText)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            if (slideNumber < 1 || slideNumber > count) return false;
            try
            {
                dynamic slide = pres.Slides[slideNumber];
                var data = new SlideData();
                if (!string.IsNullOrWhiteSpace(title)) data.Title = title;
                if (!string.IsNullOrWhiteSpace(bodyText))
                {
                    var lines = bodyText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var l in lines) data.Bullets.Add(l.Trim());
                }
                PopulateSlide(slide, data);
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.SetSlideText failed: {0}", ex.Message)); return false; }
        }

        public virtual bool ReplaceSelectedText(string newText)
        {
            if (string.IsNullOrWhiteSpace(newText)) return false;
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return false;
                dynamic activeWin = null;
                try { activeWin = app.ActiveWindow; } catch { }
                if (activeWin != null && IsNormalOrSlideView(activeWin))
                {
                    dynamic selection = null;
                    try { selection = activeWin.Selection; } catch { }
                    if (selection != null)
                    {
                        int selType = 0;
                        try { selType = Convert.ToInt32(selection.Type); } catch { }
                        if (selType == 3 && selection.TextRange != null)
                        {
                            string cleaned = CleanMarkdown(newText);
                            try { selection.TextRange.Text = cleaned; return true; }
                            catch
                            {
                                try { selection.TextRange.Delete(); selection.TextRange.InsertAfter(cleaned); return true; } catch { }
                            }
                        }
                    }
                    // Fallback: try ShapeRange.TextFrame.TextRange if ppSelectionShapes = 2
                    try
                    {
                        int selType = Convert.ToInt32(activeWin.Selection.Type);
                        if (selType == 2 && activeWin.Selection.ShapeRange != null)
                        {
                            dynamic shapeRange = activeWin.Selection.ShapeRange;
                            int sc = Convert.ToInt32(shapeRange.Count);
                            for (int i = 1; i <= sc; i++)
                            {
                                dynamic shape = shapeRange[i];
                                if (shape != null && Convert.ToInt32(shape.HasTextFrame) != 0 && Convert.ToInt32(shape.TextFrame.HasText) != 0)
                                {
                                    shape.TextFrame.TextRange.Text = CleanMarkdown(newText);
                                    return true;
                                }
                            }
                        }
                    }
                    catch { }
                }
                // Final fallback: replace text on active slide's body placeholder
                return SetSlideTextForActiveSlide(newText);
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.ReplaceSelectedText failed: {0}", ex.Message)); return false; }
        }

        private bool SetSlideTextForActiveSlide(string newText)
        {
            try
            {
                dynamic slide = GetActiveSlide();
                if (slide == null) return false;
                var data = new SlideData();
                var lines = newText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                bool first = true;
                foreach (var l in lines)
                {
                    string trimmed = l.Trim();
                    if (string.IsNullOrWhiteSpace(trimmed)) continue;
                    if (first && string.IsNullOrWhiteSpace(data.Title) && lines.Length > 1)
                    {
                        data.Title = trimmed;
                        first = false;
                    }
                    else
                    {
                        data.Bullets.Add(trimmed);
                        first = false;
                    }
                }
                PopulateSlide(slide, data);
                return true;
            }
            catch { return false; }
        }

        public virtual bool AddTableToSlide(int slideNumber, int rows, int cols, List<List<string>> data)
        {
            if (rows < 1 || cols < 1) return false;
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }

            dynamic slide = null;
            if (slideNumber > 0)
            {
                if (slideNumber > count) return false;
                try { slide = pres.Slides[slideNumber]; } catch { return false; }
            }
            else
            {
                slide = GetOrCreateActiveSlide();
            }
            if (slide == null || slide.Shapes == null) return false;

            try
            {
                dynamic tableShape = slide.Shapes.AddTable(rows, cols, 60, 140, 600, 300);
                if (tableShape == null || tableShape.Table == null) return false;
                dynamic table = tableShape.Table;
                if (data != null)
                {
                    for (int r = 0; r < Math.Min(rows, data.Count); r++)
                    {
                        var row = data[r];
                        for (int c = 0; c < Math.Min(cols, row.Count); c++)
                        {
                            try
                            {
                                dynamic cell = table.Cell(r + 1, c + 1);
                                dynamic cellShape = cell.Shape;
                                cellShape.TextFrame.TextRange.Text = row[c] ?? string.Empty;
                            }
                            catch { }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddTableToSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool AddChartToSlide(int slideNumber, string chartType, string title)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            dynamic slide = null;
            if (slideNumber > 0)
            {
                if (slideNumber > count) return false;
                try { slide = pres.Slides[slideNumber]; } catch { return false; }
            }
            else slide = GetOrCreateActiveSlide();
            if (slide == null || slide.Shapes == null) return false;

            try
            {
                int xlType = MapChartTypeToXl(chartType);
                dynamic chart = null;
                try { chart = slide.Shapes.AddChart2(201, xlType, 70, 120, 500, 300); }
                catch { chart = slide.Shapes.AddChart(xlType, 70, 120, 500, 300); }
                if (chart != null && !string.IsNullOrWhiteSpace(title))
                {
                    try
                    {
                        dynamic chartObj = chart.Chart;
                        chartObj.HasTitle = true;
                        chartObj.ChartTitle.Text = title.Trim();
                    }
                    catch { }
                }
                return chart != null;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddChartToSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool AddShapeToSlide(int slideNumber, string shapeType, string text, int left, int top, int width, int height)
        {
            dynamic pres = GetActivePresentation();
            if (pres == null || pres.Slides == null) return false;
            int count = 0;
            try { count = Convert.ToInt32(pres.Slides.Count); } catch { return false; }
            dynamic slide = null;
            if (slideNumber > 0)
            {
                if (slideNumber > count) return false;
                try { slide = pres.Slides[slideNumber]; } catch { return false; }
            }
            else slide = GetOrCreateActiveSlide();
            if (slide == null || slide.Shapes == null) return false;

            try
            {
                int msoType = MapShapeTypeToMso(shapeType);
                dynamic shape = null;
                if (msoType == 17) // textbox - use true textbox API, not rectangle
                {
                    // msoTextOrientationHorizontal = 1
                    shape = slide.Shapes.AddTextbox(1, left > 0 ? left : 100, top > 0 ? top : 100, width > 0 ? width : 200, height > 0 ? height : 80);
                }
                else
                {
                    shape = slide.Shapes.AddShape(msoType, left > 0 ? left : 100, top > 0 ? top : 100, width > 0 ? width : 200, height > 0 ? height : 80);
                }
                if (shape != null && !string.IsNullOrWhiteSpace(text))
                {
                    try { shape.TextFrame.TextRange.Text = text; } catch { }
                    try { shape.TextFrame.AutoSize = 1; } catch { }
                }
                return shape != null;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddShapeToSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual bool SetFontForSelection(string fontName, string fontSizeStr, string boldStr, string italicStr, string colorStr)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null) return false;
                dynamic activeWin = null;
                try { activeWin = app.ActiveWindow; } catch { }
                if (activeWin == null) return false;

                dynamic selection = null;
                try { selection = activeWin.Selection; } catch { }
                if (selection == null) return false;

                dynamic textRange = null;
                try
                {
                    int selType = Convert.ToInt32(selection.Type);
                    if (selType == 3 && selection.TextRange != null) textRange = selection.TextRange;
                    else if (selType == 2 && selection.ShapeRange != null)
                    {
                        // Apply to first text-containing shape in selection
                        dynamic shapeRange = selection.ShapeRange;
                        int sc = Convert.ToInt32(shapeRange.Count);
                        for (int i = 1; i <= sc; i++)
                        {
                            dynamic shape = shapeRange[i];
                            if (shape != null && Convert.ToInt32(shape.HasTextFrame) != 0) { textRange = shape.TextFrame.TextRange; break; }
                        }
                    }
                }
                catch { }

                if (textRange == null) return false;

                dynamic font = textRange.Font;
                if (!string.IsNullOrWhiteSpace(fontName)) try { font.Name = fontName.Trim(); } catch { }
                if (!string.IsNullOrWhiteSpace(fontSizeStr))
                {
                    float sz;
                    if (float.TryParse(fontSizeStr, out sz) && sz >= 6 && sz <= 72) try { font.Size = sz; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(boldStr))
                {
                    bool b = boldStr.Equals("true", StringComparison.OrdinalIgnoreCase) || boldStr == "1";
                    try { font.Bold = b ? -1 : 0; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(italicStr))
                {
                    bool it = italicStr.Equals("true", StringComparison.OrdinalIgnoreCase) || italicStr == "1";
                    try { font.Italic = it ? -1 : 0; } catch { }
                }
                if (!string.IsNullOrWhiteSpace(colorStr))
                {
                    try
                    {
                        int rgb = ParseColorToRgb(colorStr);
                        font.Color.RGB = rgb;
                    }
                    catch { }
                }
                return true;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.SetFontForSelection failed: {0}", ex.Message)); return false; }
        }

        public virtual bool FitContentToSlide(int slideNumber)
        {
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null) return false;
                int count = Convert.ToInt32(pres.Slides.Count);
                dynamic slide = null;
                if (slideNumber > 0)
                {
                    if (slideNumber > count) return false;
                    slide = pres.Slides[slideNumber];
                }
                else slide = GetActiveSlide();
                if (slide == null || slide.Shapes == null) return false;

                int sc = Convert.ToInt32(slide.Shapes.Count);
                bool anyFitted = false;
                for (int i = 1; i <= sc; i++)
                {
                    try
                    {
                        dynamic shape = slide.Shapes[i];
                        if (shape == null || Convert.ToInt32(shape.HasTextFrame) == 0) continue;
                        dynamic tf = shape.TextFrame;
                        if (tf == null) continue;
                        // AutoSize handling: if text overflows shape, shrink font incrementally
                        try
                        {
                            // Try PowerPoint AutoFit first
                            tf.AutoSize = 1; // ppAutoSizeShapeToFitText
                            anyFitted = true;
                        }
                        catch
                        {
                            try
                            {
                                dynamic tr = tf.TextRange;
                                float fontSize = Convert.ToSingle(tr.Font.Size);
                                while (fontSize > 8 && tf.TextRange.BoundHeight > shape.Height)
                                {
                                    fontSize -= 0.5f;
                                    tr.Font.Size = fontSize;
                                }
                                anyFitted = true;
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
                return anyFitted;
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.FitContentToSlide failed: {0}", ex.Message)); return false; }
        }

        public virtual string GetShapeTextForRollback(int slideNumber, string shapeNameOrIndex)
        {
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return string.Empty;
                int count = Convert.ToInt32(pres.Slides.Count);
                if (slideNumber < 1 || slideNumber > count) return string.Empty;
                dynamic slide = pres.Slides[slideNumber];
                dynamic shapes = slide.Shapes;
                int sc = Convert.ToInt32(shapes.Count);
                dynamic targetShape = null;
                int idx;
                if (!string.IsNullOrWhiteSpace(shapeNameOrIndex) && int.TryParse(shapeNameOrIndex.Trim(), out idx) && idx >= 1 && idx <= sc)
                    try { targetShape = shapes[idx]; } catch { }
                if (targetShape == null && !string.IsNullOrWhiteSpace(shapeNameOrIndex))
                {
                    string key = shapeNameOrIndex.Trim();
                    for (int i = 1; i <= sc; i++) { try { dynamic s = shapes[i]; string n = Convert.ToString(s.Name) ?? ""; if (n.Equals(key, StringComparison.OrdinalIgnoreCase)) { targetShape = s; break; } } catch { } }
                }
                if (targetShape == null) for (int i = 1; i <= sc; i++) { try { dynamic s = shapes[i]; if (Convert.ToInt32(s.HasTextFrame) != 0) { targetShape = s; break; } } catch { } }
                if (targetShape == null || Convert.ToInt32(targetShape.HasTextFrame) == 0) return string.Empty;
                return Convert.ToString(targetShape.TextFrame.TextRange.Text) ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        public HostOperationResult ExecuteTranslateDeck(string targetLanguage)
        {
            if (string.IsNullOrWhiteSpace(targetLanguage)) return HostOperationResult.Failed("Target language is required.");
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return HostOperationResult.Failed("No active presentation.");
                int count = Convert.ToInt32(pres.Slides.Count);
                var sb = new StringBuilder();
                sb.AppendLine(string.Format("[Translate deck to {0}] — per-slide actions would be:", targetLanguage.Trim()));
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic slide = pres.Slides[i];
                        string title = GetSlideTitle(slide);
                        sb.AppendLine(string.Format("Slide {0}: {1} -> [{0} translated]", i, string.IsNullOrWhiteSpace(title) ? "(untitled)" : title));
                    }
                    catch { }
                }
                sb.AppendLine("Model should emit powerpoint.set_shape_text / replace_text actions per slide; this is a read-only audit preview.");
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteTranslateDeck failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteTranslateDeck"); }
        }

        public HostOperationResult ExecuteAuditDeck()
        {
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return HostOperationResult.Failed("No active presentation.");
                int count = Convert.ToInt32(pres.Slides.Count);
                var sb = new StringBuilder();
                sb.AppendLine(string.Format("[Deck Audit | Slides: {0}]", count));
                int untitled = 0, noBody = 0, hidden = 0, missingAlt = 0;
                var titles = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic slide = pres.Slides[i];
                        string title = GetSlideTitle(slide);
                        if (string.IsNullOrWhiteSpace(title)) untitled++; else { int occ; titles.TryGetValue(title.Trim(), out occ); titles[title.Trim()] = occ + 1; }
                        string body = GetSlideTextInternal(slide, false);
                        if (CountBodyLines(body) == 0) noBody++;
                        if (IsSlideHidden(i)) hidden++;
                        // alt text check
                        try
                        {
                            dynamic shapes = slide.Shapes;
                            int sc = Convert.ToInt32(shapes.Count);
                            for (int s = 1; s <= sc; s++)
                            {
                                try { dynamic sh = shapes[s]; int type = Convert.ToInt32(sh.Type); if (type == 13 || type == 11) { string alt = Convert.ToString(sh.AlternativeText) ?? ""; if (string.IsNullOrWhiteSpace(alt)) missingAlt++; } } catch { }
                            }
                        }
                        catch { }
                        // font outliers: quick check max font size on slide
                        try
                        {
                            dynamic shapes2 = slide.Shapes;
                            int sc2 = Convert.ToInt32(shapes2.Count);
                            for (int s = 1; s <= sc2; s++) { try { dynamic sh = shapes2[s]; if (Convert.ToInt32(sh.HasTextFrame) != 0) { float sz = Convert.ToSingle(sh.TextFrame.TextRange.Font.Size); if (sz > 32 || sz < 8) { sb.AppendLine(string.Format("  Slide {0}: font size outlier {1}pt on shape '{2}'", i, sz, Convert.ToString(sh.Name))); break; } } } catch { } }
                        }
                        catch { }
                    }
                    catch { }
                }
                var dups = new List<string>(); foreach (var kv in titles) if (kv.Value > 1) dups.Add(kv.Key);
                sb.Insert(0, string.Format("Untitled:{0} NoBody:{1} Hidden:{2} MissingAlt:{3} Duplicates:{4}\n", untitled, noBody, hidden, missingAlt, dups.Count == 0 ? "none" : string.Join(", ", dups.ToArray())));
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteAuditDeck failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteAuditDeck"); }
        }

        public HostOperationResult ExecuteAuditAltText()
        {
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return HostOperationResult.Failed("No active presentation.");
                int count = Convert.ToInt32(pres.Slides.Count);
                var sb = new StringBuilder();
                int totalPics = 0, missing = 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic slide = pres.Slides[i];
                        dynamic shapes = slide.Shapes;
                        int sc = Convert.ToInt32(shapes.Count);
                        for (int s = 1; s <= sc; s++)
                        {
                            try { dynamic sh = shapes[s]; int type = Convert.ToInt32(sh.Type); if (type == 13 || type == 11) { totalPics++; string alt = Convert.ToString(sh.AlternativeText) ?? ""; if (string.IsNullOrWhiteSpace(alt)) { missing++; sb.AppendLine(string.Format("Slide {0} shape '{1}' missing alt text", i, Convert.ToString(sh.Name))); } } } catch { }
                        }
                    }
                    catch { }
                }
                if (totalPics == 0) return HostOperationResult.Ok("No pictures found.");
                sb.Insert(0, string.Format("Pictures:{0} MissingAlt:{1}\n", totalPics, missing));
                return HostOperationResult.Ok(sb.ToString().TrimEnd());
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteAuditAltText failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteAuditAltText"); }
        }

        public HostOperationResult ExecuteSetAltText(int slideNumber, string shapeNameOrIndex, string altText)
        {
            if (string.IsNullOrWhiteSpace(altText)) return HostOperationResult.Failed("Alt text cannot be empty.", 0, "Slide " + slideNumber);
            try
            {
                dynamic pres = GetActivePresentation();
                if (pres == null || pres.Slides == null) return HostOperationResult.Failed("No active presentation.", 0, "Slide " + slideNumber);
                int count = Convert.ToInt32(pres.Slides.Count);
                if (slideNumber < 1 || slideNumber > count) return HostOperationResult.Failed(string.Format("Slide {0} not found.", slideNumber), 0, "Slide " + slideNumber);
                dynamic slide = pres.Slides[slideNumber];
                dynamic shapes = slide.Shapes;
                int sc = Convert.ToInt32(shapes.Count);
                dynamic target = null;
                int idx;
                if (!string.IsNullOrWhiteSpace(shapeNameOrIndex) && int.TryParse(shapeNameOrIndex.Trim(), out idx) && idx >= 1 && idx <= sc) try { target = shapes[idx]; } catch { }
                if (target == null && !string.IsNullOrWhiteSpace(shapeNameOrIndex)) { string key = shapeNameOrIndex.Trim(); for (int i = 1; i <= sc; i++) { try { dynamic s = shapes[i]; string n = Convert.ToString(s.Name) ?? ""; if (n.Equals(key, StringComparison.OrdinalIgnoreCase)) { target = s; break; } } catch { } } }
                if (target == null) return HostOperationResult.Failed(string.Format("Shape '{0}' not found on slide {1}.", shapeNameOrIndex, slideNumber), 0, "Slide " + slideNumber);
                target.AlternativeText = altText.Trim();
                return HostOperationResult.Ok(string.Format("Set alt text on slide {0} shape '{1}'.", slideNumber, shapeNameOrIndex));
            }
            catch (Exception ex) { Logger.Error("PowerPointController.ExecuteSetAltText failed", ex); return HostOperationResult.FromException(ex, "PowerPointController.ExecuteSetAltText", "Slide " + slideNumber); }
        }

        private static void AppendBounded(StringBuilder builder, string value, int maxCharacters)
        {
            if (builder == null || string.IsNullOrEmpty(value) || builder.Length >= maxCharacters) return;
            int remaining = maxCharacters - builder.Length;
            if (value.Length <= remaining)
                builder.Append(value);
            else
                builder.Append(value.Substring(0, remaining));
        }

        private static int ParsePositiveInt(string value)
        {
            int parsed;
            return int.TryParse(value, out parsed) && parsed > 0 ? parsed : 0;
        }

        private string GetSlideTextInternal(dynamic slide, bool includeSpeakerNotes)
        {
            var sb = new StringBuilder();
            if (slide == null) return string.Empty;

            try
            {
                sb.AppendLine(string.Format("[Slide #{0}: {1}]", slide.SlideNumber, GetSlideTitle(slide)));
            }
            catch
            {
                sb.AppendLine("[Slide]");
            }

            try
            {
                dynamic shapes = slide.Shapes;
                int count = shapes != null ? Convert.ToInt32(shapes.Count) : 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        if (shape == null || Convert.ToInt32(shape.HasTextFrame) == 0) continue;
                        dynamic textFrame = shape.TextFrame;
                        if (textFrame == null || Convert.ToInt32(textFrame.HasText) == 0) continue;
                        string text = Convert.ToString(textFrame.TextRange.Text);
                        if (!string.IsNullOrWhiteSpace(text)) sb.AppendLine(text.Trim());
                    }
                    catch { }
                }
            }
            catch { }

            if (includeSpeakerNotes)
            {
                string notes = GetSpeakerNotesInternal(slide);
                if (!string.IsNullOrWhiteSpace(notes))
                {
                    sb.AppendLine("[Speaker Notes]");
                    sb.AppendLine(notes.Trim());
                }
            }
            return sb.ToString().TrimEnd();
        }

        private string GetSlideTitle(dynamic slide)
        {
            if (slide == null) return string.Empty;
            try
            {
                dynamic shapes = slide.Shapes;
                if (shapes != null)
                {
                    try
                    {
                        if (Convert.ToInt32(shapes.HasTitle) != 0 && shapes.Title != null &&
                            Convert.ToInt32(shapes.Title.HasTextFrame) != 0)
                        {
                            string title = Convert.ToString(shapes.Title.TextFrame.TextRange.Text);
                            if (!string.IsNullOrWhiteSpace(title)) return title.Trim();
                        }
                    }
                    catch { }

                    int count = Convert.ToInt32(shapes.Count);
                    for (int i = 1; i <= count; i++)
                    {
                        try
                        {
                            dynamic shape = shapes[i];
                            if (shape == null || Convert.ToInt32(shape.Type) != 14 || shape.PlaceholderFormat == null) continue;
                            int placeholderType = Convert.ToInt32(shape.PlaceholderFormat.Type);
                            if (placeholderType != 1 && placeholderType != 3) continue;
                            if (Convert.ToInt32(shape.HasTextFrame) == 0) continue;
                            string title = Convert.ToString(shape.TextFrame.TextRange.Text);
                            if (!string.IsNullOrWhiteSpace(title)) return title.Trim();
                        }
                        catch { }
                    }
                }
            }
            catch { }
            try { return Convert.ToString(slide.Name) ?? string.Empty; } catch { return string.Empty; }
        }

        private string GetSpeakerNotesInternal(dynamic slide)
        {
            if (slide == null) return string.Empty;
            try
            {
                dynamic notesPage = slide.NotesPage;
                dynamic shapes = notesPage != null ? notesPage.Shapes : null;
                int count = shapes != null ? Convert.ToInt32(shapes.Count) : 0;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic shape = shapes[i];
                        if (shape == null || shape.PlaceholderFormat == null) continue;
                        // ppPlaceholderBody = 2
                        if (Convert.ToInt32(shape.PlaceholderFormat.Type) != 2 || Convert.ToInt32(shape.HasTextFrame) == 0) continue;
                        string notes = Convert.ToString(shape.TextFrame.TextRange.Text);
                        if (!string.IsNullOrWhiteSpace(notes)) return notes.Trim();
                    }
                    catch { }
                }
            }
            catch { }
            return string.Empty;
        }

        private string GetSectionName(dynamic presentation, int slideNumber)
        {
            try
            {
                dynamic sections = presentation != null ? presentation.SectionProperties : null;
                int sectionCount = sections != null ? Convert.ToInt32(sections.Count) : 0;
                for (int i = 1; i <= sectionCount; i++)
                {
                    int firstSlide = Convert.ToInt32(sections.FirstSlide(i));
                    int slideCount = Convert.ToInt32(sections.SlidesCount(i));
                    if (slideCount > 0 && slideNumber >= firstSlide && slideNumber < firstSlide + slideCount)
                    {
                        string name = Convert.ToString(sections.Name(i));
                        return name ?? string.Empty;
                    }
                }
            }
            catch { }
            return string.Empty;
        }

        private static int CountBodyLines(string slideText)
        {
            if (string.IsNullOrWhiteSpace(slideText)) return 0;
            string[] lines = slideText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int contentLines = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.StartsWith("[Slide", StringComparison.OrdinalIgnoreCase) ||
                    line.StartsWith("[Speaker Notes]", StringComparison.OrdinalIgnoreCase))
                    continue;
                contentLines++;
            }
            // A title alone is not a substantive body.
            return contentLines > 1 ? contentLines - 1 : 0;
        }

        private dynamic AddSlideWithLayout(dynamic presentation, int insertIndex, string layoutName, dynamic sourceSlide)
        {
            if (presentation == null || presentation.Slides == null || string.IsNullOrWhiteSpace(layoutName)) return null;
            try
            {
                dynamic custom = FindCustomLayoutByName(presentation, layoutName);
                if (custom != null)
                {
                    try { return presentation.Slides.AddSlide(insertIndex, custom); }
                    catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddSlideWithLayout custom failed: {0}", ex.Message)); }
                }
                int enumVal = MapLayoutNameToEnum(layoutName);
                if (enumVal > 0)
                {
                    try { return presentation.Slides.Add(insertIndex, enumVal); }
                    catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddSlideWithLayout enum fallback failed: {0}", ex.Message)); }
                }
            }
            catch (Exception ex) { Logger.Warn(string.Format("PowerPointController.AddSlideWithLayout failed: {0}", ex.Message)); }
            return null;
        }

        private dynamic FindCustomLayoutByName(dynamic presentation, string layoutName)
        {
            if (presentation == null || string.IsNullOrWhiteSpace(layoutName)) return null;
            string key = layoutName.Trim().ToLowerInvariant();
            try
            {
                // Search all designs / masters
                dynamic designs = null;
                try { designs = presentation.Designs; } catch { }
                if (designs != null)
                {
                    int dc = 0;
                    try { dc = Convert.ToInt32(designs.Count); } catch { }
                    for (int d = 1; d <= dc; d++)
                    {
                        try
                        {
                            dynamic design = designs[d];
                            dynamic master = design != null ? design.SlideMaster : null;
                            dynamic layouts = master != null ? master.CustomLayouts : null;
                            int lc = layouts != null ? Convert.ToInt32(layouts.Count) : 0;
                            for (int i = 1; i <= lc; i++)
                            {
                                try
                                {
                                    dynamic l = layouts[i];
                                    string name = Convert.ToString(l.Name) ?? string.Empty;
                                    if (name.ToLowerInvariant().Contains(key) || key.Contains(name.ToLowerInvariant()))
                                        return l;
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                // Fallback to SlideMaster.CustomLayouts
                dynamic master2 = presentation.SlideMaster;
                dynamic layouts2 = master2 != null ? master2.CustomLayouts : null;
                int layoutCount2 = layouts2 != null ? Convert.ToInt32(layouts2.Count) : 0;
                for (int i = 1; i <= layoutCount2; i++)
                {
                    try
                    {
                        dynamic l = layouts2[i];
                        string name = Convert.ToString(l.Name) ?? string.Empty;
                        if (name.ToLowerInvariant().Contains(key) || key.Contains(name.ToLowerInvariant()))
                            return l;
                    }
                    catch { }
                }
            }
            catch { }
            return null;
        }

        private static int MapLayoutNameToEnum(string layoutName)
        {
            if (string.IsNullOrWhiteSpace(layoutName)) return 0;
            string k = layoutName.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
            switch (k)
            {
                case "title": return 1; // ppLayoutTitle
                case "text": return 2; // ppLayoutText
                case "titleandcontent": case "titlecontent": case "titlencontent": return 2;
                case "sectionheader": case "section": return 3;
                case "twocolumntext": case "twocolumn": case "twocontent": return 3;
                case "table": return 4;
                case "titleonly": case "onlytitle": return 11; // ppLayoutTitleOnly
                case "blank": return 12; // ppLayoutBlank
                case "comparison": return 3;
                case "titleslide": return 1;
                case "contentwithcaption": return 2;
                case "picturewithcaption": return 2;
                default: return 0;
            }
        }

        private static int MapChartTypeToXl(string chartType)
        {
            if (string.IsNullOrWhiteSpace(chartType)) return 51; // xlColumnClustered
            string k = chartType.Trim().ToLowerInvariant();
            switch (k)
            {
                case "column": case "columnclustered": return 51;
                case "line": return 65;
                case "bar": case "barclustered": return 57;
                case "pie": return 5;
                case "area": return 1;
                case "scatter": return -4169;
                case "doughnut": return -4120;
                default: return 51;
            }
        }

        private static int MapShapeTypeToMso(string shapeType)
        {
            if (string.IsNullOrWhiteSpace(shapeType)) return 1; // msoShapeRectangle
            string k = shapeType.Trim().ToLowerInvariant().Replace(" ", "").Replace("_", "").Replace("-", "");
            switch (k)
            {
                case "rectangle": return 1;
                case "roundedrectangle": return 5;
                case "oval": case "ellipse": case "circle": return 9;
                case "diamond": return 4;
                case "textbox": return 17; // msoShapeMixed/textbox sentinel - handled specially to use AddTextbox
                case "triangle": return 7;
                case "arrow": return 13;
                case "hexagon": return 10;
                case "star": return 12;
                default: return 1;
            }
        }

        private static int ParseColorToRgb(string color)
        {
            if (string.IsNullOrWhiteSpace(color)) return 0;
            string c = color.Trim().TrimStart('#');
            if (c.Length == 6)
            {
                int r = Convert.ToInt32(c.Substring(0, 2), 16);
                int g = Convert.ToInt32(c.Substring(2, 2), 16);
                int b = Convert.ToInt32(c.Substring(4, 2), 16);
                // Office RGB is BGR
                return (b << 16) | (g << 8) | r;
            }
            switch (c.ToLowerInvariant())
            {
                case "red": return 255;
                case "green": return 65280;
                case "blue": return 16711680;
                case "black": return 0;
                case "white": return 16777215;
                case "yellow": return 65535;
                default: return 0;
            }
        }

        private dynamic AddSlideUsingPresentationLayout(dynamic presentation, int insertIndex, dynamic sourceSlide)
        {
            if (presentation == null || presentation.Slides == null) return null;
            try
            {
                dynamic layout = null;
                try { layout = sourceSlide != null ? sourceSlide.CustomLayout : null; } catch { }
                if (layout != null)
                {
                    // AddSlide is a mutation, not a probe — a rejected custom layout must be logged,
                    // not swallowed, before falling back to the master's layouts (adversarial-review fix,
                    // same class of bug as D-14's PowerPoint/Excel bare-catch mutation swallows).
                    try { return presentation.Slides.AddSlide(insertIndex, layout); }
                    catch (Exception layoutEx) { Logger.Warn(string.Format("PowerPointController could not add slide with source slide's custom layout: {0}", layoutEx.Message)); }
                }

                try
                {
                    dynamic master = presentation.SlideMaster;
                    dynamic layouts = master != null ? master.CustomLayouts : null;
                    int layoutCount = layouts != null ? Convert.ToInt32(layouts.Count) : 0;
                    if (layoutCount > 0)
                    {
                        int layoutIndex = layoutCount >= 2 ? 2 : 1;
                        return presentation.Slides.AddSlide(insertIndex, layouts[layoutIndex]);
                    }
                }
                catch (Exception masterLayoutEx)
                {
                    Logger.Warn(string.Format("PowerPointController could not add slide with a master custom layout: {0}", masterLayoutEx.Message));
                }

                return presentation.Slides.Add(insertIndex, 2); // ppLayoutText
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.AddSlideUsingPresentationLayout failed: {0}", ex.Message));
                return null;
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

            // Keep model-proposed visuals in Notes so users can review them before using a local
            // image or an image-generation provider. This avoids silently inserting unreviewed media.
            string notesToApply = data.SpeakerNotes;
            if (!string.IsNullOrWhiteSpace(data.VisualSuggestion))
            {
                notesToApply = string.IsNullOrWhiteSpace(notesToApply)
                    ? string.Format("Visual suggestion: {0}", CleanMarkdown(data.VisualSuggestion))
                    : string.Format("{0}\rVisual suggestion: {1}", notesToApply, CleanMarkdown(data.VisualSuggestion));
            }

            // Also render visual suggestion as a visible placeholder shape/chart/table on the slide
            // so it is not silently dropped. We use heuristics to choose the most appropriate visual.
            if (!string.IsNullOrWhiteSpace(data.VisualSuggestion))
            {
                try
                {
                    string vs = CleanMarkdown(data.VisualSuggestion);
                    string vsLower = vs.ToLowerInvariant();
                    bool handled = false;
                    // Chart suggestion -> add a chart
                    if (!handled && (vsLower.Contains("chart") || vsLower.Contains("graph")))
                    {
                        string ctype = "column";
                        if (vsLower.Contains("pie")) ctype = "pie";
                        else if (vsLower.Contains("line")) ctype = "line";
                        else if (vsLower.Contains("bar")) ctype = "bar";
                        else if (vsLower.Contains("area")) ctype = "area";
                        else if (vsLower.Contains("scatter")) ctype = "scatter";
                        try
                        {
                            int xlType = MapChartTypeToXl(ctype);
                            dynamic chart = null;
                            try { chart = shapes.AddChart2(201, xlType, 70, 320, 500, 200); }
                            catch { chart = shapes.AddChart(xlType, 70, 320, 500, 200); }
                            if (chart != null && !string.IsNullOrWhiteSpace(vs))
                            {
                                try { chart.Chart.HasTitle = true; chart.Chart.ChartTitle.Text = vs.Length > 60 ? vs.Substring(0, 57) + "..." : vs; } catch { }
                                try { chart.AlternativeText = vs; } catch { }
                            }
                            handled = chart != null;
                        }
                        catch { }
                    }
                    // Table suggestion -> add a table
                    if (!handled && (vsLower.Contains("table") || vsLower.Contains("grid")))
                    {
                        try
                        {
                            dynamic tblShape = shapes.AddTable(2, 3, 60, 300, 540, 160);
                            if (tblShape != null && tblShape.Table != null)
                            {
                                try { tblShape.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = vs.Length > 40 ? vs.Substring(0, 37) + "..." : vs; } catch { }
                                try { tblShape.AlternativeText = vs; } catch { }
                            }
                            handled = tblShape != null;
                        }
                        catch { }
                    }
                    // Diagram / flow / SmartArt -> add diamond/rectangle shape
                    if (!handled && (vsLower.Contains("diagram") || vsLower.Contains("flow") || vsLower.Contains("smartart") || vsLower.Contains("process") || vsLower.Contains("cycle")))
                    {
                        try
                        {
                            dynamic dShape = shapes.AddShape(4, 80, 300, 500, 150); // diamond
                            try { dShape.TextFrame.TextRange.Text = vs; } catch { }
                            try { dShape.AlternativeText = vs; } catch { }
                            handled = dShape != null;
                        }
                        catch { }
                    }
                    // Generic image/visual placeholder -> add textbox with visual suggestion text
                    if (!handled)
                    {
                        try
                        {
                            dynamic ph = shapes.AddTextbox(1, 60, 320, 540, 100);
                            try { ph.TextFrame.TextRange.Text = string.Format("Visual: {0}", vs); } catch { }
                            try { ph.TextFrame.TextRange.Font.Italic = -1; } catch { }
                            try { ph.TextFrame.TextRange.Font.Size = 9; } catch { }
                            try { ph.AlternativeText = vs; } catch { }
                            try { ph.Line.Visible = -1; ph.Line.ForeColor.RGB = 8421504; } catch { }
                            try { ph.Fill.ForeColor.RGB = 15921906; ph.Fill.Visible = -1; ph.Fill.Transparency = 0.7f; } catch { }
                            handled = ph != null;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            // Set Speaker Notes if present
            if (!string.IsNullOrWhiteSpace(notesToApply))
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
                                    nShape.TextFrame.TextRange.Text = CleanMarkdown(notesToApply);
                                    break;
                                }
                            }
                        }
                    }
                }
                catch { }
            }
        }

        private static string CleanMarkdown(string input)
        {
            return PowerPointActionParser.CleanMarkdown(input);
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

                dynamic slide = GetOrCreateActiveSlide();
                PopulateSlide(slide, slideData);
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.AddBulletPoints failed", ex);
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

        public string GetContextReadout()
        {
            try
            {
                if (_rawAppObj == null) return string.Empty;

                dynamic app = _rawAppObj;
                dynamic presentation = null;
                try { presentation = app.ActivePresentation; } catch { }
                if (presentation == null) return string.Empty;

                dynamic slides = presentation.Slides;
                int totalSlideCount = 0;
                try { totalSlideCount = Convert.ToInt32(slides.Count); } catch { }
                if (totalSlideCount == 0) return string.Empty;

                dynamic activeSlide = GetActiveSlide();
                if (activeSlide == null) return string.Empty;

                int currentSlideNumber = 0;
                try { currentSlideNumber = Convert.ToInt32(activeSlide.SlideNumber); } catch { }
                if (currentSlideNumber == 0) return string.Empty;

                return string.Format("Slide {0} of {1}", currentSlideNumber, totalSlideCount);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetContextReadout failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        /// <summary>
        /// Navigate to a specific slide by 1-based slide number.
        /// Defensive: returns false on any error without throwing.
        /// </summary>
        public bool NavigateToSlide(int slideNumber)
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null)
                    return false;

                dynamic presentation = GetActivePresentation();
                if (presentation == null || presentation.Slides == null)
                    return false;

                int slideCount = 0;
                try { slideCount = Convert.ToInt32(presentation.Slides.Count); } catch { }

                if (slideNumber < 1 || slideNumber > slideCount)
                    return false;

                dynamic activeWin = null;
                try { activeWin = app.ActiveWindow; } catch { }
                if (activeWin == null || activeWin.View == null)
                    return false;

                // GotoSlide is the actual navigation — must not be silently swallowed, or a real
                // failure here (e.g. wrong view mode, protected view) would still report success.
                // Let it propagate to the outer catch, which correctly logs and returns false.
                activeWin.View.GotoSlide(slideNumber);

                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.NavigateToSlide failed: {0}", ex.Message));
                return false;
            }
        }
    }
}
