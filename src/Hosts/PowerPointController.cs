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

        private dynamic GetActivePresentation(bool createIfNone)
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
                    try { presentation = app.Presentations.Add(-1); } catch { }
                }
                return presentation;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetActivePresentation failed: {0}", ex.Message));
                return null;
            }
        }

        private dynamic GetOrCreateActiveSlide(bool createIfNone = false)
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
                dynamic pres = GetActivePresentation(createIfNone);

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
                dynamic presentation = GetActivePresentation(false);
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

        public string GetPresentationOutline(int maxCharacters)
        {
            if (maxCharacters <= 0) maxCharacters = 24000;
            try
            {
                dynamic presentation = GetActivePresentation(false);
                if (presentation == null || presentation.Slides == null) return string.Empty;

                var sb = new StringBuilder();
                int count = Convert.ToInt32(presentation.Slides.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic slide = presentation.Slides[i];
                    string title = GetSlideTitle(slide);
                    if (string.IsNullOrWhiteSpace(title)) title = "(untitled)";
                    string sectionName = GetSectionName(presentation, i);
                    string prefix = string.IsNullOrWhiteSpace(sectionName) ? string.Empty : string.Format(" [{0}]", sectionName);
                    AppendBounded(sb, string.Format("{0}. {1}{2}\n", i, title, prefix), maxCharacters);
                    if (sb.Length >= maxCharacters) break;
                }
                return sb.ToString().TrimEnd();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetPresentationOutline failed: {0}", ex.Message));
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
                dynamic presentation = GetActivePresentation(false);
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
                PowerPointActionParser.ParseStructuredActions(text, out cleanedText);
                text = cleanedText;
                if (string.IsNullOrWhiteSpace(text)) return false;

                dynamic app = _rawAppObj;
                if (app == null) return false;

                // 1. If in normal/slide view and user has actively selected text, insert directly
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
                                selection.TextRange.InsertAfter(CleanMarkdown(text));
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

                dynamic pres = GetActivePresentation(true);
                if (pres == null) return false;

                dynamic activeSlide = GetOrCreateActiveSlide(true);
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

        /// <summary>
        /// Builds a deck from the project slide format. If the active slide has content, new slides
        /// are appended. If the active slide is blank, it is populated first.
        /// </summary>
        public void CreateOrUpdateDeckFromOutline(string outline)
        {
            InsertText(outline);
        }

        /// <summary>
        /// Applies a single structured deck action returned by the model, setting status and result/error on the action object.
        /// </summary>
        public bool ApplyPowerPointAction(PowerPointAction action)
        {
            if (action == null) return false;
            action.Status = PowerPointActionStatus.Applying;

            try
            {
                bool success = false;
                string result = null;

                if (string.Equals(action.Type, "move_slide", StringComparison.OrdinalIgnoreCase))
                {
                    success = MoveSlide(action.Source, action.Target);
                    if (success) result = string.Format("Moved slide {0} to {1}", action.Source, action.Target);
                }
                else if (string.Equals(action.Type, "create_section", StringComparison.OrdinalIgnoreCase))
                {
                    int targetSlide = action.Slide > 0 ? action.Slide : action.Target;
                    success = CreateSectionBeforeSlide(action.Name, targetSlide);
                    if (success) result = string.Format("Created section '{0}'", action.Name);
                }
                else if (string.Equals(action.Type, "rename_section", StringComparison.OrdinalIgnoreCase))
                {
                    success = RenameSection(action.Section, action.Name);
                    if (success) result = string.Format("Renamed section {0} to '{1}'", action.Section, action.Name);
                }
                else if (string.Equals(action.Type, "set_notes", StringComparison.OrdinalIgnoreCase))
                {
                    success = SetSpeakerNotesForSlide(action.Slide, action.Notes);
                    if (success) result = string.Format("Set speaker notes on slide {0}", action.Slide);
                }
                else
                {
                    action.Status = PowerPointActionStatus.Error;
                    action.ErrorMessage = string.Format("Unsupported action type '{0}'", action.Type);
                    return false;
                }

                if (success)
                {
                    action.ResultText = result;
                    action.Status = PowerPointActionStatus.Applied;
                    Logger.Info(string.Format("Applied PowerPoint action {0}: {1}", action.Type, result));
                    return true;
                }
                else
                {
                    action.Status = PowerPointActionStatus.Error;
                    action.ErrorMessage = string.Format("Failed to execute {0}", action.Type);
                    return false;
                }
            }
            catch (Exception ex)
            {
                action.Status = PowerPointActionStatus.Error;
                action.ErrorMessage = ex.Message;
                Logger.Error(string.Format("ApplyPowerPointAction failed on {0}", action.Type), ex);
                return false;
            }
        }

        public HostOperationResult ExecutePowerPointAction(PowerPointAction action)
        {
            if (action == null)
                return HostOperationResult.Failed("PowerPoint action cannot be null.");

            if (!ApplyPowerPointAction(action))
            {
                return HostOperationResult.Failed(
                    !string.IsNullOrEmpty(action.ErrorMessage) ? action.ErrorMessage : "PowerPoint action application failed.",
                    0,
                    action.Slide > 0 ? "Slide " + action.Slide : null);
            }

            return HostOperationResult.Ok(action.ResultText, action.Slide > 0 ? "Slide " + action.Slide : null);
        }

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

        public HostOperationResult ExecuteInsertImage(string filePath, string altText = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return HostOperationResult.Failed("Image file path cannot be empty.");

            if (!System.IO.File.Exists(filePath))
                return HostOperationResult.Failed(string.Format("Image file not found: {0}", filePath));

            try
            {
                bool ok = InsertImageFromFile(filePath, altText);
                if (ok)
                    return HostOperationResult.Ok(string.Format("Inserted image from {0}", System.IO.Path.GetFileName(filePath)));
                else
                    return HostOperationResult.Failed(string.Format("Failed to insert image from {0}", filePath));
            }
            catch (Exception ex)
            {
                Logger.Error("PowerPointController.ExecuteInsertImage failed", ex);
                return HostOperationResult.FromException(ex, "PowerPointController.ExecuteInsertImage");
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

        public virtual bool MoveSlide(int sourceSlideNumber, int destinationSlideNumber)
        {
            dynamic presentation = GetActivePresentation(false);
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
            dynamic presentation = GetActivePresentation(false);
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
            dynamic presentation = GetActivePresentation(false);
            if (presentation == null || presentation.SectionProperties == null) return false;
            presentation.SectionProperties.Rename(sectionIndex, sectionName.Trim());
            return true;
        }

        public virtual bool SetSpeakerNotesForSlide(int slideNumber, string notes)
        {
            if (slideNumber < 1 || string.IsNullOrWhiteSpace(notes)) return false;
            dynamic presentation = GetActivePresentation(false);
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
                dynamic presentation = GetActivePresentation(false);
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
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return false;
            try
            {
                dynamic slide = GetOrCreateActiveSlide(true);
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
                Logger.Warn(string.Format("PowerPointController.InsertImageFromFile failed: {0}", ex.Message));
                return false;
            }
        }

        public bool UndoLastChange()
        {
            try
            {
                dynamic app = _rawAppObj;
                if (app == null || app.CommandBars == null) return false;
                app.CommandBars.ExecuteMso("Undo");
                return true;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.UndoLastChange failed: {0}", ex.Message));
                return false;
            }
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

        private dynamic AddSlideUsingPresentationLayout(dynamic presentation, int insertIndex, dynamic sourceSlide)
        {
            if (presentation == null || presentation.Slides == null) return null;
            try
            {
                dynamic layout = null;
                try { layout = sourceSlide != null ? sourceSlide.CustomLayout : null; } catch { }
                if (layout != null)
                {
                    try { return presentation.Slides.AddSlide(insertIndex, layout); } catch { }
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
                catch { }

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
