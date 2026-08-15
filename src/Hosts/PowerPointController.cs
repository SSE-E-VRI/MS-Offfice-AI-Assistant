using System;
using System.Collections.Generic;
using System.Text;
using MistralOfficeAddin.Core;
using PPT = NetOffice.PowerPointApi;

namespace MistralOfficeAddin.Hosts
{
    public class PowerPointController
    {
        private readonly object _rawAppObj;
        private PPT.Application _pptApp;

        public PowerPointController(object appObj)
        {
            _rawAppObj = appObj;
        }

        private PPT.Application GetApp()
        {
            if (_pptApp != null) return _pptApp;
            if (_rawAppObj == null) return null;
            try
            {
                _pptApp = (_rawAppObj is PPT.Application)
                    ? (PPT.Application)_rawAppObj
                    : new PPT.Application(null, _rawAppObj);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetApp failed: {0}", ex.Message));
            }
            return _pptApp;
        }

        public string GetSlideText()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveWindow != null && app.ActiveWindow.View != null)
                {
                    if (app.ActiveWindow.View.Slide is PPT.Slide)
                    {
                        var slide = (PPT.Slide)app.ActiveWindow.View.Slide;
                        var sb = new StringBuilder();
                        sb.AppendLine(string.Format("[Slide #{0}: {1}]", slide.SlideNumber, slide.Name));
                        foreach (PPT.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTextFrame == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue && shape.TextFrame.HasText == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue)
                            {
                                string text = shape.TextFrame.TextRange.Text;
                                if (!string.IsNullOrWhiteSpace(text))
                                    sb.AppendLine(text.Trim());
                            }
                        }
                        return sb.ToString().TrimEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("PowerPointController.GetSlideText failed: {0}", ex.Message));
            }
            return string.Empty;
        }

        public void AddBulletPoints(List<string> bullets)
        {
            if (bullets == null || bullets.Count == 0) return;
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveWindow != null && app.ActiveWindow.View != null)
                {
                    if (app.ActiveWindow.View.Slide is PPT.Slide)
                    {
                        var slide = (PPT.Slide)app.ActiveWindow.View.Slide;
                        PPT.Shape targetShape = null;
                        foreach (PPT.Shape shape in slide.Shapes)
                        {
                            if (shape.HasTextFrame == NetOffice.OfficeApi.Enums.MsoTriState.msoTrue)
                            { targetShape = shape; break; }
                        }
                        if (targetShape == null)
                        {
                            targetShape = slide.Shapes.AddTextbox(
                                NetOffice.OfficeApi.Enums.MsoTextOrientation.msoTextOrientationHorizontal,
                                100, 100, 500, 300);
                        }
                        var textRange = targetShape.TextFrame.TextRange;
                        foreach (var bullet in bullets)
                        {
                            if (string.IsNullOrWhiteSpace(bullet)) continue;
                            var para = textRange.InsertAfter("\r" + bullet.Trim());
                            para.ParagraphFormat.Bullet.Type = PPT.Enums.PpBulletType.ppBulletUnnumbered;
                        }
                    }
                }
            }
            catch (Exception ex) { Logger.Error("PowerPointController.AddBulletPoints failed", ex); throw; }
        }

        public void SetSpeakerNotes(string notes)
        {
            if (notes == null) notes = string.Empty;
            try
            {
                var app = GetApp();
                if (app != null && app.ActiveWindow != null && app.ActiveWindow.View != null)
                {
                    if (app.ActiveWindow.View.Slide is PPT.Slide)
                    {
                        var slide = (PPT.Slide)app.ActiveWindow.View.Slide;
                        if (slide.NotesPage != null && slide.NotesPage.Shapes != null)
                        {
                            foreach (PPT.Shape shape in slide.NotesPage.Shapes)
                            {
                                if (shape.PlaceholderFormat != null &&
                                    shape.PlaceholderFormat.Type == PPT.Enums.PpPlaceholderType.ppPlaceholderBody)
                                {
                                    shape.TextFrame.TextRange.Text = notes;
                                    return;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Logger.Error("PowerPointController.SetSpeakerNotes failed", ex); throw; }
        }

        public string GetActivePresentationName()
        {
            try
            {
                var app = GetApp();
                if (app != null && app.ActivePresentation != null)
                    return app.ActivePresentation.Name;
            }
            catch { }
            return "PowerPointPresentation";
        }
    }
}
