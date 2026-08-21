using System;
using System.Windows.Forms;
using MSOfficeAIAssistant.Core;
using MSOfficeAIAssistant.UI;

namespace MSOfficeAIAssistant.Addin
{
    public class RibbonCallback
    {
        private readonly CustomTaskPaneManager _taskPaneManager;

        public RibbonCallback(CustomTaskPaneManager taskPaneManager)
        {
            _taskPaneManager = taskPaneManager;
        }

        public void OnToggleSidebar(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnToggleSidebar");
                _taskPaneManager.TogglePane();
            }
            catch (Exception ex)
            {
                Logger.Error("OnToggleSidebar error", ex);
            }
        }

        public void OnNewChat(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnNewChat");
                _taskPaneManager.StartNewChat();
            }
            catch (Exception ex)
            {
                Logger.Error("OnNewChat error", ex);
            }
        }

        public void OnGenerate(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnGenerate");
                _taskPaneManager.ExecutePrompt("Generate a comprehensive draft based on the topic or outline provided.", "Generate Draft");
            }
            catch (Exception ex)
            {
                Logger.Error("OnGenerate error", ex);
            }
        }

        public void OnContinueWriting(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnContinueWriting");
                _taskPaneManager.ExecutePrompt("Continue writing seamlessly from the current point in the text.", "Continue Writing");
            }
            catch (Exception ex)
            {
                Logger.Error("OnContinueWriting error", ex);
            }
        }

        public void OnSummarize(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnSummarize");
                _taskPaneManager.ExecutePrompt("Provide a concise executive summary highlighting key takeaways and action items.", "Summarize");
            }
            catch (Exception ex)
            {
                Logger.Error("OnSummarize error", ex);
            }
        }

        public void OnRewrite(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnRewrite");
                _taskPaneManager.ExecutePrompt("Rewrite the selected text for maximum clarity, professional flow, and polished vocabulary.", "Rewrite");
            }
            catch (Exception ex)
            {
                Logger.Error("OnRewrite error", ex);
            }
        }

        public void OnExpand(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnExpand");
                _taskPaneManager.ExecutePrompt("Elaborate on the selected text with supporting details, explanations, and context.", "Expand");
            }
            catch (Exception ex)
            {
                Logger.Error("OnExpand error", ex);
            }
        }

        public void OnShorten(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnShorten");
                _taskPaneManager.ExecutePrompt("Condense the selected text into a tight, impactful version without losing core meaning.", "Shorten");
            }
            catch (Exception ex)
            {
                Logger.Error("OnShorten error", ex);
            }
        }

        public void OnTranslate(object control)
        {
            try
            {
                string targetLang = "Tamil";
                try
                {
                    dynamic ctrl = control;
                    string id = Convert.ToString(ctrl.Id);
                    if (id == "btnTransTA") targetLang = "Tamil";
                    else if (id == "btnTransHI") targetLang = "Hindi";
                    else if (id == "btnTransTE") targetLang = "Telugu";
                    else if (id == "btnTransKN") targetLang = "Kannada";
                    else if (id == "btnTransML") targetLang = "Malayalam";
                    else if (id == "btnTransBN") targetLang = "Bengali";
                    else if (id == "btnTransMR") targetLang = "Marathi";
                    else if (id == "btnTransGU") targetLang = "Gujarati";
                    else if (id == "btnTransEN") targetLang = "English";
                }
                catch { }

                Logger.Info(string.Format("Ribbon Callback: OnTranslate ({0})", targetLang));
                _taskPaneManager.ExecutePrompt(
                    string.Format("Translate the following text into {0} accurately and naturally. Output ONLY the clean translated text in {0}, with no introductory remarks, conversation, or explanation.", targetLang),
                    string.Format("Translate ({0})", targetLang));
            }
            catch (Exception ex)
            {
                Logger.Error("OnTranslate error", ex);
            }
        }

        public void OnOutline(object control)
        {
            try
            {
                _taskPaneManager.ExecutePrompt(
                    "Create a clear, hierarchical outline of the supplied content. Preserve key facts and show the recommended narrative flow.",
                    "Outline");
            }
            catch (Exception ex)
            {
                Logger.Error("OnOutline error", ex);
            }
        }

        public void OnActionItems(object control)
        {
            try
            {
                _taskPaneManager.ExecutePrompt(
                    "Extract decisions, action items, owners where explicitly stated, deadlines where explicitly stated, and risks. Do not invent people or dates.",
                    "Action Items");
            }
            catch (Exception ex)
            {
                Logger.Error("OnActionItems error", ex);
            }
        }

        public void OnReviewContent(object control)
        {
            try
            {
                _taskPaneManager.ExecutePrompt(
                    "Review the supplied content for clarity, gaps, consistency, duplicated ideas, and the most valuable next edits. For presentations, assess story flow and weak slides.",
                    "Review Content");
            }
            catch (Exception ex)
            {
                Logger.Error("OnReviewContent error", ex);
            }
        }

        public void OnBuildSlides(object control)
        {
            try
            {
                _taskPaneManager.ExecutePrompt(
                    "Create a concise, coherent slide deck from the supplied content. Return numbered slides with a title, 3-5 concise bullets, a Visual suggestion, and Speaker Notes for each slide.",
                    "Build Slides");
            }
            catch (Exception ex)
            {
                Logger.Error("OnBuildSlides error", ex);
            }
        }

        public void OnOpenUserManual(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnOpenUserManual");
                _taskPaneManager.OpenUserManual();
            }
            catch (Exception ex)
            {
                Logger.Error("OnOpenUserManual error", ex);
            }
        }

        public void OnOpenSettings(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnOpenSettings");
                var win = new SettingsWindow();
                if (win.ShowDialog() == true && _taskPaneManager != null)
                {
                    _taskPaneManager.ReloadConfiguredProvider();
                }
            }
            catch (Exception ex)
            {
                Logger.Error("OnOpenSettings error", ex);
                MessageBox.Show(string.Format("Could not open settings: {0}", ex.Message), "AI Assistant", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
