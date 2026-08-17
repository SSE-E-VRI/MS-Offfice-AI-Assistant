using System;
using System.Windows.Forms;
using MistralOfficeAddin.Core;
using MistralOfficeAddin.UI;

namespace MistralOfficeAddin.Addin
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
                string targetLang = "English";
                try
                {
                    dynamic ctrl = control;
                    string id = Convert.ToString(ctrl.Id);
                    if (id == "btnTransES") targetLang = "Spanish";
                    else if (id == "btnTransFR") targetLang = "French";
                    else if (id == "btnTransDE") targetLang = "German";
                    else if (id == "btnTransZH") targetLang = "Chinese";
                }
                catch { }

                Logger.Info(string.Format("Ribbon Callback: OnTranslate ({0})", targetLang));
                _taskPaneManager.ExecutePrompt(string.Format("Translate the following text accurately into {0}, preserving tone and nuance.", targetLang), string.Format("Translate ({0})", targetLang));
            }
            catch (Exception ex)
            {
                Logger.Error("OnTranslate error", ex);
            }
        }

        public void OnOpenSettings(object control)
        {
            try
            {
                Logger.Info("Ribbon Callback: OnOpenSettings");
                var win = new SettingsWindow();
                win.ShowDialog();
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
