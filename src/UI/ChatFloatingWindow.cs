using System;
using System.Windows;
using MistralOfficeAddin.Core;

namespace MistralOfficeAddin.UI
{
    /// <summary>
    /// Fallback / Standalone WPF Window hosting ChatSidebar when Office Custom Task Pane factory
    /// is not provided by the host application (e.g. during external automation or version quirks).
    /// </summary>
    public class ChatFloatingWindow : Window
    {
        private readonly ChatSidebar _sidebar;

        public ChatSidebar Sidebar
        {
            get { return _sidebar; }
        }

        public ChatFloatingWindow()
        {
            this.Title = "AI Assistant";

            this.Width = 420;
            this.Height = 720;
            this.MinWidth = 340;
            this.MinHeight = 450;
            this.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            this.Background = System.Windows.Media.Brushes.White;

            _sidebar = new ChatSidebar();
            this.Content = _sidebar;
        }

        public void InitializeHost(object appObj, string hostType)
        {
            try
            {
                if (_sidebar != null)
                {
                    _sidebar.InitializeHost(appObj, hostType);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatFloatingWindow.InitializeHost error: {0}", ex.Message));
            }
        }

        public void ExecutePrompt(string prompt, string title)
        {
            try
            {
                if (_sidebar != null)
                {
                    _sidebar.ExecuteExternalPrompt(prompt, title);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatFloatingWindow.ExecutePrompt error: {0}", ex.Message));
            }
        }

        public void StartNewChat()
        {
            try
            {
                if (_sidebar != null)
                {
                    _sidebar.StartNewChat(true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatFloatingWindow.StartNewChat error: {0}", ex.Message));
            }
        }
    }
}
