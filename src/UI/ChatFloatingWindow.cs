using System;
using System.Windows;
using System.Windows.Forms.Integration;
using System.Windows.Interop;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI
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
            this.ShowActivated = true;

            _sidebar = new ChatSidebar();
            this.Content = _sidebar;

            try
            {
                // Excel 2010 owns the native message loop. This bridge is required for a
                // modeless WPF window to receive keyboard input while Excel remains active.
                ElementHost.EnableModelessKeyboardInterop(this);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatFloatingWindow keyboard interop setup failed: {0}", ex.Message));
            }
        }

        public void InitializeHost(object appObj, string hostType)
        {
            try
            {
                if (_sidebar != null)
                {
                    _sidebar.InitializeHost(appObj, hostType);
                }

                IntPtr hostWindow = OfficeDockedPane.ResolveHostHwnd(appObj);
                if (hostWindow != IntPtr.Zero)
                {
                    new WindowInteropHelper(this).Owner = hostWindow;
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

        public void FocusPromptInput()
        {
            try
            {
                if (_sidebar != null) _sidebar.FocusPromptInput();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ChatFloatingWindow.FocusPromptInput error: {0}", ex.Message));
            }
        }
    }
}
