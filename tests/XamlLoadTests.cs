using System.Windows;
using MSOfficeAIAssistant.UI;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Constructs the two real WPF roots (ChatSidebar, SettingsWindow) and forces a layout pass.
    /// StaticResource keys are resolved at load time, not compile time — a missing/misscoped token
    /// compiles clean and only throws the moment InitializeComponent runs. Neither control needs
    /// Office to construct, so this stays inside the COM-free test project (§2.14) while closing
    /// that blind spot for the whole Theme/Tokens.xaml + Controls.xaml design system.
    /// </summary>
    public static class XamlLoadTests
    {
        public static void RunAll()
        {
            TestChatSidebarLoadsAndLaysOut();
            TestSettingsWindowLoadsAndLaysOut();
        }

        private static void TestChatSidebarLoadsAndLaysOut()
        {
            var sidebar = new ChatSidebar();
            ForceLayout(sidebar, 360, 700);
        }

        private static void TestSettingsWindowLoadsAndLaysOut()
        {
            var window = new SettingsWindow();
            ForceLayout(window, window.Width, window.Height);
        }

        private static void ForceLayout(FrameworkElement element, double width, double height)
        {
            element.Measure(new Size(width, height));
            element.Arrange(new Rect(0, 0, width, height));
            element.UpdateLayout();
        }
    }
}
