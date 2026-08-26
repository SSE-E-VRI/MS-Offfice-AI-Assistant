using System;
using System.Windows;
using System.Windows.Controls;
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
            TestChatSidebarToolbarsReflowAtNarrowWidth();
            TestSettingsWindowLoadsAndLaysOut();
        }

        private static void TestChatSidebarLoadsAndLaysOut()
        {
            var sidebar = new ChatSidebar();
            ForceLayout(sidebar, 360, 700);
        }

        /// <summary>
        /// Regression test: only ever laying out ChatSidebar at one fixed (comfortably wide) width
        /// let a real bug through — the header and toolbar rows were built as fixed-column Grids,
        /// where Auto columns never shrink below their content's natural width, so once total
        /// content exceeded the docked pane's actual width the rightmost controls (Settings/Help in
        /// the header, the "Edit" mode radio button in the toolbar) were silently clipped off-screen
        /// with no scrollbar or error. Fixed by switching both rows to WrapPanel, which reflows
        /// overflow onto a second line instead. This test lays out at a narrow width and confirms
        /// RdoEditMode (representative of "the control that was actually observed missing") stays
        /// fully within the panel's visible bounds — not just present in the tree, but positioned
        /// where a user could actually see and click it.
        /// </summary>
        private static void TestChatSidebarToolbarsReflowAtNarrowWidth()
        {
            const double narrowWidth = 240;
            var sidebar = new ChatSidebar();
            ForceLayout(sidebar, narrowWidth, 700);

            var editMode = sidebar.FindName("RdoEditMode") as FrameworkElement;
            if (editMode == null)
            {
                throw new Exception("RdoEditMode not found in the visual tree at narrow width.");
            }

            if (editMode.ActualWidth <= 0 || editMode.ActualHeight <= 0)
            {
                throw new Exception(string.Format(
                    "RdoEditMode has zero rendered size at {0}px width (ActualWidth={1}, ActualHeight={2}) — collapsed instead of reflowed.",
                    narrowWidth, editMode.ActualWidth, editMode.ActualHeight));
            }

            Point topLeft = editMode.TransformToVisual(sidebar).Transform(new Point(0, 0));
            double rightEdge = topLeft.X + editMode.ActualWidth;

            if (topLeft.X < 0 || rightEdge > narrowWidth)
            {
                throw new Exception(string.Format(
                    "RdoEditMode is positioned outside the {0}px visible width (left={1}, right={2}) — clipped, not wrapped onto a new line.",
                    narrowWidth, topLeft.X, rightEdge));
            }
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
