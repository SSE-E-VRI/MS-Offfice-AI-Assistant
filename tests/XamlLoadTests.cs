using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using MSOfficeAIAssistant.Core.QuickPrompts;
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
            TestChatSidebarSendButtonStaysVisibleAtNarrowWidth();
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

            AssertWithinVisibleBounds(sidebar, editMode, narrowWidth, "RdoEditMode");
        }

        /// <summary>
        /// Regression test: the same clipping bug class, found separately in the input area's
        /// Attach/quick-prompt-chips/Send row. That row nested a WrapPanel (Attach + chips + Insert
        /// image) inside a Grid column with Width="Auto" — a Grid measures an Auto column's child
        /// with effectively unlimited available width, so the WrapPanel never actually got a chance
        /// to wrap; it always requested its full unwrapped width. With BtnSend in the Grid's other
        /// Auto column, Send was the one silently pushed past the pane's visible edge once the
        /// combined content overflowed. Fixed by splitting onto two lines (chips WrapPanel on top,
        /// Send on its own guaranteed-visible line below) instead of one Grid row.
        ///
        /// A first version of this test passed even against the broken Grid layout — a freshly
        /// constructed ChatSidebar never populates QuickPromptsItemsControl (that only happens
        /// inside InitializeHostOnUiThread, which needs a live host and never runs in this headless
        /// test), so the row's content was far too sparse to actually trigger the overflow a real
        /// session with 5-8 populated chips hits. Manually populating the ItemsSource with
        /// representative sample chips before layout is what makes this test meaningful.
        /// </summary>
        private static void TestChatSidebarSendButtonStaysVisibleAtNarrowWidth()
        {
            const double narrowWidth = 240;
            var sidebar = new ChatSidebar();

            var quickPrompts = sidebar.FindName("QuickPromptsItemsControl") as ItemsControl;
            if (quickPrompts == null)
            {
                throw new Exception("QuickPromptsItemsControl not found in the visual tree.");
            }
            quickPrompts.ItemsSource = new List<QuickPrompt>
            {
                new QuickPrompt { Id = "Summarize", Label = "Summarize", PromptText = "x" },
                new QuickPrompt { Id = "Rewrite", Label = "Rewrite", PromptText = "x" },
                new QuickPrompt { Id = "Outline", Label = "Outline", PromptText = "x" },
                new QuickPrompt { Id = "Actions", Label = "Actions", PromptText = "x" },
                new QuickPrompt { Id = "Review", Label = "Review", PromptText = "x" },
                new QuickPrompt { Id = "Deck", Label = "Build deck", PromptText = "x" }
            };

            ForceLayout(sidebar, narrowWidth, 700);

            var sendButton = sidebar.FindName("BtnSend") as FrameworkElement;
            if (sendButton == null)
            {
                throw new Exception("BtnSend not found in the visual tree at narrow width.");
            }

            AssertWithinVisibleBounds(sidebar, sendButton, narrowWidth, "BtnSend");
        }

        private static void AssertWithinVisibleBounds(FrameworkElement container, FrameworkElement element, double containerWidth, string elementName)
        {
            if (element.ActualWidth <= 0 || element.ActualHeight <= 0)
            {
                throw new Exception(string.Format(
                    "{0} has zero rendered size at {1}px width (ActualWidth={2}, ActualHeight={3}) — collapsed instead of reflowed.",
                    elementName, containerWidth, element.ActualWidth, element.ActualHeight));
            }

            Point topLeft = element.TransformToVisual(container).Transform(new Point(0, 0));
            double rightEdge = topLeft.X + element.ActualWidth;

            if (topLeft.X < 0 || rightEdge > containerWidth)
            {
                throw new Exception(string.Format(
                    "{0} is positioned outside the {1}px visible width (left={2}, right={3}) — clipped, not wrapped onto a new line.",
                    elementName, containerWidth, topLeft.X, rightEdge));
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
