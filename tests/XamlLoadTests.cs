using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using MSOfficeAIAssistant.Core;
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
            TestActionHistoryWindowWrapsLongEntries();
            TestResponsePreviewWindowShowsCleanedContent();
            TestDocumentCompareWindowLoadsAndLaysOut();
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

        /// <summary>
        /// Regression test for the Action Log rendering every entry as one endless line with a
        /// horizontal scrollbar. The card's TextBlocks always had TextWrapping="Wrap", but a
        /// ListBox defaults to ScrollViewer.HorizontalScrollBarVisibility=Auto, which measures
        /// each item at its full desired width instead of the viewport's -- so wrapping never
        /// engaged and the Undoable column was pushed off-screen. Fixed by disabling horizontal
        /// scrolling and stretching the item containers; this asserts the container really is
        /// bounded by the list's own width.
        /// </summary>
        private static void TestActionHistoryWindowWrapsLongEntries()
        {
            const double width = 900;
            const double height = 600;

            string longEntry = "**[Your Department's Letterhead]** " +
                new string('x', 400) +
                " I am writing on behalf of the Electrical Division to request one unit of a 5 HP pump.";

            var entries = new List<ActionAuditEntry>
            {
                new ActionAuditEntry
                {
                    TimestampUtc = DateTime.UtcNow,
                    Host = "Word",
                    ActionType = "Tracked edit",
                    Target = "Selection / cursor",
                    Undoable = true,
                    Model = "mistral-large-latest",
                    Summary = longEntry
                }
            };

            var window = new ActionHistoryWindow(entries);
            ForceLayout(window, width, height);

            var list = window.FindName("ActionsListBox") as ListBox;
            if (list == null) throw new Exception("ActionsListBox not found in the visual tree.");

            if (ScrollViewer.GetHorizontalScrollBarVisibility(list) != ScrollBarVisibility.Disabled)
            {
                throw new Exception("ActionsListBox must disable horizontal scrolling, otherwise entries never wrap.");
            }

            var container = list.ItemContainerGenerator.ContainerFromIndex(0) as ListBoxItem;
            if (container == null)
            {
                var generator = list.ItemContainerGenerator as System.Windows.Controls.Primitives.IItemContainerGenerator;
                if (generator != null)
                {
                    var pos = generator.GeneratorPositionFromIndex(-1);
                    using (generator.StartAt(pos, System.Windows.Controls.Primitives.GeneratorDirection.Forward, true))
                    {
                        bool isNewlyRealized;
                        container = generator.GenerateNext(out isNewlyRealized) as ListBoxItem;
                        if (container != null)
                        {
                            generator.PrepareItemContainer(container);
                            ForceLayout(container, list.ActualWidth > 0 ? list.ActualWidth : width, height);
                        }
                    }
                }
            }
            if (container == null) throw new Exception("No container was generated for the first audit entry.");

            double maxAllowedWidth = list.ActualWidth > 0 ? list.ActualWidth : width;
            if (container.ActualWidth > maxAllowedWidth + 1.0)
            {
                throw new Exception(string.Format(
                    "Audit entry is {0:0} wide inside a {1:0}-wide list -- it is not wrapping to the viewport.",
                    container.ActualWidth, maxAllowedWidth));
            }
        }

        /// <summary>
        /// The Preview button used to call MessageBox.Show on the raw response, so it displayed
        /// literal "**" markers and the model's trailing commentary ("Key Features", "Structure")
        /// that Insert does not apply -- a preview that did not match the result. The window is
        /// built entirely in code, so this both proves it constructs and lays out, and asserts it
        /// previews the same cleaned content Insert writes while keeping the raw text available.
        /// </summary>
        private static void TestResponsePreviewWindowShowsCleanedContent()
        {
            string response =
                "Here's a polished draft for your request letter:\n\n---\n\n" +
                "**SOUTHERN RAILWAY**\n\nSir,\n\nRequest for temporary allocation of one 5 HP pump.\n\n" +
                "**Yours faithfully,**\n[Your Full Name]\n\n---\n\n" +
                "### **Key Features:**\n1. **Formal numbering** for clarity.\n2. **Structure:** reference line, numbered paragraphs.";

            var window = new ResponsePreviewWindow(response);
            ForceLayout(window, window.Width, window.Height);

            if (window.PreviewText.Contains("polished draft"))
            {
                throw new Exception("Preview still shows the model's lead-in: " + window.PreviewText);
            }
            if (window.PreviewText.Contains("Key Features"))
            {
                throw new Exception("Preview still shows the trailing commentary Insert discards.");
            }
            if (!window.PreviewText.Contains("SOUTHERN RAILWAY") || !window.PreviewText.Contains("[Your Full Name]"))
            {
                throw new Exception("Preview dropped part of the letter body: " + window.PreviewText);
            }
            if (window.RawText != response)
            {
                throw new Exception("The raw response must stay available unchanged for the toggle.");
            }
        }

        /// <summary>
        /// DocumentCompareWindow (Slice 5) merges the same Theme/Tokens.xaml + Controls.xaml
        /// dictionaries as every other window here, via the same StaticResource-resolved-at-load
        /// mechanism this file exists to test -- it was the one new window this cycle that had no
        /// construction test, so a bad/missing resource key would have compiled clean and only
        /// thrown the first time a user clicked "Compare Docs" in a live host.
        /// </summary>
        private static void TestDocumentCompareWindowLoadsAndLaysOut()
        {
            var window = new DocumentCompareWindow();
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
