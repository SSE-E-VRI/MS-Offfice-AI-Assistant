using System;
using System.Globalization;
using System.Windows;
using MSOfficeAIAssistant.UI.Converters;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// The Action Log stores the raw response that was applied, so it was rendering Markdown
    /// markers and pipe-tables verbatim, and dumping the whole 2000-character entry into the card.
    /// These cover the converters that turn a stored entry into something readable.
    /// </summary>
    public static class AuditDisplayConverterTests
    {
        private const string SampleEntry =
            "**[Your Department's Letterhead]**\n[Date]\n\n**To,**\nThe Senior Section Engineer (Works)\n\n" +
            "**Subject:** Request for Temporary Provision of a 5 HP Submersible Pump\n\n" +
            "Dear Sir,\n\nI am writing to request **one (1) unit of a 5 HP submersible pump**.\n\n" +
            "**Key Details of the Request:**\n| **Parameter** | **Specification** |\n|---|---|\n" +
            "| **Capacity** | 5 HP |\n\n**Yours faithfully,**\n**[Your Full Name]**";

        public static void RunAll()
        {
            TestPlainTextConverterStripsMarkers();
            TestPlainTextConverterCollapsesBlankRuns();
            TestPreviewConverterShortensAndEllipsizes();
            TestPreviewConverterLeavesShortTextAlone();
            TestLongTextVisibility();
            TestUndoableLabel();
            TestNullsAreSafe();
        }

        private static void TestPlainTextConverterStripsMarkers()
        {
            var c = new MarkdownPlainTextConverter();
            string plain = (string)c.Convert(SampleEntry, typeof(string), null, CultureInfo.InvariantCulture);

            Assert(!plain.Contains("**"), "Bold markers must be gone: " + plain);
            Assert(plain.Contains("Subject: Request for Temporary Provision"), "Text must survive: " + plain);
            Assert(plain.Contains("[Your Full Name]"), "Placeholders must survive");
        }

        private static void TestPlainTextConverterCollapsesBlankRuns()
        {
            var c = new MarkdownPlainTextConverter();
            string plain = (string)c.Convert("A\n\n\n\n\nB", typeof(string), null, CultureInfo.InvariantCulture);
            Assert(plain.Replace("\r\n", "\n") == "A\n\nB", "Blank runs must collapse to one: " + plain.Replace("\n", "\\n"));
        }

        private static void TestPreviewConverterShortensAndEllipsizes()
        {
            var c = new PreviewTextConverter();
            string preview = (string)c.Convert(SampleEntry, typeof(string), null, CultureInfo.InvariantCulture);

            Assert(!preview.Contains("**"), "Preview must be marker-free: " + preview);
            Assert(preview.Length <= 221 + 1, "Preview must be short, was " + preview.Length);
            Assert(preview.EndsWith("\u2026"), "Clipped preview must end with an ellipsis: " + preview);
            Assert(preview.StartsWith("[Your Department's Letterhead]"), "Preview starts at the first real line: " + preview);
            Assert(!preview.Contains("Yours faithfully"), "Preview must not reach the end of the entry");
        }

        private static void TestPreviewConverterLeavesShortTextAlone()
        {
            var c = new PreviewTextConverter();
            string preview = (string)c.Convert("Inserted at cursor", typeof(string), null, CultureInfo.InvariantCulture);
            Assert(preview == "Inserted at cursor", "Short entries must pass through unchanged: " + preview);
        }

        private static void TestLongTextVisibility()
        {
            var c = new LongTextToVisibilityConverter();
            Assert((Visibility)c.Convert(SampleEntry, typeof(Visibility), null, CultureInfo.InvariantCulture) == Visibility.Visible,
                "A long multi-line entry needs the expander");
            Assert((Visibility)c.Convert("Inserted at cursor", typeof(Visibility), null, CultureInfo.InvariantCulture) == Visibility.Collapsed,
                "A one-line entry must not show an empty expander");
        }

        private static void TestUndoableLabel()
        {
            var c = new UndoableTextConverter();
            Assert((string)c.Convert(true, typeof(string), null, CultureInfo.InvariantCulture) == "Undoable", "true -> Undoable");
            Assert((string)c.Convert(false, typeof(string), null, CultureInfo.InvariantCulture) == "Not undoable", "false -> Not undoable");
            Assert((string)c.Convert(null, typeof(string), null, CultureInfo.InvariantCulture) == "Not undoable", "null -> Not undoable");
        }

        private static void TestNullsAreSafe()
        {
            var plain = new MarkdownPlainTextConverter();
            var preview = new PreviewTextConverter();
            var visible = new LongTextToVisibilityConverter();

            Assert((string)plain.Convert(null, typeof(string), null, CultureInfo.InvariantCulture) == string.Empty, "null plain");
            Assert((string)preview.Convert(null, typeof(string), null, CultureInfo.InvariantCulture) == string.Empty, "null preview");
            Assert((Visibility)visible.Convert(null, typeof(Visibility), null, CultureInfo.InvariantCulture) == Visibility.Collapsed, "null visibility");
        }

        private static void Assert(bool condition, string message)
        {
            if (!condition)
            {
                throw new Exception("Assertion failed: " + message);
            }
        }
    }
}
