using System;
using System.Windows.Controls;
using System.Windows.Documents;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Tests citation-pattern detection and rendering in MarkdownHelper.
    /// Verifies that citation patterns ([¶N], ~Paragraph N, Sheet!Cell, Cell=Value, Slide N of M)
    /// are rendered as clickable Hyperlink elements with correct Tag values, while non-citation
    /// text remains as plain Run elements. Also verifies that evidence-level bracketed tags like
    /// [Calculated] are NOT rendered as citations.
    /// </summary>
    public static class MarkdownHelperCitationTests
    {
        public static void RunAll()
        {
            TestWordParagraphTagRendersAsHyperlink();
            TestWordExcerptLabelRendersAsHyperlink();
            TestExcelSheetQualifiedCellRendersAsHyperlink();
            TestExcelBareAddressEqualsValueRendersAsHyperlink();
            TestPowerPointSlideReferenceRendersAsHyperlink();
            TestPowerPointSlideTagBracketFormatRendersAsHyperlink();
            TestPowerPointSlideTagDashFormatRendersAsHyperlink();
            TestPlainTextWithoutCitationsHasNoHyperlinks();
            TestEvidenceLevelBracketedTagDoesNotBecomeHyperlink();
            TestCitationInSentenceRendersCorrectly();
        }

        private static void TestWordParagraphTagRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "See [¶12] for details.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "[¶12]");
            Assert(link != null, "Word paragraph tag [¶12] should render as a Hyperlink");
            Assert(link.Tag.Equals("[¶12]"), "Hyperlink Tag should be exactly '[¶12]'");
        }

        private static void TestWordExcerptLabelRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "Based on ~Paragraph 5 from the document.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "~Paragraph 5");
            Assert(link != null, "Word excerpt label ~Paragraph 5 should render as a Hyperlink");
            Assert(link.Tag.Equals("~Paragraph 5"), "Hyperlink Tag should be exactly '~Paragraph 5'");
        }

        private static void TestExcelSheetQualifiedCellRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "The value in Sheet1!B7 is shown.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "Sheet1!B7");
            Assert(link != null, "Excel sheet-qualified cell Sheet1!B7 should render as a Hyperlink");
            Assert(link.Tag.Equals("Sheet1!B7"), "Hyperlink Tag should be exactly 'Sheet1!B7'");
        }

        private static void TestExcelBareAddressEqualsValueRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "From cell B7=1234 we see the result.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "B7=");
            Assert(link != null, "Excel bare cell address B7= should render as a Hyperlink");
            Assert(link.Tag.Equals("B7="), "Hyperlink Tag should be 'B7='");
        }

        private static void TestPowerPointSlideReferenceRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "Look at Slide 3 of 12 in the presentation.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "Slide 3 of 12");
            Assert(link != null, "PowerPoint slide reference Slide 3 of 12 should render as a Hyperlink");
            Assert(link.Tag.Equals("Slide 3 of 12"), "Hyperlink Tag should be exactly 'Slide 3 of 12'");
        }

        /// <summary>
        /// PowerPointController.GetSlideTextInternal's real emitted format — this is the format
        /// that actually reaches the model as context for a full-deck fetch or review context,
        /// unlike "Slide N of M" which is only ever used for the A2 context-readout UI label and
        /// never sent as content. Found missing during an adversarial review pass; fixed alongside
        /// this test.
        /// </summary>
        private static void TestPowerPointSlideTagBracketFormatRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "See [Slide #3: Overview] for details.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "[Slide #3: Overview]");
            Assert(link != null, "PowerPoint slide tag [Slide #3: Overview] should render as a Hyperlink");
            Assert(link.Tag.Equals("[Slide #3: Overview]"), "Hyperlink Tag should be exactly '[Slide #3: Overview]'");
        }

        /// <summary>
        /// AttachmentExtractor's real emitted format for .pptx slide sections — the second of the
        /// two real, actually-used-as-content PowerPoint formats that "Slide N of M" alone missed.
        /// </summary>
        private static void TestPowerPointSlideTagDashFormatRendersAsHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "From --- Slide 5 --- we can see the trend.");

            Hyperlink link = FindHyperlinkByTag(textBlock, "--- Slide 5 ---");
            Assert(link != null, "PowerPoint slide tag --- Slide 5 --- should render as a Hyperlink");
            Assert(link.Tag.Equals("--- Slide 5 ---"), "Hyperlink Tag should be exactly '--- Slide 5 ---'");
        }

        private static void TestPlainTextWithoutCitationsHasNoHyperlinks()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "This is plain text with no citations whatsoever.");

            Hyperlink link = FindFirstHyperlink(textBlock);
            Assert(link == null, "Plain text without any citation patterns should not contain any Hyperlinks");
        }

        private static void TestEvidenceLevelBracketedTagDoesNotBecomeHyperlink()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "[Calculated] This result was derived from the data.");

            // [Calculated] is NOT a citation pattern; it's an evidence-level tag that should NOT become a hyperlink
            Hyperlink link = FindHyperlinkByTag(textBlock, "[Calculated]");
            Assert(link == null, "[Calculated] bracketed evidence tag should NOT render as a Hyperlink");

            // Verify text is still present (not lost during parsing)
            string plainText = GetPlainText(textBlock);
            Assert(plainText.Contains("Calculated"), "[Calculated] text should still appear in output");
            Assert(plainText.Contains("derived from the data"), "Citation sentence text should still appear in output");
        }

        private static void TestCitationInSentenceRendersCorrectly()
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, "See [¶5] for context on Sheet1!A1 and Slide 2 of 10.");

            Hyperlink link1 = FindHyperlinkByTag(textBlock, "[¶5]");
            Hyperlink link2 = FindHyperlinkByTag(textBlock, "Sheet1!A1");
            Hyperlink link3 = FindHyperlinkByTag(textBlock, "Slide 2 of 10");

            Assert(link1 != null, "[¶5] should be a Hyperlink");
            Assert(link2 != null, "Sheet1!A1 should be a Hyperlink");
            Assert(link3 != null, "Slide 2 of 10 should be a Hyperlink");

            // Verify sentence structure is preserved
            string plainText = GetPlainText(textBlock);
            Assert(plainText.Contains("See "), "Sentence intro should be present");
            Assert(plainText.Contains("for context on "), "Sentence connector text should be present");
            Assert(plainText.Contains("and "), "Sentence conjunction should be present");
        }

        // --- helpers -----------------------------------------------------------------------

        private static Hyperlink FindHyperlinkByTag(TextBlock textBlock, string tagValue)
        {
            foreach (Inline inline in textBlock.Inlines)
            {
                Hyperlink link = FindHyperlinkByTagInline(inline, tagValue);
                if (link != null) return link;
            }
            return null;
        }

        private static Hyperlink FindHyperlinkByTagInline(Inline inline, string tagValue)
        {
            var link = inline as Hyperlink;
            if (link != null && link.Tag != null && link.Tag.Equals(tagValue))
                return link;

            var span = inline as Span;
            if (span != null)
            {
                foreach (Inline child in span.Inlines)
                {
                    Hyperlink childLink = FindHyperlinkByTagInline(child, tagValue);
                    if (childLink != null) return childLink;
                }
            }

            return null;
        }

        private static Hyperlink FindFirstHyperlink(TextBlock textBlock)
        {
            foreach (Inline inline in textBlock.Inlines)
            {
                Hyperlink link = FindFirstHyperlinkInline(inline);
                if (link != null) return link;
            }
            return null;
        }

        private static Hyperlink FindFirstHyperlinkInline(Inline inline)
        {
            var link = inline as Hyperlink;
            if (link != null) return link;

            var span = inline as Span;
            if (span != null)
            {
                foreach (Inline child in span.Inlines)
                {
                    Hyperlink childLink = FindFirstHyperlinkInline(child);
                    if (childLink != null) return childLink;
                }
            }

            return null;
        }

        private static string GetPlainText(TextBlock textBlock)
        {
            var range = new TextRange(textBlock.ContentStart, textBlock.ContentEnd);
            return range.Text;
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
