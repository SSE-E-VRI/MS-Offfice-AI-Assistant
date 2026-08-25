using System;
using System.Windows.Controls;
using System.Windows.Documents;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Covers the D-7 incremental streaming markdown renderer: the guarantee is that whatever
    /// happens mid-stream, the state after streaming completes must be identical to a single,
    /// from-scratch full parse of the complete final string.
    /// </summary>
    public static class MarkdownHelperTests
    {
        public static void RunAll()
        {
            TestStreamedFencedCodeAndNumberedListMatchesFullParse();
            TestStreamedTextSplitMidFenceMarkerMatchesFullParse();
            TestNonStreamingSinglePassStillRendersBasics();
        }

        // A numbered list and a fenced code block, split into small chunks that land mid-line,
        // mid-marker, and mid-word — i.e. not aligned to any markdown-syntactic boundary.
        private static void TestStreamedFencedCodeAndNumberedListMatchesFullParse()
        {
            string full =
                "Steps:\n" +
                "1. First step goes here\n" +
                "2. Second step continues\n" +
                "\n" +
                "```csharp\n" +
                "var x = 1;\n" +
                "Console.WriteLine(x);\n" +
                "```\n" +
                "\n" +
                "Done.";

            string streamedPlainText = RenderViaStreamingChunks(full, 7);
            string fullParsePlainText = RenderViaSingleFullParse(full);

            Assert(streamedPlainText == fullParsePlainText,
                "Streamed (chunked) render must equal a single full parse after streaming completes.\n" +
                "Streamed: " + Escape(streamedPlainText) + "\nFull:     " + Escape(fullParsePlainText));

            // Sanity: the content actually got through structurally, not just as an empty/fallback string.
            Assert(streamedPlainText.Contains("First step goes here"), "Numbered list item 1 text missing");
            Assert(streamedPlainText.Contains("Second step continues"), "Numbered list item 2 text missing");
            Assert(streamedPlainText.Contains("var x = 1;"), "Fenced code content missing");
            Assert(streamedPlainText.Contains("Console.WriteLine(x);"), "Fenced code content missing (line 2)");
            Assert(streamedPlainText.Contains("Done."), "Trailing paragraph after code fence missing");
        }

        // Deliberately chunk so the fence delimiter "```csharp" itself is split across two deltas
        // (a chunk boundary landing inside the ``` marker), which is the exact scenario D-7 calls out.
        private static void TestStreamedTextSplitMidFenceMarkerMatchesFullParse()
        {
            string full = "Before\n```py\nprint('hi')\n```\nAfter";

            // Split so the fence marker "```py" straddles two chunks: "Before\n``" | "`py\nprint..." etc.
            string[] chunks = { "Before\n``", "`py\nprint(", "'hi')\n```", "\nAfter" };

            string streamedPlainText = RenderViaStreamingChunks(full, chunks);
            string fullParsePlainText = RenderViaSingleFullParse(full);

            Assert(streamedPlainText == fullParsePlainText,
                "Fence marker split across a chunk boundary must still match a full parse once streaming ends.\n" +
                "Streamed: " + Escape(streamedPlainText) + "\nFull:     " + Escape(fullParsePlainText));
            Assert(streamedPlainText.Contains("print('hi')"), "Code fence content missing");
            Assert(streamedPlainText.Contains("Before"), "Text before fence missing");
            Assert(streamedPlainText.Contains("After"), "Text after fence missing");
        }

        private static void TestNonStreamingSinglePassStillRendersBasics()
        {
            string markdown = "# Title\n\nSome **bold** and *italic* and `code`.\n\n- one\n- two\n\n> a quote";

            var textBlock = new TextBlock();
            MarkdownHelper.SetIsStreaming(textBlock, false);
            MarkdownHelper.SetMarkdownText(textBlock, markdown);

            string plain = GetPlainText(textBlock);
            Assert(plain.Contains("Title"), "Heading text missing");
            Assert(plain.Contains("bold"), "Bold text missing");
            Assert(plain.Contains("italic"), "Italic text missing");
            Assert(plain.Contains("code"), "Inline code text missing");
            Assert(plain.Contains("one") && plain.Contains("two"), "Bullet list items missing");
            Assert(plain.Contains("a quote"), "Blockquote text missing");
        }

        // --- helpers -----------------------------------------------------------------------

        private static string RenderViaStreamingChunks(string full, int chunkSize)
        {
            var chunks = new System.Collections.Generic.List<string>();
            for (int i = 0; i < full.Length; i += chunkSize)
            {
                int len = Math.Min(chunkSize, full.Length - i);
                chunks.Add(full.Substring(i, len));
            }
            return RenderViaStreamingChunks(full, chunks.ToArray());
        }

        private static string RenderViaStreamingChunks(string full, string[] chunks)
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetIsStreaming(textBlock, true);

            string accumulated = string.Empty;
            foreach (string chunk in chunks)
            {
                accumulated += chunk;
                MarkdownHelper.SetMarkdownText(textBlock, accumulated);
            }

            Assert(accumulated == full, "Test bug: chunks did not reassemble to the full string");

            // Mirrors AssistantSession.ProcessAssistantResponse: IsStreaming flips false, then the
            // final Content is set — triggering MarkdownHelper's mandatory full re-parse.
            MarkdownHelper.SetIsStreaming(textBlock, false);
            MarkdownHelper.SetMarkdownText(textBlock, full);

            return GetPlainText(textBlock);
        }

        private static string RenderViaSingleFullParse(string full)
        {
            var textBlock = new TextBlock();
            MarkdownHelper.SetMarkdownText(textBlock, full);
            return GetPlainText(textBlock);
        }

        private static string GetPlainText(TextBlock textBlock)
        {
            var range = new TextRange(textBlock.ContentStart, textBlock.ContentEnd);
            return range.Text;
        }

        private static string Escape(string s)
        {
            return s.Replace("\r", "\\r").Replace("\n", "\\n");
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
