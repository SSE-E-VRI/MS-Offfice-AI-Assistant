using System;
using System.Text;
using System.Text.RegularExpressions;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.Tests
{
    /// <summary>
    /// Covers the pure parts of the Copy path. The clipboard itself is never touched here --
    /// only the plain-text conversion and the CF_HTML envelope arithmetic.
    /// </summary>
    public static class MarkdownClipboardTests
    {
        public static void RunAll()
        {
            TestPlainTextStripsMarkers();
            TestPlainTextKeepsListStructure();
            TestCfHtmlOffsetsPointAtFragment();
            TestCfHtmlStaysAsciiWithNonAsciiInput();
            TestCfHtmlHandlesSurrogatePair();
            TestCfHtmlHandlesEmptyFragment();
        }

        private static void TestPlainTextStripsMarkers()
        {
            string md = "**Details of Required Pump:**\nI request **one (1) unit** of a *5 HP* pump with `spec` notes.";
            string plain = MarkdownClipboard.ToPlainText(md);

            Assert(!plain.Contains("**"), "Bold markers must be gone: " + plain);
            Assert(!plain.Contains("`"), "Code markers must be gone: " + plain);
            Assert(plain.Contains("Details of Required Pump:"), "Text must survive: " + plain);
            Assert(plain.Contains("one (1) unit"), "Bold text must survive: " + plain);
            Assert(plain.Contains("5 HP"), "Italic text must survive: " + plain);
        }

        private static void TestPlainTextKeepsListStructure()
        {
            string md = "## Pump\n- **Capacity:** 5 HP\n- **Quantity:** 1 No.\n\nSee [the manual](https://example.com/d) too.";
            string plain = MarkdownClipboard.ToPlainText(md).Replace("\r\n", "\n");

            Assert(plain.Contains("- Capacity: 5 HP"), "Bullets must be preserved: " + plain);
            Assert(plain.Contains("- Quantity: 1 No."), "Bullets must be preserved: " + plain);
            Assert(plain.StartsWith("Pump"), "Heading hashes must be stripped: " + plain);
            Assert(plain.Contains("the manual (https://example.com/d)"), "Links become text + url: " + plain);
            Assert(plain.Contains("\n\n"), "Blank-line structure must be preserved");
        }

        /// <summary>
        /// The four CF_HTML offsets are byte positions into the payload. If they are wrong Word
        /// pastes truncated or empty markup, so assert they really bracket the fragment.
        /// </summary>
        private static void TestCfHtmlOffsetsPointAtFragment()
        {
            string fragment = "<p><strong>Hello</strong> world</p>";
            string payload = MarkdownClipboard.BuildCfHtml(fragment);

            int startHtml, endHtml, startFragment, endFragment;
            ReadOffsets(payload, out startHtml, out endHtml, out startFragment, out endFragment);

            Assert(payload.Substring(startHtml, 6) == "<html>", "StartHTML must point at <html>");
            Assert(endHtml == payload.Length, string.Format("EndHTML {0} must equal payload length {1}", endHtml, payload.Length));
            Assert(payload.Substring(startFragment, endFragment - startFragment) == fragment,
                "Fragment offsets must bracket the fragment exactly");
            Assert(payload.Substring(startFragment - 20, 20) == "<!--StartFragment-->", "StartFragment sits after its marker");
            Assert(payload.Substring(endFragment, 18) == "<!--EndFragment-->", "EndFragment sits before its marker");
        }

        /// <summary>
        /// Offsets are counted in bytes but computed on chars, so the payload must be pure ASCII.
        /// A single curly quote used to be enough to shift every offset and break the paste.
        /// </summary>
        private static void TestCfHtmlStaysAsciiWithNonAsciiInput()
        {
            string payload = MarkdownClipboard.BuildCfHtml("<p>Curly \u2019quote\u2019 and \u20B9500</p>");

            Assert(Encoding.UTF8.GetByteCount(payload) == payload.Length,
                "Payload must be ASCII so byte offsets match character offsets");
            Assert(payload.Contains("&#8217;"), "Curly apostrophe must be escaped to an entity");
            Assert(payload.Contains("&#8377;"), "Rupee sign must be escaped to an entity");

            int startHtml, endHtml, startFragment, endFragment;
            ReadOffsets(payload, out startHtml, out endHtml, out startFragment, out endFragment);
            Assert(endHtml == payload.Length, "EndHTML must still equal payload length");
            Assert(payload.Substring(endFragment, 18) == "<!--EndFragment-->", "EndFragment must still be exact");
        }

        /// <summary>An emoji is one code point across two chars; escaping each half is invalid.</summary>
        private static void TestCfHtmlHandlesSurrogatePair()
        {
            string payload = MarkdownClipboard.BuildCfHtml("<p>ok \uD83D\uDE00</p>");

            Assert(payload.Contains("&#128512;"), "Surrogate pair must become one code point entity: " + payload);
            Assert(Encoding.UTF8.GetByteCount(payload) == payload.Length, "Payload must stay ASCII");
        }

        private static void TestCfHtmlHandlesEmptyFragment()
        {
            string payload = MarkdownClipboard.BuildCfHtml("");
            int startHtml, endHtml, startFragment, endFragment;
            ReadOffsets(payload, out startHtml, out endHtml, out startFragment, out endFragment);
            Assert(startFragment == endFragment, "Empty fragment means equal offsets");
            Assert(endHtml == payload.Length, "EndHTML must equal payload length");
        }

        private static void ReadOffsets(string payload, out int startHtml, out int endHtml,
            out int startFragment, out int endFragment)
        {
            var m = Regex.Match(payload,
                @"StartHTML:(\d+)\r\nEndHTML:(\d+)\r\nStartFragment:(\d+)\r\nEndFragment:(\d+)\r\n");
            if (!m.Success) throw new Exception("Assertion failed: CF_HTML header is malformed:\n" + payload);
            startHtml = int.Parse(m.Groups[1].Value);
            endHtml = int.Parse(m.Groups[2].Value);
            startFragment = int.Parse(m.Groups[3].Value);
            endFragment = int.Parse(m.Groups[4].Value);
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
