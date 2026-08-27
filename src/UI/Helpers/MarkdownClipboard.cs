using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using Markdig;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI.Helpers
{
    /// <summary>
    /// Puts an AI response on the clipboard the way a user expects to paste it.
    ///
    /// The Copy button used to call Clipboard.SetText on the raw response, so pasting into Word
    /// produced the model's chat wrapper plus literal Markdown markers ("**Details of Required
    /// Pump:**"). This writes two flavours instead: CF_HTML, which Word, Outlook and browsers
    /// paste as real bold / headings / lists / tables, and a marker-stripped plain-text fallback
    /// for Notepad and anything else that only takes text. Both are run through
    /// ResponseContentCleaner first, so Copy and Insert produce the same content.
    /// </summary>
    public static class MarkdownClipboard
    {
        private static readonly Regex HeadingPrefix = new Regex(@"^\s{0,3}#{1,6}\s+", RegexOptions.CultureInvariant);
        private static readonly Regex EmphasisPair = new Regex(@"(?<!\w)[*_](?=\S)(.+?)(?<=\S)[*_](?!\w)", RegexOptions.CultureInvariant);
        private static readonly Regex MarkdownLink = new Regex(@"\[([^\]]+)\]\(([^)\s]+)[^)]*\)", RegexOptions.CultureInvariant);

        /// <summary>
        /// Copies the response to the clipboard as formatted HTML plus clean plain text.
        /// Returns false only if the clipboard itself refused the data.
        /// </summary>
        public static bool SetResponse(string markdown)
        {
            if (string.IsNullOrEmpty(markdown)) return false;

            string content = markdown;
            try
            {
                bool strip = ConfigManager.Instance == null || ConfigManager.Instance.StripConversationalWrapper;
                if (strip) content = ResponseContentCleaner.ExtractInsertableContent(markdown);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("MarkdownClipboard: cleaner skipped ({0})", ex.Message));
                content = markdown;
            }

            string plain = ToPlainText(content);

            try
            {
                var pipeline = new MarkdownPipelineBuilder().UseAdvancedExtensions().Build();
                string html = Markdown.ToHtml(content, pipeline);

                var data = new DataObject();
                data.SetData(DataFormats.UnicodeText, plain);
                data.SetData(DataFormats.Text, plain);
                data.SetData(DataFormats.Html, BuildCfHtml(html));
                Clipboard.SetDataObject(data, true);
                return true;
            }
            catch (Exception ex)
            {
                // Rich copy is a convenience; never lose the user's copy over it.
                Logger.Warn(string.Format("MarkdownClipboard: HTML copy failed, falling back to text ({0})", ex.Message));
                try
                {
                    Clipboard.SetText(plain);
                    return true;
                }
                catch (Exception inner)
                {
                    Logger.Error("MarkdownClipboard: clipboard unavailable", inner);
                    return false;
                }
            }
        }

        /// <summary>
        /// Removes Markdown emphasis, heading and code markers while keeping the line structure
        /// (list bullets, numbering, blank lines) that makes a pasted letter still read as a letter.
        /// </summary>
        public static string ToPlainText(string markdown)
        {
            if (string.IsNullOrEmpty(markdown)) return markdown;

            string[] lines = markdown.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            var sb = new StringBuilder();
            for (int i = 0; i < lines.Length; i++)
            {
                string line = HeadingPrefix.Replace(lines[i], string.Empty);
                line = MarkdownLink.Replace(line, "$1 ($2)");
                line = line.Replace("**", string.Empty).Replace("__", string.Empty);
                line = EmphasisPair.Replace(line, "$1");
                line = line.Replace("`", string.Empty);
                if (i > 0) sb.Append(Environment.NewLine);
                sb.Append(line.TrimEnd());
            }
            return sb.ToString();
        }

        /// <summary>
        /// Wraps an HTML fragment in the CF_HTML envelope Windows requires: a header whose four
        /// offsets are counted in BYTES. Non-ASCII is escaped to numeric entities first so the
        /// fragment is pure ASCII and byte count equals character count -- otherwise the offsets
        /// drift the moment a response contains a curly quote, and Word pastes truncated markup.
        /// Public so the offset arithmetic can be unit-tested without touching the clipboard.
        /// </summary>
        public static string BuildCfHtml(string fragment)
        {
            const string headerFormat =
                "Version:0.9\r\n" +
                "StartHTML:{0:0000000000}\r\n" +
                "EndHTML:{1:0000000000}\r\n" +
                "StartFragment:{2:0000000000}\r\n" +
                "EndFragment:{3:0000000000}\r\n";
            const string preFragment = "<html><body>\r\n<!--StartFragment-->";
            const string postFragment = "<!--EndFragment-->\r\n</body></html>";

            string ascii = EscapeNonAscii(fragment);

            int headerLength = string.Format(CultureInfo.InvariantCulture, headerFormat, 0, 0, 0, 0).Length;
            int startFragment = headerLength + preFragment.Length;
            int endFragment = startFragment + ascii.Length;
            int endHtml = endFragment + postFragment.Length;

            return string.Format(CultureInfo.InvariantCulture, headerFormat, headerLength, endHtml, startFragment, endFragment)
                   + preFragment + ascii + postFragment;
        }

        private static string EscapeNonAscii(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;

            var sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c < 128)
                {
                    sb.Append(c);
                    continue;
                }

                int codePoint = c;
                // Surrogate pairs must be emitted as the single combined code point; escaping each
                // half separately produces two invalid entities.
                if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    codePoint = char.ConvertToUtf32(c, text[i + 1]);
                    i++;
                }
                sb.Append("&#").Append(codePoint.ToString(CultureInfo.InvariantCulture)).Append(';');
            }
            return sb.ToString();
        }
    }
}
