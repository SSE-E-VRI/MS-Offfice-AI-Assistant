using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Strips the conversational wrapper an AI response carries around the content the user
    /// actually asked for -- the "Here's a polished draft for your request letter:" lead-in and
    /// the trailing "Notes for Customization" / "Let me know if..." block. Those belong in the
    /// chat pane, not in the document; users were finding them inserted into finished letters.
    ///
    /// Deliberately conservative: it removes a leading line only when it matches a known
    /// conversational opener, and removes a trailer only when the section is clearly marked
    /// (a thematic break followed by a note heading, or an unmistakable heading in the closing
    /// portion of the response). Anything it cannot classify is left untouched, and if the
    /// heuristics would empty the content the original text is returned unchanged.
    /// </summary>
    public static class ResponseContentCleaner
    {
        private static readonly Regex PreambleCue = new Regex(
            @"^(sure|certainly|of course|absolutely|got it|no problem|here's|here is|here are|here you go|below is|below are|the following|following is|as requested|as promised|this is|i've|i have|i'll|i will|happy to help|hope this helps|great question)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Used only after a thematic break, which is a strong "the document ended here" signal,
        // so a bare "Notes:" heading is safe to treat as a trailer in that position. Only the LAST
        // break is considered, so a letter's own "---" before an "Encl.:" block is never the cut
        // point. "Key features" lives here rather than in the strict list because a real spec or
        // product document can carry that heading in its body.
        private static readonly Regex TrailerCueLoose = new Regex(
            @"^(notes?|customization notes?|notes? for customization|next steps?|tips?|suggestions?|additional notes?|key changes|key features?|key points?|highlights?|why this works|how to use( this| it)?|usage notes?|format(ting)? notes?|template notes?|changes made|summary of changes|what i changed|explanation|instructions?|let me know|feel free|would you like|hope this helps|i hope this helps|happy to (help|assist)|if you'd like|if you want)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Used without a break, so "Note:" and "Instructions:" are excluded -- those appear
        // inside real letters and memos and must never be cut out of the body.
        private static readonly Regex TrailerCueStrict = new Regex(
            @"^(customization notes?|notes? for customization|next steps?|key changes|changes made|summary of changes|what i changed|let me know|feel free|would you like|hope this helps|i hope this helps|happy to (help|assist)|if you'd like|if you want)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex ThematicBreak = new Regex(
            @"^([-*_])[ \t]*(\1[ \t]*){2,}$", RegexOptions.CultureInvariant);

        private static readonly Regex OrderedListItem = new Regex(
            @"^\d+[.)]\s", RegexOptions.CultureInvariant);

        private static readonly Regex BoldHeadingLine = new Regex(
            @"^(\*\*|__).+(\*\*|__):?$", RegexOptions.CultureInvariant);

        // Signals that a response is a REVIEW of the text (a report about it) rather than a
        // single replacement FOR it. Seen in practice: asking for "a grammar check" on selected
        // text sometimes gets back a multi-section critique -- a revised draft, a table comparing
        // original vs. corrected wording, a "Key Improvements" rundown, and a second "recommended"
        // draft -- and Insert would apply the ENTIRE thing (table included) as the replacement.
        // Two independent tells, either one sufficient: a comparison-table header row, or two or
        // more analysis-report headings. One heading alone isn't enough -- a real letter can have
        // a body section titled e.g. "Key Details" without being a critique of itself.
        private static readonly Regex ComparisonTableKeyword = new Regex(
            @"^(original( text)?|issue|correction(s)?|corrected( text)?|reason|before|after|change(s)?)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex AnalysisHeading = new Regex(
            @"^(key improvements?|key changes|grammar (&|and) clarity corrections?|final recommendation|" +
            @"revised version|why this works|why\?|summary of changes|what( i|'s)? changed|corrections? made)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        /// <summary>
        /// Returns the response body with its conversational preamble and trailing notes removed.
        /// Never returns null or whitespace for non-empty input.
        /// </summary>
        public static string ExtractInsertableContent(string content)
        {
            if (string.IsNullOrEmpty(content)) return content;

            try
            {
                string normalized = content.Replace("\r\n", "\n").Replace("\r", "\n");
                var lines = new List<string>(normalized.Split('\n'));

                StripPreamble(lines);
                StripTrailer(lines);
                TrimBlankEdges(lines);

                string result = string.Join(Environment.NewLine, lines.ToArray());
                if (string.IsNullOrWhiteSpace(result)) return content;
                return result;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ResponseContentCleaner failed, inserting raw response: {0}", ex.Message));
                return content;
            }
        }

        private static void StripPreamble(List<string> lines)
        {
            // Loops because removing the lead-in sentence usually exposes the "---" rule the
            // model placed between its chatter and the actual draft.
            for (int guard = 0; guard < 8; guard++)
            {
                int i = FirstNonBlank(lines, 0);
                if (i < 0) return;

                string raw = lines[i].Trim();
                if (ThematicBreak.IsMatch(raw))
                {
                    RemoveThrough(lines, i);
                    continue;
                }

                if (IsStructuralLine(raw)) return;

                string bare = StripEmphasis(raw);
                if (!PreambleCue.IsMatch(bare)) return;

                // A cue word alone is not enough -- real content can open with "This is...".
                // Require the shape of a lead-in: a short line, or one ending in a colon.
                bool endsWithColon = bare.EndsWith(":");
                if (!endsWithColon && CountWords(bare) > 30) return;

                RemoveThrough(lines, i);
            }
        }

        private static void StripTrailer(List<string> lines)
        {
            // Strongest signal: a thematic break near the end whose following section is a
            // note / next-steps heading. Everything from the break down is chat, not document.
            int lastBreak = -1;
            for (int i = lines.Count - 1; i > 0; i--)
            {
                if (ThematicBreak.IsMatch(lines[i].Trim())) { lastBreak = i; break; }
            }

            if (lastBreak > 0)
            {
                int next = FirstNonBlank(lines, lastBreak + 1);
                if (next < 0)
                {
                    TruncateAt(lines, lastBreak);
                    return;
                }
                if (TrailerCueLoose.IsMatch(StripEmphasis(lines[next].Trim())))
                {
                    TruncateAt(lines, lastBreak);
                    return;
                }
            }

            // Weaker fallback: an unmistakable trailer heading in the closing half of the
            // response. Restricted to the second half so a body heading cannot be mistaken for one.
            int half = lines.Count / 2;
            for (int i = lines.Count - 1; i > half && i > 0; i--)
            {
                string t = lines[i].Trim();
                if (t.Length == 0) continue;
                if (!TrailerCueStrict.IsMatch(StripEmphasis(t))) continue;
                if (IsHeadingLike(t) || IsClosingRemark(t))
                {
                    TruncateAt(lines, i);
                    return;
                }
            }
        }

        private static bool IsStructuralLine(string line)
        {
            if (line.Length == 0) return false;
            if (line.StartsWith("#")) return true;
            if (line.StartsWith("|")) return true;
            if (line.StartsWith("```")) return true;
            if (line.StartsWith(">")) return true;
            if (line.StartsWith("- ") || line.StartsWith("* ") || line.StartsWith("+ ")) return true;
            if (OrderedListItem.IsMatch(line)) return true;
            return false;
        }

        private static bool IsHeadingLike(string line)
        {
            if (line.StartsWith("#")) return true;
            if (BoldHeadingLine.IsMatch(line)) return true;
            string bare = StripEmphasis(line);
            return bare.EndsWith(":") && CountWords(bare) <= 8;
        }

        private static bool IsClosingRemark(string line)
        {
            return CountWords(StripEmphasis(line)) <= 40;
        }

        /// <summary>
        /// True when a response reads as an analysis OF the text -- a comparison table and/or
        /// several "here's what I changed and why" headings -- rather than a single piece of
        /// replacement text. Insert applies a response verbatim in place of the selection, so this
        /// is the signal used to warn the user before a critique (table, rationale, alternate
        /// drafts included) lands in their document instead of just the corrected wording.
        /// Deliberately narrow: it flags, it never trims -- guessing which paragraph is "the real
        /// draft" out of a report like this is too easy to get wrong in a way that silently drops
        /// the user's actual content.
        /// </summary>
        public static bool LooksLikeEditAnalysisReport(string content)
        {
            if (string.IsNullOrWhiteSpace(content)) return false;

            try
            {
                string[] lines = content.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                if (HasComparisonTableHeader(lines)) return true;

                int analysisHeadings = 0;
                for (int i = 0; i < lines.Length; i++)
                {
                    string bare = StripEmphasis(lines[i]);
                    if (bare.Length == 0) continue;
                    if (AnalysisHeading.IsMatch(bare)) analysisHeadings++;
                }
                return analysisHeadings >= 2;
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ResponseContentCleaner.LooksLikeEditAnalysisReport failed: {0}", ex.Message));
                return false;
            }
        }

        private static bool HasComparisonTableHeader(string[] lines)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (!line.StartsWith("|") || !line.EndsWith("|")) continue;

                string[] cells = line.Substring(1, line.Length - 2).Split('|');
                int matches = 0;
                for (int c = 0; c < cells.Length; c++)
                {
                    string cell = StripEmphasis(cells[c]);
                    if (ComparisonTableKeyword.IsMatch(cell)) matches++;
                }
                if (matches >= 2) return true;
            }
            return false;
        }

        private static string StripEmphasis(string line)
        {
            string s = line.Trim().TrimStart('#', '*', '_', ' ', '\t').TrimEnd('*', '_', ' ', '\t');
            // Models emit curly apostrophes (U+2018/U+2019); normalise so the cue patterns
            // match either form. Written as escapes so the source stays pure ASCII and cannot be
            // mis-decoded by a compiler falling back to the system codepage.
            return s.Replace('\u2019', '\'').Replace('\u2018', '\'');
        }

        private static int CountWords(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;
            return text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length;
        }

        private static int FirstNonBlank(List<string> lines, int start)
        {
            for (int i = Math.Max(0, start); i < lines.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(lines[i])) return i;
            }
            return -1;
        }

        /// <summary>Removes everything from the start of the list through index i (inclusive).</summary>
        private static void RemoveThrough(List<string> lines, int i)
        {
            lines.RemoveRange(0, Math.Min(i + 1, lines.Count));
        }

        /// <summary>Removes index i and everything after it.</summary>
        private static void TruncateAt(List<string> lines, int i)
        {
            if (i < 0 || i >= lines.Count) return;
            lines.RemoveRange(i, lines.Count - i);
        }

        private static void TrimBlankEdges(List<string> lines)
        {
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[0])) lines.RemoveAt(0);
            while (lines.Count > 0 && string.IsNullOrWhiteSpace(lines[lines.Count - 1])) lines.RemoveAt(lines.Count - 1);
        }
    }
}
