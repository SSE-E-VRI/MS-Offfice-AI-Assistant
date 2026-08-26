using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace MSOfficeAIAssistant.UI.Helpers
{
    /// <summary>
    /// Attached property that renders Markdown text directly into a WPF TextBlock as formatted Inlines.
    /// Handles headers (###), bold (**), italic (*), lists (-, 1.), inline code (`), and code blocks (```).
    ///
    /// Colors mirror src/UI/Theme/Tokens.xaml (InkStrongBrush, InkBrush, PrimaryBrush, InkMutedBrush,
    /// CodeInlineTextBrush, SurfaceAltBrush). Kept as plain static brushes here rather than loaded from
    /// the XAML dictionary because this runs on the streaming hot path (D-7) — see AppendMarkdown below.
    ///
    /// Streaming (D-7 part 2): when <see cref="IsStreamingProperty"/> is true and the new MarkdownText
    /// value is a pure append onto the text previously rendered for that TextBlock, only the newly
    /// completed lines are parsed and appended to the existing Inlines — the whole accumulated string is
    /// no longer re-parsed on every delta. A markdown marker (heading, bullet, fence, ...) can only be
    /// decided once its whole line has arrived — deciding from a bare prefix is exactly what breaks when
    /// e.g. a fenced ``` delimiter itself lands split across two deltas — so the still-open trailing line
    /// of each snapshot is deliberately left unrendered until a newline completes it (or the stream ends).
    /// Complete, already-committed lines are never revisited. Once streaming ends, IsStreaming flips to
    /// false and MarkdownHelper always does one full, from-scratch re-parse of the final text — so the
    /// end state is guaranteed identical to what a single full parse of the complete string produces,
    /// regardless of how the text happened to be chunked while it was arriving.
    /// </summary>
    public static class MarkdownHelper
    {
        public static readonly DependencyProperty MarkdownTextProperty =
            DependencyProperty.RegisterAttached(
                "MarkdownText",
                typeof(string),
                typeof(MarkdownHelper),
                new PropertyMetadata(string.Empty, OnMarkdownTextChanged));

        public static string GetMarkdownText(DependencyObject obj)
        {
            return (string)obj.GetValue(MarkdownTextProperty);
        }

        public static void SetMarkdownText(DependencyObject obj, string value)
        {
            obj.SetValue(MarkdownTextProperty, value);
        }

        /// <summary>
        /// Bind to a ChatMessage's IsStreaming flag. While true, MarkdownText updates are treated as
        /// incremental appends where possible. When it is false (including the final "stream finished"
        /// update), MarkdownHelper always does a full re-parse and discards any incremental state.
        /// </summary>
        public static readonly DependencyProperty IsStreamingProperty =
            DependencyProperty.RegisterAttached(
                "IsStreaming",
                typeof(bool),
                typeof(MarkdownHelper),
                new PropertyMetadata(false, OnIsStreamingChanged));

        public static bool GetIsStreaming(DependencyObject obj)
        {
            return (bool)obj.GetValue(IsStreamingProperty);
        }

        public static void SetIsStreaming(DependencyObject obj, bool value)
        {
            obj.SetValue(IsStreamingProperty, value);
        }

        // Per-TextBlock incremental parse state. Weak-keyed so it never outlives the TextBlock
        // (message bubbles are recycled/discarded routinely, especially now that the message list
        // virtualizes — see D-7 part 1).
        private static readonly ConditionalWeakTable<TextBlock, RenderState> _states =
            new ConditionalWeakTable<TextBlock, RenderState>();

        private sealed class RenderState
        {
            // The full markdown text last seen for this TextBlock (used to validate that the next
            // update is a pure append onto it).
            public string LastText = string.Empty;
            // How many leading characters of LastText are already reflected in textBlock.Inlines.
            // Always sits at a line boundary (right after a newline) except at the very end of a full
            // parse, where it equals LastText.Length. Anything past this point is an as-yet-incomplete
            // trailing line that incremental appends deliberately hold back — see class remarks.
            public int CommittedLength;
            public bool InCodeBlock;
            public readonly List<string> CodeBlockLines = new List<string>();
        }

        private static readonly SolidColorBrush InkStrongBrush = Freeze(Color.FromRgb(0x0F, 0x17, 0x2A));
        private static readonly SolidColorBrush InkBrush = Freeze(Color.FromRgb(0x1E, 0x29, 0x3B));
        private static readonly SolidColorBrush PrimaryBrush = Freeze(Color.FromRgb(0x25, 0x63, 0xEB));
        private static readonly SolidColorBrush InkMutedBrush = Freeze(Color.FromRgb(0x47, 0x55, 0x69));
        private static readonly SolidColorBrush CodeInlineTextBrush = Freeze(Color.FromRgb(0xB4, 0x53, 0x09));
        private static readonly SolidColorBrush SurfaceAltBrush = Freeze(Color.FromRgb(0xF1, 0xF5, 0xF9));

        private static SolidColorBrush Freeze(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }

        private static void OnMarkdownTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var textBlock = d as TextBlock;
            if (textBlock == null) return;

            string markdown = e.NewValue as string;
            if (markdown == null) markdown = string.Empty;

            bool streaming = GetIsStreaming(textBlock);

            RenderState state;
            bool hasState = _states.TryGetValue(textBlock, out state);

            if (streaming && hasState && IsPureAppend(markdown, state.LastText))
            {
                try
                {
                    AppendIncremental(textBlock, state, markdown);
                    state.LastText = markdown;
                }
                catch
                {
                    // Fallback to plain text in case of any parsing exception.
                    textBlock.Inlines.Clear();
                    textBlock.Inlines.Add(new Run(markdown));
                    _states.Remove(textBlock);
                }
            }
            else
            {
                // Not streaming (this is the authoritative final render), no prior state for this
                // TextBlock, or the text wasn't a pure append (e.g. a fresh message reusing a
                // recycled container) — always a full, from-scratch parse.
                RenderFullSafe(textBlock, markdown);
            }
        }

        /// <summary>
        /// Streaming has just ended (or this confirms the message was never streaming). Forces a full
        /// re-parse of whatever MarkdownText currently holds right now — this must NOT wait for a
        /// subsequent MarkdownText change, because WPF suppresses the MarkdownText change callback
        /// entirely when the final Content value coincidentally equals the last streamed snapshot,
        /// which would otherwise leave a held-back trailing line (see AppendIncremental) permanently
        /// unrendered.
        /// </summary>
        private static void OnIsStreamingChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            bool streaming = (bool)e.NewValue;
            if (streaming) return;

            var textBlock = d as TextBlock;
            if (textBlock == null) return;

            RenderFullSafe(textBlock, GetMarkdownText(textBlock));
        }

        /// <summary>
        /// Clears and fully re-renders textBlock from markdown, resetting all incremental state.
        /// The one authoritative rendering path — guarantees streamed output converges to exactly
        /// what a single full parse of the complete string produces.
        /// </summary>
        private static void RenderFullSafe(TextBlock textBlock, string markdown)
        {
            if (markdown == null) markdown = string.Empty;

            try
            {
                textBlock.Inlines.Clear();
                var state = new RenderState();
                _states.Remove(textBlock);
                _states.Add(textBlock, state);

                if (markdown.Length > 0)
                {
                    RenderFull(textBlock, state, markdown);
                }
                state.LastText = markdown;
                state.CommittedLength = markdown.Length;
            }
            catch
            {
                // Fallback to plain text in case of any parsing exception.
                textBlock.Inlines.Clear();
                textBlock.Inlines.Add(new Run(markdown));
                _states.Remove(textBlock);
            }
        }

        private static bool IsPureAppend(string newText, string previous)
        {
            if (previous.Length == 0) return true;
            if (newText.Length < previous.Length) return false;
            return string.CompareOrdinal(newText, 0, previous, 0, previous.Length) == 0;
        }

        private static string[] SplitLines(string text)
        {
            return text.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        }

        /// <summary>
        /// Full, from-scratch parse of the complete text (used for non-streaming messages and for the
        /// mandatory final re-parse once streaming ends). Behaviorally identical to the original
        /// single-pass renderer this replaces.
        /// </summary>
        private static void RenderFull(TextBlock textBlock, RenderState state, string markdown)
        {
            string[] lines = SplitLines(markdown);
            for (int i = 0; i < lines.Length; i++)
            {
                bool isLastLine = (i == lines.Length - 1);
                ProcessLine(textBlock, state, lines[i], isLastLine);
            }

            if (state.InCodeBlock && state.CodeBlockLines.Count > 0)
            {
                AppendCodeSpan(textBlock, state.CodeBlockLines, /*trailingLineBreak:*/ false);
                state.CodeBlockLines.Clear();
            }
        }

        /// <summary>
        /// Renders only the lines that have newly become complete (terminated by an actual newline)
        /// since <see cref="RenderState.CommittedLength"/>. The trailing line — which may still be
        /// mid-word, or mid-marker like an unfinished ``` fence — is deliberately left unrendered;
        /// it will be picked up by a later call once it completes, or by the final full re-parse.
        /// </summary>
        private static void AppendIncremental(TextBlock textBlock, RenderState state, string markdown)
        {
            int scanPos = state.CommittedLength;
            while (true)
            {
                int nextStart;
                int lineEnd = FindLineEnd(markdown, scanPos, out nextStart);
                if (lineEnd < 0)
                {
                    // No newline yet in the remainder — hold it back uncommitted.
                    break;
                }

                string line = markdown.Substring(scanPos, lineEnd - scanPos);
                ProcessLine(textBlock, state, line, /*isLastLine:*/ false);
                state.CommittedLength = nextStart;
                scanPos = nextStart;
            }
        }

        /// <summary>
        /// Finds the end of the line starting at <paramref name="start"/> (index of the first line-ending
        /// character, treating \r\n as a single terminator), returning -1 if <paramref name="text"/> has
        /// no line terminator at or after <paramref name="start"/>. <paramref name="nextStart"/> is the
        /// index immediately following the terminator (undefined when -1 is returned).
        /// </summary>
        private static int FindLineEnd(string text, int start, out int nextStart)
        {
            for (int i = start; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\n')
                {
                    nextStart = i + 1;
                    return i;
                }
                if (c == '\r')
                {
                    nextStart = (i + 1 < text.Length && text[i + 1] == '\n') ? i + 2 : i + 1;
                    return i;
                }
            }
            nextStart = -1;
            return -1;
        }

        private static void ProcessLine(TextBlock textBlock, RenderState state, string line, bool isLastLine)
        {
            // 1. Code Block Toggle (```)
            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal))
            {
                if (state.InCodeBlock)
                {
                    AppendCodeSpan(textBlock, state.CodeBlockLines, /*trailingLineBreak:*/ true);
                    state.CodeBlockLines.Clear();
                    state.InCodeBlock = false;
                }
                else
                {
                    state.InCodeBlock = true;
                    state.CodeBlockLines.Clear();
                }
                return;
            }

            if (state.InCodeBlock)
            {
                state.CodeBlockLines.Add(line);
                return;
            }

            // 2. Empty Line -> Paragraph spacing
            if (string.IsNullOrWhiteSpace(line))
            {
                textBlock.Inlines.Add(new LineBreak());
                return;
            }

            string trimmed = line.TrimStart();

            // 3. Headings (# H1, ## H2, ### H3, #### H4)
            if (trimmed.StartsWith("#### ", StringComparison.Ordinal))
            {
                AppendHeading(textBlock, trimmed.Substring(5), 13.0, InkBrush);
                return;
            }
            if (trimmed.StartsWith("### ", StringComparison.Ordinal))
            {
                AppendHeading(textBlock, trimmed.Substring(4), 13.5, InkStrongBrush);
                return;
            }
            if (trimmed.StartsWith("## ", StringComparison.Ordinal))
            {
                AppendHeading(textBlock, trimmed.Substring(3), 14.5, InkStrongBrush);
                return;
            }
            if (trimmed.StartsWith("# ", StringComparison.Ordinal))
            {
                AppendHeading(textBlock, trimmed.Substring(2), 16.0, InkStrongBrush);
                return;
            }

            // 4. Bullet Lists (- item, * item, + item)
            if (trimmed.StartsWith("- ") || trimmed.StartsWith("* ") || trimmed.StartsWith("+ "))
            {
                var bulletRun = new Run(" •  ") { FontWeight = FontWeights.Bold, Foreground = PrimaryBrush };
                textBlock.Inlines.Add(bulletRun);
                ParseInlineFormatting(textBlock, trimmed.Substring(2));
                textBlock.Inlines.Add(new LineBreak());
                return;
            }

            // 5. Numbered Lists (1. item, 2. item, etc.)
            Match numMatch = Regex.Match(trimmed, @"^(\d+)\.\s+(.*)$");
            if (numMatch.Success)
            {
                string num = numMatch.Groups[1].Value;
                string rest = numMatch.Groups[2].Value;
                var numRun = new Run(string.Format(" {0}. ", num)) { FontWeight = FontWeights.Bold, Foreground = PrimaryBrush };
                textBlock.Inlines.Add(numRun);
                ParseInlineFormatting(textBlock, rest);
                textBlock.Inlines.Add(new LineBreak());
                return;
            }

            // 6. Blockquote (> quote)
            if (trimmed.StartsWith("> "))
            {
                var quoteSpan = new Span { FontStyle = FontStyles.Italic, Foreground = InkMutedBrush };
                quoteSpan.Inlines.Add(new Run("▎ "));
                ParseInlineFormatting(quoteSpan, trimmed.Substring(2));
                textBlock.Inlines.Add(quoteSpan);
                textBlock.Inlines.Add(new LineBreak());
                return;
            }

            // 7. Regular paragraph text with inline formatting
            ParseInlineFormatting(textBlock, line);

            // Add line break if not the last line of this batch (full text on a full parse; the
            // newly-arrived suffix on an incremental append).
            if (!isLastLine)
            {
                textBlock.Inlines.Add(new LineBreak());
            }
        }

        private static void AppendHeading(TextBlock textBlock, string content, double fontSize, SolidColorBrush foreground)
        {
            var headerSpan = new Span { FontWeight = FontWeights.Bold, FontSize = fontSize, Foreground = foreground };
            ParseInlineFormatting(headerSpan, content);
            textBlock.Inlines.Add(headerSpan);
            textBlock.Inlines.Add(new LineBreak());
        }

        private static void AppendCodeSpan(TextBlock textBlock, List<string> codeBlockLines, bool trailingLineBreak)
        {
            var codeSpan = new Span
            {
                FontFamily = new FontFamily("Consolas, Courier New, monospace"),
                FontSize = 11.5,
                Foreground = InkStrongBrush,
                Background = SurfaceAltBrush
            };
            string codeContent = string.Join("\n", codeBlockLines.ToArray());
            codeSpan.Inlines.Add(new Run(codeContent));
            textBlock.Inlines.Add(codeSpan);
            if (trailingLineBreak)
            {
                textBlock.Inlines.Add(new LineBreak());
            }
        }

        private static void ParseInlineFormatting(Span parentSpan, string text)
        {
            ParseInlineCore(parentSpan.Inlines, text);
        }

        private static void ParseInlineFormatting(TextBlock parentBlock, string text)
        {
            ParseInlineCore(parentBlock.Inlines, text);
        }

        private static void ParseInlineCore(InlineCollection targetInlines, string text)
        {
            if (string.IsNullOrEmpty(text)) return;

            // Regex tokenizes:
            // 1. `inline code`
            // 2. ***bold italic***
            // 3. **bold** or __bold__
            // 4. *italic* or _italic_
            string pattern = @"(`[^`]+`|\*\*\*[^*]+\*\*\*|\*\*[^*]+\*\*|__[^_]+__|\*[^*]+\*|_[^_]+_)";
            string[] parts = Regex.Split(text, pattern);

            foreach (string part in parts)
            {
                if (string.IsNullOrEmpty(part)) continue;

                // 1. Inline Code `code`
                if (part.StartsWith("`") && part.EndsWith("`") && part.Length >= 2)
                {
                    string code = part.Substring(1, part.Length - 2);
                    var codeRun = new Run(code)
                    {
                        FontFamily = new FontFamily("Consolas, Courier New, monospace"),
                        FontSize = 12.0,
                        Background = SurfaceAltBrush,
                        Foreground = CodeInlineTextBrush
                    };
                    targetInlines.Add(codeRun);
                }
                // 2. Bold + Italic ***text***
                else if (part.StartsWith("***") && part.EndsWith("***") && part.Length >= 6)
                {
                    string inner = part.Substring(3, part.Length - 6);
                    var bi = new Bold(new Italic(new Run(inner)));
                    targetInlines.Add(bi);
                }
                // 3. Bold **text** or __text__
                else if ((part.StartsWith("**") && part.EndsWith("**") && part.Length >= 4) ||
                         (part.StartsWith("__") && part.EndsWith("__") && part.Length >= 4))
                {
                    string inner = part.Substring(2, part.Length - 4);
                    var bold = new Bold(new Run(inner));
                    targetInlines.Add(bold);
                }
                // 4. Italic *text* or _text_
                else if ((part.StartsWith("*") && part.EndsWith("*") && part.Length >= 2) ||
                         (part.StartsWith("_") && part.EndsWith("_") && part.Length >= 2))
                {
                    string inner = part.Substring(1, part.Length - 2);
                    var italic = new Italic(new Run(inner));
                    targetInlines.Add(italic);
                }
                // 5. Plain Text (may contain citation patterns that become hyperlinks)
                else
                {
                    AddPlainTextWithCitations(targetInlines, part);
                }
            }
        }

        /// <summary>
        /// Tokenizes plain text for citation patterns and renders them as Hyperlink elements,
        /// while keeping plain text as inert Run elements.
        /// Five citation patterns are recognized:
        /// 1. [¶N] — Word paragraph tag (e.g. [¶12])
        /// 2. ~Paragraph N — Word excerpt label (e.g. ~Paragraph 5)
        /// 3. SheetName!Address — Excel sheet-qualified cell (e.g. Sheet1!B7)
        /// 4. Address=Value — Excel bare cell tag (e.g. B7=1234; only Address is clickable)
        /// 5. Slide N of M — PowerPoint slide reference (e.g. Slide 3 of 12)
        /// </summary>
        private static void AddPlainTextWithCitations(InlineCollection targetInlines, string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            // Combined regex that captures all 5 citation patterns
            // Pattern 1: [¶N] - Word paragraph tag
            // Pattern 2: ~Paragraph N - Word excerpt label
            // Pattern 3: Sheet!Address - Excel sheet-qualified cell
            // Pattern 4: Address=Value - Excel bare cell tag
            // Pattern 5: Slide N of M - PowerPoint slide reference
            string pattern =
                @"(\[¶\d+\])" +                                    // Group 1: [¶N]
                @"|" + @"(~Paragraph\s+\d+)" +                     // Group 2: ~Paragraph N
                @"|" + @"([A-Za-z0-9_]+!\$?[A-Z]+\$?\d+)" +        // Group 3: Sheet!Address
                @"|" + @"([A-Z]{1,3}\d{1,7}=)" +                   // Group 4: Address= (full match with =)
                @"|" + @"(Slide\s+\d+\s+of\s+\d+)";                // Group 5: Slide N of M

            string[] parts = Regex.Split(text, pattern);

            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i];
                if (string.IsNullOrEmpty(part))
                    continue;

                // Determine if this part is a citation by checking if it matches any pattern
                bool isCitation = false;

                // Check each pattern
                if (Regex.IsMatch(part, @"^\[¶\d+\]$"))
                {
                    isCitation = true;
                }
                else if (Regex.IsMatch(part, @"^~Paragraph\s+\d+$"))
                {
                    isCitation = true;
                }
                else if (Regex.IsMatch(part, @"^[A-Za-z0-9_]+!\$?[A-Z]+\$?\d+$"))
                {
                    isCitation = true;
                }
                else if (Regex.IsMatch(part, @"^[A-Z]{1,3}\d{1,7}=$"))
                {
                    isCitation = true;
                }
                else if (Regex.IsMatch(part, @"^Slide\s+\d+\s+of\s+\d+$"))
                {
                    isCitation = true;
                }

                if (isCitation)
                {
                    // Create a clickable hyperlink for this citation
                    var link = new Hyperlink(new Run(part))
                    {
                        Tag = part,
                        Foreground = PrimaryBrush
                    };
                    link.TextDecorations = null;  // Remove default underline for subtle inline citation
                    targetInlines.Add(link);
                }
                else
                {
                    // Plain text, not a citation
                    targetInlines.Add(new Run(part));
                }
            }
        }
    }
}
