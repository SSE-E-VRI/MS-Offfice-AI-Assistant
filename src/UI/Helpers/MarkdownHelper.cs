using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;

namespace MistralOfficeAddin.UI.Helpers
{
    /// <summary>
    /// Attached property that renders Markdown text directly into a WPF TextBlock as formatted Inlines.
    /// Handles headers (###), bold (**), italic (*), lists (-, 1.), inline code (`), and code blocks (```).
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

        private static void OnMarkdownTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            var textBlock = d as TextBlock;
            if (textBlock == null) return;

            string markdown = e.NewValue as string;
            textBlock.Inlines.Clear();

            if (string.IsNullOrEmpty(markdown))
                return;

            try
            {
                RenderMarkdown(textBlock, markdown);
            }
            catch
            {
                // Fallback to plain text in case of any parsing exception
                textBlock.Inlines.Clear();
                textBlock.Inlines.Add(new Run(markdown));
            }
        }

        private static void RenderMarkdown(TextBlock textBlock, string markdown)
        {
            string[] lines = markdown.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            bool inCodeBlock = false;
            var codeBlockLines = new List<string>();

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                // 1. Code Block Toggle (```)
                if (line.TrimStart().StartsWith("```"))
                {
                    if (inCodeBlock)
                    {
                        // End code block
                        var codeSpan = new Span
                        {
                            FontFamily = new FontFamily("Consolas, Courier New, monospace"),
                            FontSize = 11.5,
                            Foreground = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)),
                            Background = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9))
                        };
                        string codeContent = string.Join("\n", codeBlockLines.ToArray());
                        codeSpan.Inlines.Add(new Run(codeContent));
                        textBlock.Inlines.Add(codeSpan);
                        textBlock.Inlines.Add(new LineBreak());

                        codeBlockLines.Clear();
                        inCodeBlock = false;
                    }
                    else
                    {
                        // Start code block
                        inCodeBlock = true;
                        codeBlockLines.Clear();
                    }
                    continue;
                }

                if (inCodeBlock)
                {
                    codeBlockLines.Add(line);
                    continue;
                }

                // 2. Empty Line -> Paragraph spacing
                if (string.IsNullOrWhiteSpace(line))
                {
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }

                string trimmed = line.TrimStart();

                // 3. Headings (# H1, ## H2, ### H3, #### H4)
                if (trimmed.StartsWith("#### "))
                {
                    var headerSpan = new Span { FontWeight = FontWeights.Bold, FontSize = 13.0, Foreground = new SolidColorBrush(Color.FromRgb(0x1E, 0x29, 0x3B)) };
                    ParseInlineFormatting(headerSpan, trimmed.Substring(5));
                    textBlock.Inlines.Add(headerSpan);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }
                if (trimmed.StartsWith("### "))
                {
                    var headerSpan = new Span { FontWeight = FontWeights.Bold, FontSize = 13.5, Foreground = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)) };
                    ParseInlineFormatting(headerSpan, trimmed.Substring(4));
                    textBlock.Inlines.Add(headerSpan);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }
                if (trimmed.StartsWith("## "))
                {
                    var headerSpan = new Span { FontWeight = FontWeights.Bold, FontSize = 14.5, Foreground = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)) };
                    ParseInlineFormatting(headerSpan, trimmed.Substring(3));
                    textBlock.Inlines.Add(headerSpan);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }
                if (trimmed.StartsWith("# "))
                {
                    var headerSpan = new Span { FontWeight = FontWeights.Bold, FontSize = 16.0, Foreground = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)) };
                    ParseInlineFormatting(headerSpan, trimmed.Substring(2));
                    textBlock.Inlines.Add(headerSpan);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }

                // 4. Bullet Lists (- item, * item, + item)
                if (trimmed.StartsWith("- ") || trimmed.StartsWith("* ") || trimmed.StartsWith("+ "))
                {
                    var bulletRun = new Run(" •  ") { FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xEB)) };
                    textBlock.Inlines.Add(bulletRun);
                    ParseInlineFormatting(textBlock, trimmed.Substring(2));
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }

                // 5. Numbered Lists (1. item, 2. item, etc.)
                Match numMatch = Regex.Match(trimmed, @"^(\d+)\.\s+(.*)$");
                if (numMatch.Success)
                {
                    string num = numMatch.Groups[1].Value;
                    string rest = numMatch.Groups[2].Value;
                    var numRun = new Run(string.Format(" {0}. ", num)) { FontWeight = FontWeights.Bold, Foreground = new SolidColorBrush(Color.FromRgb(0x25, 0x63, 0xEB)) };
                    textBlock.Inlines.Add(numRun);
                    ParseInlineFormatting(textBlock, rest);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }

                // 6. Blockquote (> quote)
                if (trimmed.StartsWith("> "))
                {
                    var quoteSpan = new Span
                    {
                        FontStyle = FontStyles.Italic,
                        Foreground = new SolidColorBrush(Color.FromRgb(0x47, 0x55, 0x69))
                    };
                    quoteSpan.Inlines.Add(new Run("▎ "));
                    ParseInlineFormatting(quoteSpan, trimmed.Substring(2));
                    textBlock.Inlines.Add(quoteSpan);
                    textBlock.Inlines.Add(new LineBreak());
                    continue;
                }

                // 7. Regular paragraph text with inline formatting
                ParseInlineFormatting(textBlock, line);

                // Add line break if not the last line
                if (i < lines.Length - 1)
                {
                    textBlock.Inlines.Add(new LineBreak());
                }
            }

            // Handle unclosed code block
            if (inCodeBlock && codeBlockLines.Count > 0)
            {
                var codeSpan = new Span
                {
                    FontFamily = new FontFamily("Consolas, Courier New, monospace"),
                    FontSize = 11.5,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x0F, 0x17, 0x2A)),
                    Background = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9))
                };
                codeSpan.Inlines.Add(new Run(string.Join("\n", codeBlockLines.ToArray())));
                textBlock.Inlines.Add(codeSpan);
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
                        Background = new SolidColorBrush(Color.FromRgb(0xF1, 0xF5, 0xF9)),
                        Foreground = new SolidColorBrush(Color.FromRgb(0xB4, 0x53, 0x09)) // Amber-700
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
                // 5. Plain Text
                else
                {
                    targetInlines.Add(new Run(part));
                }
            }
        }
    }
}
