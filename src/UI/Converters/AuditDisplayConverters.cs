using System;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Data;
using MSOfficeAIAssistant.UI.Helpers;

namespace MSOfficeAIAssistant.UI.Converters
{
    /// <summary>
    /// Shared plain-text normalisation for the audit log. Entries store the raw response that was
    /// applied, so the log was rendering "**Details of Required Pump:**" and table pipes verbatim.
    /// </summary>
    internal static class AuditText
    {
        public static string ToPlain(object value)
        {
            string text = value as string;
            if (string.IsNullOrEmpty(text)) return string.Empty;

            string plain = MarkdownClipboard.ToPlainText(text).Replace("\r\n", "\n").Replace("\r", "\n");

            // Collapse runs of blank lines so a log card stays compact.
            var sb = new StringBuilder(plain.Length);
            int blanks = 0;
            string[] lines = plain.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                bool blank = string.IsNullOrWhiteSpace(lines[i]);
                if (blank)
                {
                    blanks++;
                    if (blanks > 1) continue;
                }
                else blanks = 0;

                if (sb.Length > 0) sb.Append(Environment.NewLine);
                sb.Append(lines[i].TrimEnd());
            }
            return sb.ToString().Trim();
        }
    }

    /// <summary>Renders a stored response as readable text, with Markdown markers removed.</summary>
    public class MarkdownPlainTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return AuditText.ToPlain(value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>
    /// First few lines of an entry, for the always-visible part of a log card. The full text
    /// stays one click away in the expander rather than filling the window.
    /// </summary>
    public class PreviewTextConverter : IValueConverter
    {
        public int MaxLines { get; set; }
        public int MaxCharacters { get; set; }

        public PreviewTextConverter()
        {
            MaxLines = 3;
            MaxCharacters = 220;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string plain = AuditText.ToPlain(value);
            if (plain.Length == 0) return string.Empty;

            string[] lines = plain.Replace("\r\n", "\n").Split('\n');
            var sb = new StringBuilder();
            int used = 0;
            bool clipped = false;

            for (int i = 0; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                if (used >= MaxLines) { clipped = true; break; }
                if (sb.Length > 0) sb.Append(Environment.NewLine);
                sb.Append(lines[i].Trim());
                used++;
            }

            string preview = sb.ToString();
            if (preview.Length > MaxCharacters)
            {
                preview = preview.Substring(0, MaxCharacters).TrimEnd();
                clipped = true;
            }
            else if (!clipped && preview.Length < plain.Length)
            {
                clipped = true;
            }

            return clipped ? preview + "\u2026" : preview;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>Shows the "full content" expander only when there is more than the preview shows.</summary>
    public class LongTextToVisibilityConverter : IValueConverter
    {
        public int Threshold { get; set; }

        public LongTextToVisibilityConverter()
        {
            Threshold = 220;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            string plain = AuditText.ToPlain(value);
            if (plain.Length > Threshold) return Visibility.Visible;
            return plain.IndexOf('\n') >= 0 ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    /// <summary>Turns the Undoable flag into a label, instead of binding a bare "True" into the card.</summary>
    public class UndoableTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool undoable = (value is bool) && (bool)value;
            return undoable ? "Undoable" : "Not undoable";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
