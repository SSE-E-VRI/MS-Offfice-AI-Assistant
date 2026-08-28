using System;
using System.Collections.Generic;
using MSOfficeAIAssistant.API.Models;

namespace MSOfficeAIAssistant.UI.Helpers
{
    /// <summary>
    /// Splits a "3 Variants" response (RibbonCallback.OnRewriteVariants) on the
    /// ChatMessage.VariantDelimiter marker into individual candidate rewrites. Pure and COM-free
    /// so it's directly testable without a Word host (AI_Assistant_SSOT.md testability rule).
    /// </summary>
    public static class RewriteVariantParser
    {
        /// <summary>
        /// Returns the non-blank segments of <paramref name="content"/> split on lines that
        /// contain only the variant delimiter. Returns an empty list if content is null/blank;
        /// returns a single-item list (the whole trimmed content) if the delimiter never appears.
        /// </summary>
        public static List<string> Split(string content)
        {
            var result = new List<string>();
            if (string.IsNullOrWhiteSpace(content)) return result;

            string[] lines = content.Replace("\r\n", "\n").Split('\n');
            var current = new List<string>();

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].Trim() == ChatMessage.VariantDelimiter)
                {
                    AddIfNotBlank(result, current);
                    current = new List<string>();
                }
                else
                {
                    current.Add(lines[i]);
                }
            }
            AddIfNotBlank(result, current);

            return result;
        }

        private static void AddIfNotBlank(List<string> result, List<string> lines)
        {
            string joined = string.Join("\n", lines.ToArray()).Trim();
            if (!string.IsNullOrEmpty(joined)) result.Add(joined);
        }
    }
}
