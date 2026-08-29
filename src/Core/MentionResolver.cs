using System;
using System.Collections.Generic;
using System.IO;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Pure, COM-free helper for @-mention parsing and candidate filtering.
    /// No filesystem or WPF dependencies, fully unit-testable.
    /// </summary>
    public static class MentionResolver
    {
        public const int MaxResults = 10;

        private static readonly HashSet<string> SupportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".docx", ".xlsx", ".pptx", ".pdf",
            ".png", ".jpg", ".jpeg", ".webp", ".gif", ".bmp",
            ".txt", ".csv", ".json", ".xml", ".md", ".cs", ".py", ".js", ".html", ".css", ".sql", ".log", ".ini", ".yaml", ".yml"
        };

        public static bool IsSupportedExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension)) return false;
            string ext = extension.Trim();
            if (!ext.StartsWith(".")) ext = "." + ext;
            return SupportedExtensions.Contains(ext);
        }

        public static bool IsSupportedFile(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return false;
            string ext = Path.GetExtension(fileName);
            return IsSupportedExtension(ext);
        }

        /// <summary>
        /// Attempts to extract the @-mention query at the caret position.
        /// Returns true when an '@' trigger is found before caretIndex and the text
        /// after it (up to caret) is a valid query (letters, digits, dot, dash, underscore).
        /// Query may be empty (bare '@' to show all).
        /// </summary>
        public static bool TryExtractQuery(string text, int caretIndex, out string query)
        {
            query = null;
            if (text == null) return false;
            if (caretIndex < 0 || caretIndex > text.Length) return false;
            if (caretIndex == 0) return false;

            // Search backwards from caret-1 for the nearest '@'
            int atPos = -1;
            for (int i = caretIndex - 1; i >= 0; i--)
            {
                char c = text[i];
                if (c == '@')
                {
                    atPos = i;
                    break;
                }
                if (c == ' ' || c == '\n' || c == '\r' || c == '\t')
                {
                    // A whitespace break before '@' means no active mention
                    break;
                }
                // If we hit a char that is not a valid query char and not '@', keep scanning
                // but if we encounter a non-query char other than '@', the mention is still valid
                // as long as '@' is found before whitespace - so we continue scanning
            }

            if (atPos < 0) return false;

            // Ensure '@' is at start or preceded by whitespace / newline / '(' / '"' / "'"
            if (atPos > 0)
            {
                char prev = text[atPos - 1];
                bool allowedBefore = prev == ' ' || prev == '\n' || prev == '\r' || prev == '\t' || prev == '(' || prev == '"' || prev == '\'' || prev == '[';
                if (!allowedBefore)
                {
                    // email-like pattern a@b should not trigger; require separator before @
                    // However we still allow trigger if @ is the first char of the string
                    return false;
                }
            }

            // Extract query substring between @ and caret
            int queryStart = atPos + 1;
            int queryLength = caretIndex - queryStart;
            if (queryLength < 0) return false;
            string raw = text.Substring(queryStart, queryLength);

            // Validate query characters: allow letters, digits, dot, dash, underscore
            for (int i = 0; i < raw.Length; i++)
            {
                char c = raw[i];
                bool valid = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9') || c == '.' || c == '-' || c == '_';
                if (!valid) return false;
            }

            query = raw;
            return true;
        }

        /// <summary>
        /// Filters candidate file names/paths by query substring (case-insensitive).
        /// Ranks starts-with higher than contains, then alphabetically.
        /// Caps to maxResults.
        /// </summary>
        public static List<string> FilterCandidates(IEnumerable<string> candidates, string query, int maxResults)
        {
            if (maxResults <= 0) maxResults = MaxResults;
            var result = new List<string>();
            if (candidates == null) return result;

            string normalizedQuery = query != null ? query.Trim().ToLowerInvariant() : string.Empty;
            bool hasQuery = !string.IsNullOrEmpty(normalizedQuery);

            // Separate into two buckets: starts-with and contains
            List<string> startsWith = new List<string>();
            List<string> contains = new List<string>();

            foreach (string candidate in candidates)
            {
                if (string.IsNullOrWhiteSpace(candidate)) continue;
                if (!IsSupportedFile(candidate)) continue;

                string fileName = Path.GetFileName(candidate);
                if (string.IsNullOrEmpty(fileName)) fileName = candidate;
                string lowerName = fileName.ToLowerInvariant();

                if (!hasQuery)
                {
                    // No query -> include all up to cap, sorted below
                    contains.Add(candidate);
                    continue;
                }

                if (lowerName.StartsWith(normalizedQuery, StringComparison.Ordinal))
                {
                    startsWith.Add(candidate);
                }
                else if (lowerName.IndexOf(normalizedQuery, StringComparison.Ordinal) >= 0)
                {
                    contains.Add(candidate);
                }
            }

            // Sort each bucket alphabetically (case-insensitive)
            startsWith.Sort(StringComparer.OrdinalIgnoreCase);
            contains.Sort(StringComparer.OrdinalIgnoreCase);

            foreach (string s in startsWith)
            {
                if (result.Count >= maxResults) break;
                result.Add(s);
            }
            foreach (string s in contains)
            {
                if (result.Count >= maxResults) break;
                result.Add(s);
            }

            // If query empty, result is currently unsorted contains list - sort it alphabetically
            // and already added in sorted order, but we built contains sorted above so ok
            return result;
        }

        public static List<string> FilterCandidates(IEnumerable<string> candidates, string query)
        {
            return FilterCandidates(candidates, query, MaxResults);
        }
    }
}
