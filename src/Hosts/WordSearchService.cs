using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MSOfficeAIAssistant.Hosts
{
    /// <summary>
    /// COM-free query builder for Word smart search. Converts natural language queries
    /// into regex or simple contains checks, testable without Office.
    /// </summary>
    public static class WordSearchService
    {
        private static readonly HashSet<string> StopWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "the","and","or","a","an","in","on","at","for","to","of","with","by","is","are","was","were","find","search","locate","show","where"
        };

        public static string BuildFindPattern(string naturalQuery)
        {
            if (string.IsNullOrWhiteSpace(naturalQuery)) return string.Empty;
            List<string> terms = ExtractTerms(naturalQuery);
            if (terms.Count == 0) return string.Empty;
            // For now, return the most significant term as search pattern.
            // Prefer longest term (most specific)
            string best = terms[0];
            foreach (string t in terms)
            {
                if (t.Length > best.Length) best = t;
            }
            return best;
        }

        public static List<string> ExtractTerms(string query)
        {
            var terms = new List<string>();
            if (string.IsNullOrWhiteSpace(query)) return terms;
            MatchCollection matches = Regex.Matches(query.ToLowerInvariant(), @"[\p{L}\p{Nd}_-]{3,}");
            foreach (Match m in matches)
            {
                string term = m.Value;
                if (StopWords.Contains(term)) continue;
                if (terms.Contains(term)) continue;
                terms.Add(term);
            }
            return terms;
        }

        public static bool MatchesParagraph(string paragraphText, string query)
        {
            if (string.IsNullOrWhiteSpace(paragraphText) || string.IsNullOrWhiteSpace(query)) return false;
            string lowerPara = paragraphText.ToLowerInvariant();
            string lowerQuery = query.Trim().ToLowerInvariant();
            if (lowerPara.IndexOf(lowerQuery, StringComparison.Ordinal) >= 0) return true;
            List<string> terms = ExtractTerms(query);
            foreach (string term in terms)
            {
                if (lowerPara.IndexOf(term, StringComparison.Ordinal) >= 0) return true;
            }
            return false;
        }

        public static string BuildCoachingPrompt(string selectedText)
        {
            if (string.IsNullOrWhiteSpace(selectedText))
                return "Review the document for clarity, gaps, ambiguous language, and consistency. Provide findings as summary, do not propose document changes.";
            return "Review the following text for clarity, gaps, ambiguous language, and consistency. Provide findings as summary (do NOT propose insertions or replacements; this is a read-only coaching review):\n\n\"\"\"\n" + selectedText + "\n\"\"\"";
        }
    }
}
