using System;
using System.Collections.Generic;

namespace MSOfficeAIAssistant.Core
{
    public enum DiffOp { Equal, Insert, Delete }

    public class DiffPiece
    {
        public DiffOp Operation { get; set; }
        public string Text { get; set; }
        public int OldLine { get; set; }
        public int NewLine { get; set; }
    }

    /// <summary>
    /// Simple line-based diff using LCS (Myers-like). COM-free, deterministic.
    /// Used for document comparison viewer (Slice 5) on extracted text.
    /// </summary>
    public static class DiffEngine
    {
        /// <summary>
        /// The LCS table below is O(m*n) time and memory. Above this many lines per side, the DP
        /// table alone would run into the multi-gigabyte / billion-iteration range and hang or OOM
        /// the caller (originally the WPF UI thread in DocumentCompareWindow). DiffLinesOrNull
        /// returns null past this bound instead of attempting the full LCS.
        /// </summary>
        public const int MaxLinesPerSide = 6000;

        /// <summary>
        /// Same as <see cref="DiffLines"/> but returns null instead of attempting an O(m*n) LCS when
        /// either side exceeds <see cref="MaxLinesPerSide"/> lines. Callers on a UI thread should use
        /// this (or run DiffLines on a background thread) rather than DiffLines directly.
        /// </summary>
        public static List<DiffPiece> DiffLinesOrNull(string oldText, string newText)
        {
            int m, n;
            CountLines(oldText, out m);
            CountLines(newText, out n);
            if (m > MaxLinesPerSide || n > MaxLinesPerSide) return null;
            return DiffLines(oldText, newText);
        }

        private static void CountLines(string text, out int count)
        {
            if (string.IsNullOrEmpty(text)) { count = 0; return; }
            string normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
            count = normalized.Split('\n').Length;
        }

        public static List<DiffPiece> DiffLines(string oldText, string newText)
        {
            var result = new List<DiffPiece>();
            if (oldText == null) oldText = string.Empty;
            if (newText == null) newText = string.Empty;

            string[] oldLines = oldText.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            string[] newLines = newText.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');

            // Handle empty strings (Split returns 1 element "" for empty)
            if (oldText.Length == 0) oldLines = new string[0];
            if (newText.Length == 0) newLines = new string[0];

            int m = oldLines.Length;
            int n = newLines.Length;
            int[,] dp = new int[m + 1, n + 1];

            for (int i = m - 1; i >= 0; i--)
            {
                for (int j = n - 1; j >= 0; j--)
                {
                    if (oldLines[i] == newLines[j])
                        dp[i, j] = dp[i + 1, j + 1] + 1;
                    else
                        dp[i, j] = dp[i + 1, j] >= dp[i, j + 1] ? dp[i + 1, j] : dp[i, j + 1];
                }
            }

            int oi = 0, ni = 0;
            int oldLineNo = 1, newLineNo = 1;
            while (oi < m && ni < n)
            {
                if (oldLines[oi] == newLines[ni])
                {
                    result.Add(new DiffPiece { Operation = DiffOp.Equal, Text = oldLines[oi], OldLine = oldLineNo, NewLine = newLineNo });
                    oi++; ni++; oldLineNo++; newLineNo++;
                }
                else if (dp[oi + 1, ni] >= dp[oi, ni + 1])
                {
                    result.Add(new DiffPiece { Operation = DiffOp.Delete, Text = oldLines[oi], OldLine = oldLineNo, NewLine = -1 });
                    oi++; oldLineNo++;
                }
                else
                {
                    result.Add(new DiffPiece { Operation = DiffOp.Insert, Text = newLines[ni], OldLine = -1, NewLine = newLineNo });
                    ni++; newLineNo++;
                }
            }
            while (oi < m)
            {
                result.Add(new DiffPiece { Operation = DiffOp.Delete, Text = oldLines[oi], OldLine = oldLineNo, NewLine = -1 });
                oi++; oldLineNo++;
            }
            while (ni < n)
            {
                result.Add(new DiffPiece { Operation = DiffOp.Insert, Text = newLines[ni], OldLine = -1, NewLine = newLineNo });
                ni++; newLineNo++;
            }
            return result;
        }

        public static string RenderPlain(List<DiffPiece> pieces)
        {
            if (pieces == null) return string.Empty;
            var sb = new System.Text.StringBuilder();
            foreach (DiffPiece p in pieces)
            {
                string prefix = p.Operation == DiffOp.Equal ? "  " : (p.Operation == DiffOp.Insert ? "+ " : "- ");
                sb.Append(prefix).Append(p.Text).AppendLine();
            }
            return sb.ToString();
        }
    }
}
