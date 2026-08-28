using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using MSOfficeAIAssistant.Core;

namespace MSOfficeAIAssistant.UI.Helpers
{
    /// <summary>
    /// Filesystem helper for @-mention picker. Enumerates local candidates from
    /// Documents and the current document's directory, then filters via MentionResolver.
    /// No cloud calls. Defensive: never throws.
    /// </summary>
    public static class MentionPicker
    {
        private const int MaxFilesPerDirectory = 200;

        public static List<string> CollectCandidates(string query, string currentDocumentPath)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var allCandidates = new List<string>();

            // 1. Current document directory (highest relevance)
            TryAddDirectory(currentDocumentPath, allCandidates, seen);

            // 2. User's Documents folder
            try
            {
                string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (!string.IsNullOrWhiteSpace(docs) && Directory.Exists(docs))
                {
                    TryAddDirectory(docs, allCandidates, seen);
                }
            }
            catch { }

            // 3. Desktop as fallback (often contains working files)
            try
            {
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                if (!string.IsNullOrWhiteSpace(desktop) && Directory.Exists(desktop))
                {
                    TryAddDirectory(desktop, allCandidates, seen);
                }
            }
            catch { }

            return MentionResolver.FilterCandidates(allCandidates, query, MentionResolver.MaxResults);
        }

        public static List<string> CollectCandidates(string query)
        {
            return CollectCandidates(query, null);
        }

        private static void TryAddDirectory(string directoryOrFilePath, List<string> target, HashSet<string> seen)
        {
            if (string.IsNullOrWhiteSpace(directoryOrFilePath)) return;
            string dir = directoryOrFilePath;
            try
            {
                // If a file path was supplied, use its directory
                if (File.Exists(dir))
                {
                    dir = Path.GetDirectoryName(dir);
                }
                if (string.IsNullOrWhiteSpace(dir)) return;
                if (!Directory.Exists(dir)) return;

                string[] files = Directory.GetFiles(dir);
                int added = 0;
                foreach (string file in files)
                {
                    if (added >= MaxFilesPerDirectory) break;
                    if (string.IsNullOrWhiteSpace(file)) continue;
                    if (seen.Contains(file)) continue;
                    if (!MentionResolver.IsSupportedFile(file)) continue;

                    // Skip very large files (>30 MB) early to avoid wasted filtering
                    try
                    {
                        var fi = new FileInfo(file);
                        if (fi.Length > 30L * 1024 * 1024) continue;
                    }
                    catch { }

                    seen.Add(file);
                    target.Add(file);
                    added++;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("MentionPicker.TryAddDirectory failed for '{0}': {1}", dir, ex.Message));
            }
        }
    }
}
