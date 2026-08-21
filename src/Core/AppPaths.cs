using System;
using System.IO;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Resolves the per-user data directory and migrates data written by builds that
    /// used the legacy "MistralOfficeAddin" folder name.
    ///
    /// This type must stay dependency-free: <see cref="Logger"/> calls it from its own
    /// static constructor, so using Logger here would create a initialization cycle.
    /// All failures are swallowed and fall back to the legacy path.
    /// </summary>
    public static class AppPaths
    {
        private const string FolderName = "MSOfficeAIAssistant";
        private const string LegacyFolderName = "MistralOfficeAddin";

        private static readonly object _sync = new object();
        private static string _dataDirectory;

        /// <summary>
        /// %LOCALAPPDATA%\MSOfficeAIAssistant, created on first access. Content written by
        /// older builds is moved across once, so DPAPI-encrypted credentials, conversation
        /// history and the action audit trail survive the rename.
        /// </summary>
        public static string DataDirectory
        {
            get
            {
                if (_dataDirectory != null) return _dataDirectory;

                lock (_sync)
                {
                    if (_dataDirectory != null) return _dataDirectory;

                    string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    string current = Path.Combine(localAppData, FolderName);
                    string legacy = Path.Combine(localAppData, LegacyFolderName);

                    try
                    {
                        MigrateLegacyData(legacy, current);
                        if (!Directory.Exists(current)) Directory.CreateDirectory(current);
                        _dataDirectory = current;
                    }
                    catch
                    {
                        // If the new location is unusable for any reason, keep using the
                        // legacy folder rather than losing access to existing settings.
                        _dataDirectory = Directory.Exists(legacy) ? legacy : current;
                    }

                    return _dataDirectory;
                }
            }
        }

        /// <summary>Returns a path inside <see cref="DataDirectory"/>.</summary>
        public static string InDataDirectory(string fileName)
        {
            return Path.Combine(DataDirectory, fileName);
        }

        private static void MigrateLegacyData(string legacy, string current)
        {
            if (!Directory.Exists(legacy)) return;

            // Fresh rename: nothing at the new location yet, so move the whole tree.
            if (!Directory.Exists(current))
            {
                try
                {
                    Directory.Move(legacy, current);
                    return;
                }
                catch
                {
                    // A locked file (for example an open log handle) blocks the move.
                    // Fall through and copy what we can instead.
                    Directory.CreateDirectory(current);
                }
            }

            CopyMissing(legacy, current);
        }

        /// <summary>
        /// Copies files the destination does not already have. Existing files are never
        /// overwritten, so a partially migrated profile converges without losing newer data.
        /// </summary>
        private static void CopyMissing(string sourceDir, string targetDir)
        {
            if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

            string[] files;
            try { files = Directory.GetFiles(sourceDir); }
            catch { return; }

            for (int i = 0; i < files.Length; i++)
            {
                try
                {
                    string destination = Path.Combine(targetDir, Path.GetFileName(files[i]));
                    if (!File.Exists(destination)) File.Copy(files[i], destination);
                }
                catch
                {
                    // Skip individual unreadable/locked files; migration is best-effort.
                }
            }

            string[] subDirectories;
            try { subDirectories = Directory.GetDirectories(sourceDir); }
            catch { return; }

            for (int i = 0; i < subDirectories.Length; i++)
            {
                try
                {
                    CopyMissing(subDirectories[i], Path.Combine(targetDir, Path.GetFileName(subDirectories[i])));
                }
                catch
                {
                }
            }
        }
    }
}
