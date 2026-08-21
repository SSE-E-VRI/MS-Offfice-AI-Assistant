using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Stores a compact, local, encrypted audit trail of user-approved AI actions.
    /// It intentionally records metadata and a bounded preview rather than full document content.
    /// </summary>
    public sealed class ActionAuditStore
    {
        private const int MaxEntries = 250;
        private const int MaxPreviewLength = 2000;
        private static readonly object SyncRoot = new object();
        private static ActionAuditStore _instance;
        private readonly string _filePath;

        public static ActionAuditStore Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (SyncRoot)
                    {
                        if (_instance == null) _instance = new ActionAuditStore();
                    }
                }
                return _instance;
            }
        }

        public ActionAuditStore()
        {
            _filePath = AppPaths.InDataDirectory("action-audit.dat");
        }

        public ActionAuditStore(string customFilePath)
        {
            _filePath = customFilePath;
            if (!string.IsNullOrEmpty(_filePath))
            {
                string dir = Path.GetDirectoryName(_filePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
            }
        }

        public void Record(string host, string actionType, string target, string summary, bool undoable)
        {
            Record(host, actionType, target, summary, undoable, null, null, null, null, null);
        }

        public void Record(
            string host,
            string actionType,
            string target,
            string summary,
            bool undoable,
            string prompt,
            string sourceContext,
            string model,
            string fullProposedAction,
            string applyResult)
        {
            try
            {
                lock (SyncRoot)
                {
                    List<ActionAuditEntry> entries = LoadUnsafe();
                    entries.Add(new ActionAuditEntry
                    {
                        TimestampUtc = DateTime.UtcNow,
                        Host = host ?? "Office",
                        ActionType = actionType ?? "Apply",
                        Target = target ?? string.Empty,
                        Summary = Truncate(summary),
                        Undoable = undoable,
                        Prompt = Truncate(prompt),
                        SourceContext = Truncate(sourceContext),
                        Model = model ?? string.Empty,
                        FullProposedAction = Truncate(fullProposedAction),
                        ApplyResult = Truncate(applyResult)
                    });

                    if (entries.Count > MaxEntries)
                    {
                        entries.RemoveRange(0, entries.Count - MaxEntries);
                    }
                    SaveUnsafe(entries);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Action audit could not be saved: {0}", ex.Message));
            }
        }

        public List<ActionAuditEntry> GetRecent(int maxEntries)
        {
            lock (SyncRoot)
            {
                List<ActionAuditEntry> entries = LoadUnsafe();
                entries.Reverse();
                if (maxEntries > 0 && entries.Count > maxEntries)
                    entries.RemoveRange(maxEntries, entries.Count - maxEntries);
                return entries;
            }
        }

        private List<ActionAuditEntry> LoadUnsafe()
        {
            if (!File.Exists(_filePath)) return new List<ActionAuditEntry>();
            try
            {
                byte[] encrypted = File.ReadAllBytes(_filePath);
                byte[] plain = ProtectedData.Unprotect(encrypted, null, DataProtectionScope.CurrentUser);
                string json = Encoding.UTF8.GetString(plain);
                return JsonConvert.DeserializeObject<List<ActionAuditEntry>>(json) ?? new List<ActionAuditEntry>();
            }
            catch (System.Security.Cryptography.CryptographicException ex)
            {
                // DPAPI failure (profile/machine change, elevation mismatch) — quarantine, do not overwrite
                Logger.Warn(string.Format("ActionAuditStore: DPAPI decryption failed ({0}). Existing audit log quarantined as .bak. History will restart.", ex.Message));
                try
                {
                    string bakPath = _filePath + ".bak";
                    if (File.Exists(bakPath)) File.Delete(bakPath);
                    File.Move(_filePath, bakPath);
                }
                catch (Exception moveEx)
                {
                    Logger.Warn(string.Format("ActionAuditStore: Could not quarantine corrupt file: {0}", moveEx.Message));
                }
                return new List<ActionAuditEntry>();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ActionAuditStore: Could not load audit log ({0}). Starting fresh.", ex.Message));
                return new List<ActionAuditEntry>();
            }
        }

        private void SaveUnsafe(List<ActionAuditEntry> entries)
        {
            string json = JsonConvert.SerializeObject(entries, Formatting.None);
            byte[] plain = Encoding.UTF8.GetBytes(json);
            byte[] encrypted = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(_filePath, encrypted);
        }

        private static string Truncate(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            return value.Length <= MaxPreviewLength ? value : value.Substring(0, MaxPreviewLength) + "...[truncated]";
        }
    }

    public class ActionAuditEntry
    {
        public DateTime TimestampUtc { get; set; }
        public string Host { get; set; }
        public string ActionType { get; set; }
        public string Target { get; set; }
        public string Summary { get; set; }
        public bool Undoable { get; set; }
        public string Prompt { get; set; }
        public string SourceContext { get; set; }
        public string Model { get; set; }
        public string FullProposedAction { get; set; }
        public string ApplyResult { get; set; }
    }
}
