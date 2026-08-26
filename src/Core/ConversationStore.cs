using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using MSOfficeAIAssistant.API.Models;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core
{
    public class ConversationSessionSummary
    {
        public string DocumentKey { get; set; }
        public string Title { get; set; }
        public int MessageCount { get; set; }
        public DateTime LastUpdatedUtc { get; set; }
    }

    public class ConversationStore
    {
        private static readonly object _lock = new object();
        private static ConversationStore _instance;
        private readonly Dictionary<string, List<ChatMessage>> _sessionHistories = new Dictionary<string, List<ChatMessage>>(StringComparer.OrdinalIgnoreCase);
        private readonly string _storageDir;

        public static ConversationStore Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new ConversationStore();
                        }
                    }
                }
                return _instance;
            }
        }

        public ConversationStore()
        {
            _storageDir = AppPaths.InDataDirectory("Conversations");
            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    Directory.CreateDirectory(_storageDir);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ConversationStore could not create directory '{0}': {1}", _storageDir, ex.Message));
            }
        }

        public ConversationStore(string customDir)
        {
            _storageDir = customDir ?? AppPaths.InDataDirectory("Conversations");
            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    Directory.CreateDirectory(_storageDir);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ConversationStore could not create directory '{0}': {1}", _storageDir, ex.Message));
            }
        }

        public void ClearMemoryCache()
        {
            lock (_lock)
            {
                _sessionHistories.Clear();
            }
        }

        public List<ChatMessage> GetHistory(string documentKey)
        {
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                documentKey = "GlobalSession";
            }

            lock (_lock)
            {
                List<ChatMessage> history;
                if (_sessionHistories.TryGetValue(documentKey, out history))
                {
                    return history;
                }

                var loaded = LoadFromDisk(documentKey);
                _sessionHistories[documentKey] = loaded;
                return loaded;
            }
        }

        public void SaveHistory(string documentKey, List<ChatMessage> history)
        {
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                documentKey = "GlobalSession";
            }

            lock (_lock)
            {
                _sessionHistories[documentKey] = new List<ChatMessage>(history);
                SaveToDisk(documentKey, history);
            }
        }

        public void ClearHistory(string documentKey)
        {
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                documentKey = "GlobalSession";
            }

            lock (_lock)
            {
                _sessionHistories.Remove(documentKey);
                try
                {
                    string safeKey = GetSafeFilename(documentKey);
                    string filePath = Path.Combine(_storageDir, string.Format("{0}.dat", safeKey));
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }

                    // Clean up histories saved by releases before encrypted persistence.
                    string legacyFilePath = Path.Combine(_storageDir, string.Format("{0}.json", safeKey));
                    if (File.Exists(legacyFilePath))
                    {
                        File.Delete(legacyFilePath);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("ConversationStore could not delete history file for '{0}': {1}", documentKey, ex.Message));
                }
            }
        }

        public List<ConversationSessionSummary> ListSessions()
        {
            var summaries = new List<ConversationSessionSummary>();

            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    return summaries;
                }

                string[] datFiles = Directory.GetFiles(_storageDir, "*.dat");
                foreach (string filePath in datFiles)
                {
                    try
                    {
                        string filename = Path.GetFileNameWithoutExtension(filePath);
                        string documentKey = filename;

                        var history = LoadFromDisk(documentKey);
                        if (history == null || history.Count == 0)
                        {
                            continue;
                        }

                        string title = documentKey;
                        var firstUserMsg = history.FirstOrDefault(m => m.IsUser);
                        if (firstUserMsg != null && !string.IsNullOrWhiteSpace(firstUserMsg.Content))
                        {
                            string content = firstUserMsg.Content.Trim();
                            if (content.Length > 60)
                            {
                                title = content.Substring(0, 60) + "…";
                            }
                            else
                            {
                                title = content;
                            }
                        }

                        DateTime lastUpdated = File.GetLastWriteTimeUtc(filePath);

                        summaries.Add(new ConversationSessionSummary
                        {
                            DocumentKey = documentKey,
                            Title = title,
                            MessageCount = history.Count,
                            LastUpdatedUtc = lastUpdated
                        });
                    }
                    catch (Exception ex)
                    {
                        Logger.Warn(string.Format("ConversationStore.ListSessions could not process file '{0}': {1}", filePath, ex.Message));
                    }
                }

                summaries = summaries.OrderByDescending(s => s.LastUpdatedUtc).ToList();
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("ConversationStore.ListSessions failed: {0}", ex.Message));
            }

            return summaries;
        }

        private List<ChatMessage> LoadFromDisk(string documentKey)
        {
            try
            {
                string safeKey = GetSafeFilename(documentKey);
                string filePath = Path.Combine(_storageDir, string.Format("{0}.dat", safeKey));
                if (File.Exists(filePath))
                {
                    byte[] encryptedBytes = File.ReadAllBytes(filePath);
                    byte[] plainBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                    string json = Encoding.UTF8.GetString(plainBytes);
                    var list = JsonConvert.DeserializeObject<List<ChatMessage>>(json);
                    if (list != null) return list;
                }

                // One-time, best-effort migration of plaintext histories created by v0.3.
                string legacyFilePath = Path.Combine(_storageDir, string.Format("{0}.json", safeKey));
                if (File.Exists(legacyFilePath))
                {
                    string legacyJson = File.ReadAllText(legacyFilePath);
                    var legacyList = JsonConvert.DeserializeObject<List<ChatMessage>>(legacyJson);
                    if (legacyList != null)
                    {
                        SaveToDisk(documentKey, legacyList);
                        try { File.Delete(legacyFilePath); } catch { }
                        return legacyList;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to load conversation for {0}: {1}", documentKey, ex.Message));
            }
            return new List<ChatMessage>();
        }

        private void SaveToDisk(string documentKey, List<ChatMessage> history)
        {
            try
            {
                string safeKey = GetSafeFilename(documentKey);
                string filePath = Path.Combine(_storageDir, string.Format("{0}.dat", safeKey));
                string json = JsonConvert.SerializeObject(history, Formatting.Indented);
                byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                File.WriteAllBytes(filePath, encryptedBytes);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to persist conversation for {0}: {1}", documentKey, ex.Message));
            }
        }

        private string GetSafeFilename(string key)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                key = key.Replace(c, '_');
            }
            return key;
        }
    }
}
