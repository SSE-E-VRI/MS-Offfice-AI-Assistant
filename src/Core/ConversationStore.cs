using System;
using System.Collections.Generic;
using System.IO;
using MistralOfficeAddin.API.Models;
using Newtonsoft.Json;

namespace MistralOfficeAddin.Core
{
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
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            _storageDir = Path.Combine(localAppData, "MistralOfficeAddin", "Conversations");
            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    Directory.CreateDirectory(_storageDir);
                }
            }
            catch { }
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
                    string filePath = Path.Combine(_storageDir, string.Format("{0}.json", safeKey));
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                }
                catch { }
            }
        }

        private List<ChatMessage> LoadFromDisk(string documentKey)
        {
            try
            {
                string safeKey = GetSafeFilename(documentKey);
                string filePath = Path.Combine(_storageDir, string.Format("{0}.json", safeKey));
                if (File.Exists(filePath))
                {
                    string json = File.ReadAllText(filePath);
                    var list = JsonConvert.DeserializeObject<List<ChatMessage>>(json);
                    if (list != null) return list;
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
                string filePath = Path.Combine(_storageDir, string.Format("{0}.json", safeKey));
                string json = JsonConvert.SerializeObject(history, Formatting.Indented);
                File.WriteAllText(filePath, json);
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
