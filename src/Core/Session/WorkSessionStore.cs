using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core.Session
{
    /// <summary>
    /// Manages persistent storage of WorkSession objects with DPAPI encryption.
    /// Each WorkSession is stored in a separate encrypted file named by WorkSessionId.
    /// Supports querying sessions by DocumentKey for "reopen recent session" UI workflows.
    /// </summary>
    public class WorkSessionStore
    {
        private static readonly object _lock = new object();
        private static WorkSessionStore _instance;
        private readonly Dictionary<string, WorkSession> _sessionCache = new Dictionary<string, WorkSession>(StringComparer.OrdinalIgnoreCase);
        private readonly string _storageDir;

        public static WorkSessionStore Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new WorkSessionStore();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Production constructor - uses AppPaths.InDataDirectory("WorkSessions") for storage.
        /// </summary>
        public WorkSessionStore()
        {
            _storageDir = AppPaths.InDataDirectory("WorkSessions");
            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    Directory.CreateDirectory(_storageDir);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkSessionStore could not create directory '{0}': {1}", _storageDir, ex.Message));
            }
        }

        /// <summary>
        /// Test constructor - accepts a custom storage directory for headless testing.
        /// </summary>
        public WorkSessionStore(string customDir)
        {
            _storageDir = customDir ?? AppPaths.InDataDirectory("WorkSessions");
            try
            {
                if (!Directory.Exists(_storageDir))
                {
                    Directory.CreateDirectory(_storageDir);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkSessionStore could not create directory '{0}': {1}", _storageDir, ex.Message));
            }
        }

        /// <summary>
        /// Saves a WorkSession to encrypted storage.
        /// Updates the session's UpdatedUtc timestamp before saving.
        /// </summary>
        public void Save(WorkSession session)
        {
            if (session == null) return;

            try
            {
                lock (_lock)
                {
                    session.UpdatedUtc = DateTime.UtcNow;
                    string filePath = GetFilePath(session.WorkSessionId);
                    string json = JsonConvert.SerializeObject(session, Formatting.Indented);
                    byte[] plainBytes = Encoding.UTF8.GetBytes(json);
                    byte[] encryptedBytes = ProtectedData.Protect(plainBytes, null, DataProtectionScope.CurrentUser);
                    File.WriteAllBytes(filePath, encryptedBytes);

                    // Update cache
                    _sessionCache[session.WorkSessionId] = new WorkSession
                    {
                        WorkSessionId = session.WorkSessionId,
                        DocumentKey = session.DocumentKey,
                        Title = session.Title,
                        CreatedUtc = session.CreatedUtc,
                        UpdatedUtc = session.UpdatedUtc,
                        Status = session.Status,
                        SourceHosts = session.SourceHosts != null ? new List<string>(session.SourceHosts) : new List<string>()
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("Failed to save work session '{0}': {1}", session.WorkSessionId, ex.Message));
            }
        }

        /// <summary>
        /// Loads a WorkSession from encrypted storage by WorkSessionId.
        /// Returns null if the session is not found or fails to deserialize.
        /// Corrupted files are quarantined to .bak and logged as warnings.
        /// </summary>
        public WorkSession Load(string workSessionId)
        {
            if (string.IsNullOrWhiteSpace(workSessionId)) return null;

            lock (_lock)
            {
                // Check cache first
                WorkSession cached;
                if (_sessionCache.TryGetValue(workSessionId, out cached))
                {
                    // Reload from disk to get full state including Plan
                    return LoadFromDisk(workSessionId);
                }

                return LoadFromDisk(workSessionId);
            }
        }

        /// <summary>
        /// Lists all WorkSessions for a given DocumentKey, most-recent-first.
        /// Used for "reopen recent session" UI workflows.
        /// </summary>
        public List<WorkSession> ListByDocumentKey(string documentKey)
        {
            if (string.IsNullOrWhiteSpace(documentKey))
            {
                documentKey = "GlobalSession";
            }

            lock (_lock)
            {
                List<WorkSession> result = new List<WorkSession>();
                if (!Directory.Exists(_storageDir))
                {
                    return result;
                }

                try
                {
                    string[] files = Directory.GetFiles(_storageDir, "*.dat");
                    foreach (string filePath in files)
                    {
                        try
                        {
                            string sessionId = Path.GetFileNameWithoutExtension(filePath);
                            WorkSession session = LoadFromDisk(sessionId);
                            if (session != null && string.Equals(session.DocumentKey, documentKey, StringComparison.OrdinalIgnoreCase))
                            {
                                result.Add(session);
                            }
                        }
                        catch
                        {
                            // Skip individual files that fail to load
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("Failed to list work sessions for document '{0}': {1}", documentKey, ex.Message));
                }

                // Sort by UpdatedUtc descending (most recent first)
                result.Sort((a, b) => b.UpdatedUtc.CompareTo(a.UpdatedUtc));
                return result;
            }
        }

        /// <summary>
        /// Deletes a WorkSession from storage.
        /// </summary>
        public void Delete(string workSessionId)
        {
            if (string.IsNullOrWhiteSpace(workSessionId)) return;

            lock (_lock)
            {
                try
                {
                    string filePath = GetFilePath(workSessionId);
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                    _sessionCache.Remove(workSessionId);
                }
                catch (Exception ex)
                {
                    Logger.Warn(string.Format("Failed to delete work session '{0}': {1}", workSessionId, ex.Message));
                }
            }
        }

        /// <summary>
        /// Clears the in-memory cache.
        /// Used for test isolation between test runs.
        /// </summary>
        public void ClearMemoryCache()
        {
            lock (_lock)
            {
                _sessionCache.Clear();
            }
        }

        private WorkSession LoadFromDisk(string workSessionId)
        {
            try
            {
                string filePath = GetFilePath(workSessionId);
                if (!File.Exists(filePath))
                {
                    return null;
                }

                byte[] encryptedBytes = File.ReadAllBytes(filePath);
                byte[] plainBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                string json = Encoding.UTF8.GetString(plainBytes);
                WorkSession session = JsonConvert.DeserializeObject<WorkSession>(json);
                return session;
            }
            catch (CryptographicException ex)
            {
                // DPAPI failure (profile/machine change, elevation mismatch) — quarantine, do not overwrite
                Logger.Warn(string.Format("WorkSessionStore: DPAPI decryption failed for '{0}' ({1}). Existing session quarantined as .bak. Session will not be recoverable.", workSessionId, ex.Message));
                try
                {
                    string filePath = GetFilePath(workSessionId);
                    string bakPath = filePath + ".bak";
                    if (File.Exists(bakPath)) File.Delete(bakPath);
                    if (File.Exists(filePath)) File.Move(filePath, bakPath);
                }
                catch (Exception moveEx)
                {
                    Logger.Warn(string.Format("WorkSessionStore: Could not quarantine corrupt file for '{0}': {1}", workSessionId, moveEx.Message));
                }
                return null;
            }
            catch (Exception ex)
            {
                // JSON deserialization or other errors — quarantine the file
                Logger.Warn(string.Format("WorkSessionStore: Failed to load work session '{0}' ({1}). Session file quarantined as .bak.", workSessionId, ex.Message));
                try
                {
                    string filePath = GetFilePath(workSessionId);
                    string bakPath = filePath + ".bak";
                    if (File.Exists(bakPath)) File.Delete(bakPath);
                    if (File.Exists(filePath)) File.Move(filePath, bakPath);
                }
                catch (Exception moveEx)
                {
                    Logger.Warn(string.Format("WorkSessionStore: Could not quarantine corrupt file for '{0}': {1}", workSessionId, moveEx.Message));
                }
                return null;
            }
        }

        private string GetFilePath(string workSessionId)
        {
            return Path.Combine(_storageDir, string.Format("{0}.dat", workSessionId));
        }
    }
}
