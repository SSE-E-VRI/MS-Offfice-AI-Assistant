using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using MSOfficeAIAssistant.Hosts;
using Newtonsoft.Json;

namespace MSOfficeAIAssistant.Core
{
    /// <summary>
    /// Local per-workbook rules / personalization. No cloud, no Graph.
    /// Sources (in priority):
    /// 1) Hidden worksheet named ".Rules" (if present) — read via ExcelController on UI thread, injected per-turn.
    /// 2) Local JSON file %LOCALAPPDATA%\MSOfficeAIAssistant\WorkbookRules\{safeName}.json (plain, not DPAPI — rules are not secrets)
    /// 3) Global profile %LOCALAPPDATA%\MSOfficeAIAssistant\profile.json (optional)
    /// All file I/O is best-effort; never throws to caller.
    /// </summary>
    public static class WorkbookRulesStore
    {
        private static string WorkbookRulesDir
        {
            get { return Path.Combine(AppPaths.InDataDirectory("WorkbookRules")); }
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "DefaultWorkbook";
            string safe = name.Trim();
            foreach (char c in Path.GetInvalidFileNameChars()) safe = safe.Replace(c, '_');
            safe = safe.Replace(" ", "_");
            if (safe.Length > 64) safe = safe.Substring(0, 64);
            return safe;
        }

        public static string GetRulesFilePath(string workbookName)
        {
            string dir = WorkbookRulesDir;
            try { Directory.CreateDirectory(dir); } catch { }
            string safe = SanitizeFileName(workbookName);
            return Path.Combine(dir, safe + ".json");
        }

        public static string GetGlobalProfilePath()
        {
            string dir = WorkbookRulesDir;
            try { Directory.CreateDirectory(dir); } catch { }
            return Path.Combine(dir, "_global_profile.json");
        }

        public static Dictionary<string, string> LoadRules(string workbookName)
        {
            try
            {
                string path = GetRulesFilePath(workbookName);
                if (!File.Exists(path)) return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                string json = File.ReadAllText(path, Encoding.UTF8);
                var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                return dict != null ? new Dictionary<string, string>(dict, StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkbookRulesStore.LoadRules failed for {0}: {1}", workbookName, ex.Message));
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public static Dictionary<string, string> LoadGlobalProfile()
        {
            try
            {
                string path = GetGlobalProfilePath();
                if (!File.Exists(path)) return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                string json = File.ReadAllText(path, Encoding.UTF8);
                var dict = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                return dict != null ? new Dictionary<string, string>(dict, StringComparer.OrdinalIgnoreCase) : new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkbookRulesStore.LoadGlobalProfile failed: {0}", ex.Message));
                return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        public static void SaveRules(string workbookName, Dictionary<string, string> rules)
        {
            try
            {
                string path = GetRulesFilePath(workbookName);
                string dir = Path.GetDirectoryName(path);
                Directory.CreateDirectory(dir);
                string json = JsonConvert.SerializeObject(rules, Formatting.Indented);
                File.WriteAllText(path, json, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkbookRulesStore.SaveRules failed for {0}: {1}", workbookName, ex.Message));
            }
        }

        public static void SaveGlobalProfile(Dictionary<string, string> profile)
        {
            try
            {
                string path = GetGlobalProfilePath();
                string dir = Path.GetDirectoryName(path);
                Directory.CreateDirectory(dir);
                string json = JsonConvert.SerializeObject(profile, Formatting.Indented);
                File.WriteAllText(path, json, new UTF8Encoding(false));
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkbookRulesStore.SaveGlobalProfile failed: {0}", ex.Message));
            }
        }

        public static string FormatRulesForPrompt(Dictionary<string, string> rules)
        {
            if (rules == null || rules.Count == 0) return string.Empty;
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Workbook-specific rules / personalization (local, must obey):");
            foreach (KeyValuePair<string, string> kv in rules)
            {
                if (string.IsNullOrWhiteSpace(kv.Key) || string.IsNullOrWhiteSpace(kv.Value)) continue;
                sb.AppendLine(string.Format("- {0}: {1}", kv.Key.Trim(), kv.Value.Trim()));
            }
            return sb.ToString().TrimEnd();
        }

        public static string FormatRulesForPrompt(string workbookName)
        {
            Dictionary<string, string> wb = LoadRules(workbookName);
            Dictionary<string, string> global = LoadGlobalProfile();
            // merge: workbook overrides global
            Dictionary<string, string> merged = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in global) merged[kv.Key] = kv.Value;
            foreach (var kv in wb) merged[kv.Key] = kv.Value;
            return FormatRulesForPrompt(merged);
        }

        // Helper to read .Rules sheet via ExcelController (COM, UI thread)
        public static string TryReadRulesSheet(ExcelController excelCtrl)
        {
            if (excelCtrl == null) return string.Empty;
            try
            {
                // Use snapshot-like read: look for sheet named ".Rules"
                dynamic app = excelCtrl.GetRawAppObj();
                if (app == null) return string.Empty;
                dynamic ws = null;
                try { ws = app.Worksheets[".Rules"]; } catch { return string.Empty; }
                if (ws == null) return string.Empty;
                dynamic used = null;
                try { used = ws.UsedRange; } catch { return string.Empty; }
                if (used == null) return string.Empty;
                object val = used.Value2;
                if (val is object[,])
                {
                    var arr = (object[,])val;
                    int rr = arr.GetLength(0);
                    int cc = arr.GetLength(1);
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("Rules sheet (.Rules):");
                    for (int r = 1; r <= Math.Min(rr, 50); r++)
                    {
                        List<string> rowVals = new List<string>();
                        for (int c = 1; c <= Math.Min(cc, 5); c++)
                        {
                            string s = Convert.ToString(arr[r, c]) ?? string.Empty;
                            s = s.Trim();
                            if (!string.IsNullOrWhiteSpace(s)) rowVals.Add(s);
                        }
                        if (rowVals.Count >= 2)
                            sb.AppendLine(string.Format("- {0}: {1}", rowVals[0], string.Join(" ", rowVals.GetRange(1, rowVals.Count - 1).ToArray())));
                        else if (rowVals.Count == 1)
                            sb.AppendLine(string.Format("- {0}", rowVals[0]));
                    }
                    string result = sb.ToString().TrimEnd();
                    if (result == "Rules sheet (.Rules):") return string.Empty;
                    return result;
                }
                else if (val != null)
                {
                    string s = Convert.ToString(val).Trim();
                    if (!string.IsNullOrWhiteSpace(s)) return "Rules sheet (.Rules): " + s;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn(string.Format("WorkbookRulesStore.TryReadRulesSheet failed: {0}", ex.Message));
            }
            return string.Empty;
        }
    }
}
