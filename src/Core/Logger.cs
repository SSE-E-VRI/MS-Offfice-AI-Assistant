using System;
using System.IO;
using System.Threading;

namespace MSOfficeAIAssistant.Core
{
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logFilePath;

        static Logger()
        {
            try
            {
                _logFilePath = AppPaths.InDataDirectory("addin.log");
            }
            catch
            {
                _logFilePath = Path.Combine(Path.GetTempPath(), "MSOfficeAIAssistant.log");
            }
        }

        public static void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public static void Warn(string message)
        {
            WriteLog("WARN", message);
        }

        public static void Error(string message, Exception ex = null)
        {
            string msg = message;
            if (ex != null)
            {
                msg = string.Format("{0} | Exception: {1}: {2}\nStack: {3}", message, ex.GetType().Name, ex.Message, ex.StackTrace);
                if (ex.InnerException != null)
                {
                    msg = string.Format("{0}\nInner Exception: {1}: {2}", msg, ex.InnerException.GetType().Name, ex.InnerException.Message);
                }
            }
            WriteLog("ERROR", msg);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                string sanitized = SanitizeSecrets(message);
                lock (_lock)
                {
                    string line = string.Format("{0:yyyy-MM-dd HH:mm:ss.fff} [{1:D2}] [{2}] {3}", DateTime.Now, Thread.CurrentThread.ManagedThreadId, level, sanitized);
                    File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Silently swallow logging failures to prevent host disruption
            }
        }

        public static string SanitizeSecrets(string message)
        {
            if (string.IsNullOrEmpty(message)) return string.Empty;
            string result = message;
            try
            {
                result = System.Text.RegularExpressions.Regex.Replace(
                    result,
                    @"Bearer\s+[A-Za-z0-9_\-\.]{8,}",
                    "Bearer [REDACTED]",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(
                    result,
                    @"(?i)(api[_-]?key|password|client[_-]?secret|token)\s*[:=]\s*[""']?([^""'\s,;]+)[""']?",
                    "$1=[REDACTED]");

                result = System.Text.RegularExpressions.Regex.Replace(
                    result,
                    @"""(?i)(apiKey|api_key|password|secret|token)""\s*:\s*""[^""]+""",
                    "\"$1\":\"[REDACTED]\"");

                result = System.Text.RegularExpressions.Regex.Replace(
                    result,
                    @"\b(sk-[A-Za-z0-9_\-]{20,}|gsk_[A-Za-z0-9_\-]{20,}|AIza[A-Za-z0-9_\-]{30,})\b",
                    "[REDACTED_KEY]");
            }
            catch { }
            return result;
        }
    }
}
