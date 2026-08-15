using System;
using System.IO;
using System.Threading;

namespace MistralOfficeAddin.Core
{
    public static class Logger
    {
        private static readonly object _lock = new object();
        private static string _logFilePath;

        static Logger()
        {
            try
            {
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string dir = Path.Combine(localAppData, "MistralOfficeAddin");
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }
                _logFilePath = Path.Combine(dir, "addin.log");
            }
            catch
            {
                _logFilePath = Path.Combine(Path.GetTempPath(), "MistralAddinLog.txt");
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
                lock (_lock)
                {
                    string line = string.Format("{0:yyyy-MM-dd HH:mm:ss.fff} [{1:D2}] [{2}] {3}", DateTime.Now, Thread.CurrentThread.ManagedThreadId, level, message);
                    File.AppendAllText(_logFilePath, line + Environment.NewLine);
                }
            }
            catch
            {
                // Silently swallow logging failures to prevent host disruption
            }
        }
    }
}
