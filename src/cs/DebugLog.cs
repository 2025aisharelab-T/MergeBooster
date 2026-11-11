using System;
using System.IO;
using System.Text;

namespace ExcelDiffAggregator
{
    // Appends timestamped lines to logs\MergeBooster_debug_yyyyMMdd.log next to the exe
    // (falls back to %LOCALAPPDATA%\MergeBooster\logs when that folder is not writable).
    public static class DebugLog
    {
        private static readonly object Sync = new object();
        private static string logDir;

        public static string CurrentPath
        {
            get { return Path.Combine(GetLogDir(), "MergeBooster_debug_" + DateTime.Now.ToString("yyyyMMdd") + ".log"); }
        }

        public static void Write(string message)
        {
            Write("INFO", message);
        }

        public static void Error(string message, Exception ex)
        {
            Write("ERROR", message + (ex != null ? Environment.NewLine + ex.ToString() : ""));
        }

        public static void Write(string level, string message)
        {
            try
            {
                string line = string.Format("{0} [{1}] [T{2}] {3}{4}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    level,
                    System.Threading.Thread.CurrentThread.ManagedThreadId,
                    message,
                    Environment.NewLine);
                lock (Sync)
                {
                    File.AppendAllText(CurrentPath, line, Encoding.UTF8);
                }
            }
            catch
            {
                // Logging must never break the merge itself
            }
        }

        private static string GetLogDir()
        {
            if (logDir != null) return logDir;
            lock (Sync)
            {
                if (logDir != null) return logDir;
                string primary = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                if (TryEnsureWritable(primary))
                {
                    logDir = primary;
                }
                else
                {
                    string fallback = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MergeBooster", "logs");
                    TryEnsureWritable(fallback);
                    logDir = fallback;
                }
                return logDir;
            }
        }

        private static bool TryEnsureWritable(string dir)
        {
            try
            {
                Directory.CreateDirectory(dir);
                string probe = Path.Combine(dir, ".write_probe");
                File.WriteAllText(probe, "");
                File.Delete(probe);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
