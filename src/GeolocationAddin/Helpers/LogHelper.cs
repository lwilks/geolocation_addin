using System;
using System.IO;

namespace GeolocationAddin.Helpers
{
    public static class LogHelper
    {
        private static readonly string LogDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData),
            "GeolocationAddin");

        public static readonly string LogFilePath = Path.Combine(LogDir, "geolocation.log");
        private static readonly string LogFile = LogFilePath;

        public static void Info(string message)
        {
            WriteLog("INFO", message);
        }

        public static void Error(string message)
        {
            WriteLog("ERROR", message);
        }

        private static void WriteLog(string level, string message)
        {
            try
            {
                Directory.CreateDirectory(LogDir);
                var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}";
                File.AppendAllText(LogFile, line + Environment.NewLine);
            }
            catch
            {
                // Logging should never crash the addin
            }
        }
    }
}
