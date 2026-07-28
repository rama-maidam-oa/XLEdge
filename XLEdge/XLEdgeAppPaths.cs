using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLEdge.Utilities;

namespace XLEdge
{
    public static class XLEdgeAppPaths
    {
        private static readonly Lazy<string> _baseFolder = new(() =>
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"ORBIT\Excel_Logs");

            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _xlEdgeLogsFolder = new(() =>
        {
            string path = Path.Combine(BaseFolder, "XLEdge_Logs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _logFolder = new(() =>
        {
            string path = Path.Combine(XLEdgeLogsFolder, "Logs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _tempFolder = new(() =>
        {
            string path = Path.Combine(XLEdgeLogsFolder, "Temp");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _tempUrlsPath = new(() =>
        {
            string path = Path.Combine(BaseFolder, "ORBIT_URLS.xml");
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                EnsureDirectoryExists(directory);
            }
            return path;
        });

        private static readonly Lazy<string> _browserLogsFolder = new(() =>
        {
            string path = Path.Combine(XLEdgeLogsFolder, "BrowserLogs");
            EnsureDirectoryExists(path);
            return path;
        });

        private static readonly Lazy<string> _loginBrowserLogs = new(() =>
        {
            string path = Path.Combine(BrowserLogsFolder, "Login");
            EnsureDirectoryExists(path);
            return path;
        });

        // Public properties
        public static string BaseFolder => _baseFolder.Value;
        public static string LogFolder => _logFolder.Value;
        public static string XLEdgeLogsFolder => _xlEdgeLogsFolder.Value;
        public static string TempUrlsPath => _tempUrlsPath.Value;
        public static string BrowserLogsFolder => _browserLogsFolder.Value;
        public static string LoginBrowserLogsPath => _loginBrowserLogs.Value;
        public static string TempFolder => _tempFolder.Value;

        private static void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                try { Directory.CreateDirectory(path); }
                catch (Exception ex)
                {
                    // Worth investigating if this recurs - a required app folder couldn't be created
                    // (e.g. permissions issue), which could break logging/temp-file storage downstream.
                    // Uses Debug.WriteLine as a fallback too since LogUtility's own logger may not be
                    // initialized this early in startup (paths are resolved before LogHelper runs).
                    System.Diagnostics.Debug.WriteLine($"{nameof(EnsureDirectoryExists)}: failed to create directory '{path}' - {ex.Message}");
                    LogUtility.LogWarn($"{nameof(EnsureDirectoryExists)}: failed to create directory '{path}' - {ex.Message}");
                }
            }
        }
    }
}
