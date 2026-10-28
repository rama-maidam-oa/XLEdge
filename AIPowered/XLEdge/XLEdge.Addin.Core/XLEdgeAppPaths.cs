using System;
using System.IO;
using XLEdge.Addin.Core.Infrastructure;
using XLEdge.Contracts;

namespace XLEdge
{
    /// <summary>
    /// Today's XLEdge data folders (same members), for code ported into the core: the shell's paths when the core is
    /// attached (they are the same folders), today's own computation otherwise. Every folder is created on first use, as today.
    /// </summary>
    public static class XLEdgeAppPaths
    {
        private static string DefaultBase =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"ORBIT\Excel_Logs");

        private static IPathProvider Shell
        {
            get
            {
                try { return CoreServices.Context?.Paths; }
                catch (Exception) { return null; }
            }
        }

        /// <summary>%LOCALAPPDATA%\ORBIT\Excel_Logs (the folder of ORBIT_URLS.xml).</summary>
        public static string BaseFolder => Folder(Shell is IPathProvider p ? Path.GetDirectoryName(p.UrlsFile) : DefaultBase);

        /// <summary>...\Excel_Logs\XLEdge_Logs</summary>
        public static string XLEdgeLogsFolder => Folder(Shell?.Root ?? Path.Combine(DefaultBase, "XLEdge_Logs"));

        /// <summary>...\XLEdge_Logs\Logs</summary>
        public static string LogFolder => Folder(Shell?.Logs ?? Path.Combine(XLEdgeLogsFolder, "Logs"));

        /// <summary>...\XLEdge_Logs\Temp</summary>
        public static string TempFolder => Folder(Shell?.Temp ?? Path.Combine(XLEdgeLogsFolder, "Temp"));

        /// <summary>...\Excel_Logs\ORBIT_URLS.xml (its folder is created, not the file).</summary>
        public static string TempUrlsPath
        {
            get
            {
                string path = Shell?.UrlsFile ?? Path.Combine(DefaultBase, "ORBIT_URLS.xml");
                Folder(Path.GetDirectoryName(path));
                return path;
            }
        }

        /// <summary>...\XLEdge_Logs\BrowserLogs (the WebView2 user data folder).</summary>
        public static string BrowserLogsFolder => Folder(Shell?.BrowserDataPath ?? Path.Combine(XLEdgeLogsFolder, "BrowserLogs"));

        /// <summary>...\BrowserLogs\Login</summary>
        public static string LoginBrowserLogsPath => Folder(Path.Combine(BrowserLogsFolder, "Login"));

        private static readonly object CacheGate = new object();
        private static readonly System.Collections.Generic.Dictionary<string, bool> Ensured =
            new System.Collections.Generic.Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Today's EnsureDirectoryExists (once per folder, as today's Lazy values did); the path is always returned.</summary>
        private static string Folder(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return path;
            lock (CacheGate)
            {
                if (Ensured.ContainsKey(path)) return path;
                if (!Directory.Exists(path))
                {
                    try { Directory.CreateDirectory(path); }
                    catch (Exception ex)
                    {
                        // Worth investigating if this recurs - a required app folder couldn't be created
                        // (e.g. permissions issue), which could break logging/temp-file storage downstream.
                        System.Diagnostics.Debug.WriteLine($"EnsureDirectoryExists: failed to create directory '{path}' - {ex.Message}");
                        XLEdge.Utilities.LogUtility.LogWarn($"EnsureDirectoryExists: failed to create directory '{path}' - {ex.Message}");
                        return path;
                    }
                }
                Ensured[path] = true;
            }
            return path;
        }
    }
}
