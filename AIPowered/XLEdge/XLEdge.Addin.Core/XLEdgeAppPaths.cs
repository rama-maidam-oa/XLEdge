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

        private static string Folder(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return path;
            try
            {
                if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Today's behaviour: the path is still returned; the first write to it reports the problem.
            }
            return path;
        }
    }
}
