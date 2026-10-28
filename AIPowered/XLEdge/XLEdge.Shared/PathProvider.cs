using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using XLEdge.Contracts;

namespace XLEdge.Shared
{
    /// <summary>
    /// XLEdge folders. Per-user data stays where XLEdge has always kept it
    /// (%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\..., ORBIT_URLS.xml next to it); the hot-reload
    /// state lives next to XLEdge.dll under AddinCore\ (the per-user install folder, writable without admin).
    /// Created once by the shell and remoted to the core.
    /// </summary>
    public class PathProvider : MarshalByRefObject, IPathProvider
    {
        private static readonly JsonSerializerOptions ManifestJsonOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };

        /// <param name="installRoot">Folder XLEdge.dll runs from (Assembly.Location of the shell, not AppDomain.BaseDirectory, which can be Excel's folder).</param>
        public PathProvider(string installRoot)
            : this(installRoot, Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ORBIT", "Excel_Logs"))
        {
        }

        /// <param name="installRoot">Folder XLEdge.dll runs from.</param>
        /// <param name="dataBasePath">Base folder for per-user data (normally %LOCALAPPDATA%\ORBIT\Excel_Logs).</param>
        public PathProvider(string installRoot, string dataBasePath)
        {
            if (string.IsNullOrWhiteSpace(installRoot)) throw new ArgumentException("Install root is required.", nameof(installRoot));
            if (string.IsNullOrWhiteSpace(dataBasePath)) throw new ArgumentException("Data base path is required.", nameof(dataBasePath));

            Root = Path.Combine(dataBasePath, "XLEdge_Logs");
            Logs = Path.Combine(Root, "Logs");
            Temp = Path.Combine(Root, "Temp");
            BrowserDataPath = Path.Combine(Root, "BrowserLogs");
            UrlsFile = Path.Combine(dataBasePath, "ORBIT_URLS.xml");

            string addinCore = Path.Combine(installRoot, "AddinCore");
            VersionsPath = Path.Combine(addinCore, "Versions");
            SharedDependenciesPath = Path.Combine(addinCore, "SharedDependencies");
            ManifestDirectory = Path.Combine(addinCore, "Manifest");
            ManifestFile = Path.Combine(ManifestDirectory, "manifest.json");
            ReleaseHistoryFile = Path.Combine(addinCore, "ReleaseHistory.json");
        }

        public string Root { get; }
        public string Logs { get; }
        public string Temp { get; }
        public string BrowserDataPath { get; }
        public string UrlsFile { get; }
        public string VersionsPath { get; }
        public string SharedDependenciesPath { get; }
        public string ManifestFile { get; }
        public string ManifestDirectory { get; }
        public string ReleaseHistoryFile { get; }
        // Version and date are published together as one object, so a reader (possibly through the remoting
        // proxy, on another thread) never sees one without the other or a half-refreshed pair.
        private volatile LatestEntry _latest;

        private sealed class LatestEntry
        {
            public LatestEntry(string version, string releaseDate) { Version = version; ReleaseDate = releaseDate; }
            public string Version { get; }
            public string ReleaseDate { get; }
        }

        public string LatestVersion => _latest?.Version;
        public string LatestReleaseDate => _latest?.ReleaseDate;

        public void Ensure()
        {
            foreach (var folder in new[] { Logs, Temp, BrowserDataPath, VersionsPath, SharedDependenciesPath, ManifestDirectory })
                Directory.CreateDirectory(folder);
        }

        /// <summary>
        /// Reads the staged manifest.json. A missing, empty or unreadable manifest leaves both values null;
        /// the file is never created here (a default manifest once pointed the loader at a release that
        /// did not exist).
        /// </summary>
        public void Refresh() => _latest = ReadLatest();

        private LatestEntry ReadLatest()
        {
            try
            {
                if (!File.Exists(ManifestFile)) return null;
                string json = File.ReadAllText(ManifestFile);
                if (string.IsNullOrWhiteSpace(json)) return null;

                var entries = JsonSerializer.Deserialize<List<VersionInfo>>(json, ManifestJsonOptions);
                var latest = entries?
                    .Select(e => new { Entry = e, Parsed = TryParseVersion(e?.Version) })
                    .Where(x => x.Parsed != null)
                    .OrderByDescending(x => x.Parsed)
                    .Select(x => x.Entry)
                    .FirstOrDefault();
                return latest == null ? null : new LatestEntry(latest.Version, latest.ReleaseDate);
            }
            catch (JsonException) { return null; }
            catch (IOException) { return null; }
            catch (UnauthorizedAccessException) { return null; }
        }

        public override object InitializeLifetimeService() => null;

        private static Version TryParseVersion(string value) => Version.TryParse(value, out var parsed) ? parsed : null;
    }
}
