namespace XLEdge.Contracts
{
    /// <summary>
    /// Folders and files used by XLEdge, plus the hot-reload layout under the install folder
    /// (AddinCore\Manifest, AddinCore\Versions, AddinCore\SharedDependencies, ReleaseHistory.json).
    /// Owned by the shell and remoted to the core, so its implementation must be a MarshalByRefObject.
    /// </summary>
    public interface IPathProvider
    {
        // Per-user data (%LOCALAPPDATA%\ORBIT\Excel_Logs\XLEdge_Logs\...)
        string Root { get; }
        string Logs { get; }
        string Temp { get; }

        /// <summary>WebView2 user data folder.</summary>
        string BrowserDataPath { get; }

        /// <summary>ORBIT_URLS.xml, the server list (shared with GLSense).</summary>
        string UrlsFile { get; }

        // Hot-reload layout (next to XLEdge.dll)
        string VersionsPath { get; }
        string SharedDependenciesPath { get; }
        string ManifestFile { get; }
        string ManifestDirectory { get; }
        string ReleaseHistoryFile { get; }

        /// <summary>
        /// Version of the highest entry in the staged manifest.json, or null when there is none.
        /// Describes what is staged for installation, not what is installed: the installed releases
        /// are the release history entries whose folders exist.
        /// </summary>
        string LatestVersion { get; }

        string LatestReleaseDate { get; }

        /// <summary>Creates any missing folder.</summary>
        void Ensure();

        /// <summary>Re-reads manifest.json and refreshes LatestVersion / LatestReleaseDate.</summary>
        void Refresh();
    }
}
