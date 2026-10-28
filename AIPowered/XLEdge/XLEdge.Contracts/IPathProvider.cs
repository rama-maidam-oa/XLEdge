using System.Collections.Generic;

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

        string LoginBrowserPath { get; }

        /// <summary>ORBIT_URLS.xml, the server list (shared with GLSense).</summary>
        string UrlsFile { get; }

        // Hot-reload layout (next to XLEdge.dll)
        string VersionsPath { get; }
        string SharedDependenciesPath { get; }
        string ManifestFile { get; }
        string ManifestDirectory { get; }
        string ReleaseHistoryFile { get; }

        // Latest entry of manifest.json
        string LatestVersion { get; }
        string LatestReleaseDate { get; }
        string LatestDownloadUrl { get; }
        string LatestChecksum { get; }
        string LatestNotes { get; }
        bool LatestMandatory { get; }
        IReadOnlyList<VersionInfo> AllVersions { get; }

        /// <summary>Creates any missing folder.</summary>
        void Ensure();

        /// <summary>Re-reads manifest.json and refreshes the Latest* properties.</summary>
        void Refresh();

        /// <summary>Overwrites manifest.json with this single entry, then calls Refresh.</summary>
        void WriteManifest(VersionInfo info);
    }
}
