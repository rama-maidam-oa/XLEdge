using System;

namespace XLEdge.Contracts
{
    /// <summary>
    /// One XLEdge core release, as described by a manifest.json entry.
    /// [Serializable] because it crosses the shell/core AppDomain boundary by value
    /// (IPathProvider.AllVersions, IPathProvider.WriteManifest, IXLEdgeContext.AllVersions).
    /// </summary>
    [Serializable]
    public class VersionInfo
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }

        // Manifest fields; System.Text.Json binds them case-insensitively to
        // "downloadUrl", "checksum", "notes", "mandatory", "folderName", "fileName".
        public string DownloadUrl { get; set; }
        public string Checksum { get; set; }
        public string Notes { get; set; }
        public bool Mandatory { get; set; }

        /// <summary>Release folder under AddinCore\Versions\, "v{version}_{releaseDate}".</summary>
        public string FolderName { get; set; }

        /// <summary>Release zip name, "{FolderName}.zip".</summary>
        public string FileName { get; set; }
    }
}
