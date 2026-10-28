using System;

namespace XLEdge.Shared
{
    /// <summary>One extracted core release in ReleaseHistory.json (the permanent catalog behind Release History).</summary>
    [Serializable]
    public class ReleaseEntry
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }

        /// <summary>Folder under AddinCore\Versions\ the release was extracted to.</summary>
        public string FolderName { get; set; }

        public string Checksum { get; set; }
        public string Notes { get; set; }

        /// <summary>Where the release came from: "Install" (MSI seed or a rebuild picked up at startup), "Online" or "Offline".</summary>
        public string Source { get; set; } = "Install";

        /// <summary>XLEdge.Contracts.ContractInfo.Version the release was built against; 0 when unknown.</summary>
        public int ContractVersion { get; set; }
    }
}
