namespace XLEdge.Loader.Core
{
    /// <summary>The core release chosen to load: its folder under AddinCore\Versions\ plus what to show and check.</summary>
    public class ResolvedRelease
    {
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string FolderName { get; set; }

        /// <summary>XLEdge.Contracts.ContractInfo.Version the release was built against; 0 when unknown.</summary>
        public int ContractVersion { get; set; }
    }
}
