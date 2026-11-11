using System;
using System.Collections.Generic;
using System.Linq;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Views
{
    /// <summary>One line of the Version History window.</summary>
    internal sealed class ReleaseRow
    {
        public string Version { get; set; }
        /// <summary>The version as shown (11.1.3, not 11.1.3.0).</summary>
        public string DisplayVersion => VersionText.Display(Version);
        public string ReleaseDate { get; set; }
        public string ReleaseDateDisplay => ReleaseDates.Display(ReleaseDate);
        public string Source { get; set; }
        public string SourceDescription => DescribeSource(Source);
        public string FolderName { get; set; }
        public bool IsCurrentlyLoaded { get; set; }
        public int ContractVersion { get; set; }

        /// <summary>OISR-22872: the Source cell tooltip explains where the version came from.</summary>
        internal static string DescribeSource(string source)
        {
            switch (source?.Trim().ToLowerInvariant())
            {
                case "install": return "Install - came with the XLEdge installer and was set up automatically when Excel started.";
                case "online": return "Online - downloaded from the server with Reload Add-in > Online.";
                case "offline": return "Offline - loaded from a folder on this computer with Reload Add-in > Offline.";
                default: return source ?? string.Empty;
            }
        }
    }

    internal static class ReleaseHistoryRows
    {
        /// <summary>
        /// Rows for the catalogued releases, newest first. The loaded one is the release folder the core runs from, or — when no
        /// folder is recorded — the entry with the loaded version and release date.
        /// </summary>
        public static List<ReleaseRow> Build(IEnumerable<ReleaseEntry> entries, IXLEdgeContext context)
        {
            string activeFolder = context?.ActiveFolderName;
            string version = context?.Version?.Trim();
            string date = context?.ReleaseDate?.Trim();

            return entries
                .Select(e => new ReleaseRow
                {
                    Version = e.Version,
                    ReleaseDate = e.ReleaseDate,
                    Source = e.Source,
                    FolderName = e.FolderName,
                    ContractVersion = e.ContractVersion,
                    IsCurrentlyLoaded = !string.IsNullOrWhiteSpace(activeFolder)
                        ? string.Equals(e.FolderName, activeFolder, StringComparison.OrdinalIgnoreCase)
                        : !string.IsNullOrEmpty(version) &&
                          string.Equals(e.Version?.Trim(), version, StringComparison.OrdinalIgnoreCase) &&
                          string.Equals(e.ReleaseDate?.Trim(), date, StringComparison.OrdinalIgnoreCase),
                })
                .OrderByDescending(r => ReleaseDates.TryParse(r.ReleaseDate) ?? DateTime.MinValue)
                .ToList();
        }
    }
}
