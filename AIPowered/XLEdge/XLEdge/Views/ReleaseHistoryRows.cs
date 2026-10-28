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
        public string ReleaseDate { get; set; }
        public string Source { get; set; }
        public string Notes { get; set; }
        public string FolderName { get; set; }
        public bool IsCurrentlyLoaded { get; set; }
        public int ContractVersion { get; set; }
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
                    Notes = e.Notes,
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
