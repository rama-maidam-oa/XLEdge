using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Online
{
    /// <summary>
    /// The reload picker's Online list from the server's index and this machine's release history (GLSense's
    /// OnlineReleaseClassifier, XLEdge-named): loaded, already downloaded or new; newest first; the newest new one ticked.
    /// </summary>
    internal static class OnlineReleaseClassifier
    {
        public static List<OnlineReleaseRow> Classify(IEnumerable<VersionInfo> serverEntries, IReadOnlyList<ReleaseEntry> localEntries,
                                                      string currentVersion, string currentReleaseDate)
        {
            var rows = new List<OnlineReleaseRow>();
            if (serverEntries == null) return rows;
            localEntries = localEntries ?? new List<ReleaseEntry>();
            foreach (VersionInfo entry in serverEntries)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Version)) continue;
                bool loaded = Same(entry.Version, currentVersion) && Same(entry.ReleaseDate, currentReleaseDate);
                bool downloaded = !loaded && localEntries.Any(l => l != null && Same(l.Version, entry.Version) && Same(l.ReleaseDate, entry.ReleaseDate));
                rows.Add(new OnlineReleaseRow
                {
                    Version = entry.Version, ReleaseDate = entry.ReleaseDate, Mandatory = entry.Mandatory, FolderName = entry.FolderName,
                    FileName = entry.FileName, Checksum = entry.Checksum, Notes = entry.Notes,
                    IsCurrentlyLoaded = loaded, IsAlreadyDownloaded = downloaded,
                });
            }
            rows = rows.OrderByDescending(r => ReleaseDateOrMin(r.ReleaseDate)).ToList();
            OnlineReleaseRow newestNew = rows.FirstOrDefault(r => r.IsSelectable);
            if (newestNew != null) newestNew.IsChecked = true;
            return rows;
        }

        /// <summary>Newest/oldest is decided by the release date in one place (the list order and which release loads).</summary>
        public static DateTime ReleaseDateOrMin(string releaseDate) =>
            DateTime.TryParse(releaseDate, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date) ? date : DateTime.MinValue;

        private static bool Same(string a, string b) => string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
    }
}
