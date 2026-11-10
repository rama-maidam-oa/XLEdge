using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Views
{
    /// <summary>Result of checking an offline folder for a core release (a manifest.json + zip pair).</summary>
    internal sealed class OfflineScanResult
    {
        public bool Ok { get; set; }
        public string Message { get; set; }
        public string ManifestPath { get; set; }
        public string ZipPath { get; set; }
        public VersionInfo Release { get; set; }

        /// <summary>The release is valid but not newer than the one loaded.</summary>
        public bool NotNewer { get; set; }
    }

    /// <summary>The logic behind the Offline option of the reload window: find, verify and stage a release from a folder.</summary>
    internal static class OfflineReleaseScanner
    {
        /// <summary>
        /// Checks <paramref name="folder"/> for the newest manifest*.json and the newest v*.zip, and verifies them: parseable,
        /// checksum present and matching, a contract this shell can load, and newer than the loaded release.
        /// </summary>
        public static OfflineScanResult Scan(string folder, string loadedVersion, string loadedReleaseDate)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return Fail("Folder not found: " + folder);

            string manifest = Newest(folder, "manifest*.json");
            string zip = Newest(folder, "v*.zip");
            if (manifest == null || zip == null)
                return Fail("No manifest.json + zip pair found in this folder.");

            var parsed = new VersionParser().ParseVersionFile(manifest);
            VersionInfo release = parsed.Success ? parsed.AllVersions.FirstOrDefault() : null;
            if (release == null)
                return Fail("Could not parse manifest.json: " + (parsed.ErrorMessage ?? "no release entry"));
            if (string.IsNullOrWhiteSpace(release.Checksum))
                return Fail("manifest.json has no checksum recorded - cannot verify, refusing to reload.");

            // The zip the manifest names wins over the newest one (another zip may sit in the same folder).
            if (!string.IsNullOrWhiteSpace(release.FileName) && release.FileName.IndexOfAny(Path.GetInvalidFileNameChars()) < 0)
            {
                string named = Path.Combine(folder, release.FileName);
                if (File.Exists(named)) zip = named;
            }

            long length;
            string actual;
            using (var stream = new FileStream(zip, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var sha = SHA256.Create())
            {
                length = stream.Length;
                actual = BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
            }
            if (!string.Equals(actual, release.Checksum, StringComparison.OrdinalIgnoreCase))
                return Fail($"Checksum mismatch - the zip may be corrupt or incomplete. Expected {release.Checksum}, got {actual}.");

            string refusal = ContractCompatibility.RefusalReason(release.ContractVersion);
            if (refusal != null)
                return Fail($"This version cannot be loaded: it {refusal}.");

            var result = new OfflineScanResult { ManifestPath = manifest, ZipPath = zip, Release = release };
            DateTime? loaded = ReleaseDates.TryParse(loadedReleaseDate);
            DateTime? candidate = ReleaseDates.TryParse(release.ReleaseDate);
            if (loaded != null && (candidate == null || candidate <= loaded))
            {
                result.NotNewer = true;
                result.Message = $"No updates found.\n\nYou already have the latest version: {VersionText.Display(loadedVersion)} (version date {ReleaseDates.Display(loadedReleaseDate)}).\n\nChecked version: {VersionText.Display(release.Version)} ({ReleaseDates.Display(release.ReleaseDate)}).";
                return result;
            }

            result.Ok = true;
            result.Message = $"Ready to reload: version {VersionText.Display(release.Version)}, version date {ReleaseDates.Display(release.ReleaseDate)} ({Math.Round(length / 1024.0):0} KB).";
            return result;
        }

        /// <summary>
        /// Stages a verified release for the bootstrapper: the zip and manifest.json go into AddinCore\Manifest\ (unless already
        /// there), and any other staged zip is removed. Throws on a refused scan or an I/O failure.
        /// </summary>
        public static void Stage(OfflineScanResult result, IPathProvider paths)
        {
            if (result == null || !result.Ok) throw new InvalidOperationException("Only a verified release can be staged.");
            Directory.CreateDirectory(paths.ManifestDirectory);

            string stagedZip = Path.Combine(paths.ManifestDirectory, Path.GetFileName(result.ZipPath));
            foreach (string other in Directory.GetFiles(paths.ManifestDirectory, "*.zip"))
                if (!SamePath(other, result.ZipPath) && !SamePath(other, stagedZip)) File.Delete(other);

            if (!SamePath(result.ZipPath, stagedZip)) File.Copy(result.ZipPath, stagedZip, overwrite: true);
            if (!SamePath(result.ManifestPath, paths.ManifestFile)) File.Copy(result.ManifestPath, paths.ManifestFile, overwrite: true);
        }

        private static OfflineScanResult Fail(string message) => new OfflineScanResult { Ok = false, Message = message };

        private static string Newest(string folder, string pattern) =>
            Directory.GetFiles(folder, pattern).OrderByDescending(File.GetLastWriteTimeUtc).FirstOrDefault();

        private static bool SamePath(string a, string b) =>
            string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Release dates as written by the core build ("yyyy-MM-ddTHH:mm:ss"), plus two older forms.</summary>
    internal static class ReleaseDates
    {
        private static readonly string[] Formats = { "yyyy-MM-ddTHH:mm:ss", "dd-MMM-yyyy", "yyyy-MM-dd" };

        public static DateTime? TryParse(string value) =>
            DateTime.TryParseExact(value?.Trim(), Formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed) ? parsed : (DateTime?)null;

        /// <summary>OISR-22872: a version date as the windows show it (like the About window's build date); display only.</summary>
        public static string Display(string value)
        {
            DateTime? parsed = TryParse(value);
            if (parsed == null) return value?.Trim() ?? string.Empty;
            return parsed.Value.ToString(value.Trim().Length > 11 ? "dd-MMM-yyyy hh:mm tt" : "dd-MMM-yyyy", CultureInfo.InvariantCulture);
        }
    }
}
