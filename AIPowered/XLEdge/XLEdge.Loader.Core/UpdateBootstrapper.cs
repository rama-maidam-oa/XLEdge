using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Decides which core release to load and turns release zips into release folders:
    /// AddinCore\Versions\{folder}\ (the release's DLLs + a manifest.json snapshot), catalogued in ReleaseHistory.json.
    /// Third-party DLLs that are identical across releases live once in AddinCore\SharedDependencies\.
    /// </summary>
    public class UpdateBootstrapper
    {
        private const string PartialMarker = ".partial-";

        // Third-party files shared between releases (matched by file name anywhere in the release, relative path kept).
        // XLEdge's own assemblies are never shared: each release must carry the ones it was built with.
        private static readonly HashSet<string> SharedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "MahApps.Metro.IconPacks.Core.dll", "MahApps.Metro.IconPacks.FontAwesome.dll",
            "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.Wpf.dll", "Microsoft.Web.WebView2.WinForms.dll", "WebView2Loader.dll",
            "Wpf.Ui.dll", "Wpf.Ui.Abstractions.dll", "Microsoft.Xaml.Behaviors.dll",
            "NLog.dll",
            "System.Text.Json.dll", "System.Text.Encodings.Web.dll", "System.Buffers.dll", "System.Memory.dll", "System.IO.Pipelines.dll",
            "System.Numerics.Vectors.dll", "System.Runtime.CompilerServices.Unsafe.dll", "System.Threading.Tasks.Extensions.dll",
            "Microsoft.Bcl.AsyncInterfaces.dll", "System.ValueTuple.dll",
        };

        private static readonly string[] ReleaseDateFormats = { "yyyy-MM-ddTHH:mm:ss", "dd-MMM-yyyy", "yyyy-MM-dd" };

        /// <summary>
        /// Chooses the release to load at startup or after a reload request. Never throws; null means nothing can be loaded.
        /// <list type="number">
        /// <item>No catalog yet and a seed staged (installer: manifest.json + zip) — first run: clear Versions, extract the seed,
        /// and remove the Manifest folder only if that worked.</item>
        /// <item>Catalog unreadable, or empty while releases exist — rebuilt from the releases' manifest.json snapshots.</item>
        /// <item>A manifest + zip staged (offline/online update, or a rebuilt dev core) — extract it; the zip is deleted once
        /// catalogued. If it is refused, the newest catalogued release is used instead.</item>
        /// <item>Otherwise the newest catalogued release whose folder holds DLLs and whose contract this shell can load.</item>
        /// </list>
        /// The staged manifest.json never decides which installed release is used.
        /// </summary>
        public ResolvedRelease ResolveVersionToLoad(IXLEdgeContext context, string source = "Install")
        {
            ILogger logger = null;
            try
            {
                logger = context.Logger;
                IPathProvider paths = context.Paths;
                paths.Ensure();
                paths.Refresh();
                RemoveStalePartialFolders(paths.VersionsPath, logger);

                bool catalogExists = File.Exists(paths.ReleaseHistoryFile);
                string stagedZip = FindStagedZip(paths, logger);

                if (!catalogExists && stagedZip != null)
                {
                    logger?.LogInfo("UpdateBootstrapper: first run - installing the staged core release.");
                    ClearVersions(paths.VersionsPath, logger);
                    ResolvedRelease seeded = ExtractStaged(context, stagedZip, "Install");
                    if (seeded != null) TryDeleteDirectory(paths.ManifestDirectory, logger);
                    else logger?.LogWarn("UpdateBootstrapper: the staged release could not be installed; it is kept for the next start.");
                    return seeded;
                }

                List<ReleaseEntry> catalog = catalogExists ? ReadOrRecoverCatalog(paths, logger) : new List<ReleaseEntry>();

                if (stagedZip != null)
                {
                    VersionInfo staged = new VersionParser(logger).ParseVersionFile(paths.ManifestFile).AllVersions.FirstOrDefault();
                    if (staged != null && catalog.Any(e => string.Equals(e.Version, staged.Version, StringComparison.OrdinalIgnoreCase) &&
                                                           string.Equals(e.ReleaseDate, staged.ReleaseDate, StringComparison.OrdinalIgnoreCase)))
                    {
                        logger?.LogInfo($"UpdateBootstrapper: release {staged.Version} is already catalogued (reinstall); reconciling the catalog.");
                        catalog = ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);
                    }
                    ResolvedRelease extracted = ExtractStaged(context, stagedZip, source);
                    if (extracted != null) return extracted;
                    logger?.LogWarn("UpdateBootstrapper: the staged release was not installed; using the newest catalogued release.");
                    catalog = ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
                }

                ReleaseEntry newest = catalog
                    .Where(e => ReleaseHistoryStore.IsSafeFolderName(e.FolderName) && HasDlls(Path.Combine(paths.VersionsPath, e.FolderName)))
                    .Where(e =>
                    {
                        string refusal = ContractCompatibility.RefusalReason(e.ContractVersion);
                        if (refusal != null) logger?.LogWarn($"UpdateBootstrapper: skipping catalogued release '{e.FolderName}': {refusal}.");
                        return refusal == null;
                    })
                    .OrderByDescending(e => ParseReleaseDate(e.ReleaseDate))
                    .FirstOrDefault();
                if (newest != null)
                {
                    logger?.LogInfo($"UpdateBootstrapper: using catalogued release '{newest.FolderName}'.");
                    return new ResolvedRelease { Version = newest.Version, ReleaseDate = newest.ReleaseDate, FolderName = newest.FolderName, ContractVersion = newest.ContractVersion };
                }

                logger?.LogError($"UpdateBootstrapper: no staged release in '{paths.ManifestDirectory}' and no usable catalogued release - nothing to load.");
                return null;
            }
            catch (Exception ex)
            {
                try { logger?.LogException(ex.ToString(), "UpdateBootstrapper.ResolveVersionToLoad"); } catch (Exception) { }
                return null;
            }
        }

        /// <summary>
        /// Extracts a release zip described by <paramref name="manifestJson"/> (a one-entry manifest array) into its release
        /// folder and catalogues it. Returns null — logged, nothing changed — when the release is refused: no usable entry,
        /// unsafe folder name, the active release's folder, missing or wrong checksum, or a contract this shell cannot load.
        /// An unreadable zip throws and leaves no folder behind.
        /// </summary>
        public ResolvedRelease ExtractAndCatalog(IXLEdgeContext context, string manifestJson, byte[] zipBytes, string source)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            ILogger logger = context.Logger;
            IPathProvider paths = context.Paths;

            var parsed = new VersionParser(logger).ParseVersionJson(manifestJson);
            VersionInfo info = parsed.Success ? parsed.AllVersions.FirstOrDefault() : null;
            if (info == null || string.IsNullOrWhiteSpace(info.Version))
                return Refuse(logger, "the manifest has no usable release entry" + (parsed.ErrorMessage == null ? "" : $" ({parsed.ErrorMessage})"));

            string folderName = string.IsNullOrWhiteSpace(info.FolderName) ? ReleaseHistoryStore.BuildFolderName(info.Version, info.ReleaseDate) : info.FolderName;
            if (!ReleaseHistoryStore.IsSafeFolderName(folderName))
                return Refuse(logger, $"invalid release folder name '{folderName}'");
            if (string.Equals(folderName, context.ActiveFolderName, StringComparison.OrdinalIgnoreCase))
                return Refuse(logger, $"'{folderName}' is the release currently loaded");

            string contractRefusal = ContractCompatibility.RefusalReason(info.ContractVersion);
            if (contractRefusal != null)
                return Refuse(logger, $"release {info.Version} {contractRefusal}");

            if (string.IsNullOrWhiteSpace(info.Checksum))
                return Refuse(logger, $"release {info.Version} has no checksum");
            if (zipBytes == null)
                return Refuse(logger, $"release {info.Version} has no zip");
            string actual = Sha256(zipBytes);
            if (!string.Equals(actual, info.Checksum, StringComparison.OrdinalIgnoreCase))
                return Refuse(logger, $"checksum mismatch for release {info.Version} (manifest {info.Checksum}, zip {actual})");

            string versionFolder = Path.Combine(paths.VersionsPath, folderName);
            if (HasDlls(versionFolder))
            {
                // Never re-extract into a populated folder: native DLLs loaded from it stay locked for the life of Excel.
                logger?.LogDebug($"UpdateBootstrapper: '{folderName}' is already extracted; not extracting again.");
            }
            else
            {
                ExtractAtomically(paths.VersionsPath, versionFolder, folderName, manifestJson, zipBytes, logger);
            }
            MaterializeSharedDependencies(versionFolder, paths.SharedDependenciesPath, logger);

            ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, new ReleaseEntry
            {
                Version = info.Version,
                ReleaseDate = info.ReleaseDate,
                FolderName = folderName,
                Checksum = info.Checksum,
                Notes = string.IsNullOrWhiteSpace(info.Notes) ? "XLEdge core release" : info.Notes,
                Source = source,
                ContractVersion = info.ContractVersion,
            });
            logger?.LogInfo($"UpdateBootstrapper: extracted and catalogued '{folderName}' (source={source}).");
            return new ResolvedRelease { Version = info.Version, ReleaseDate = info.ReleaseDate, FolderName = folderName, ContractVersion = info.ContractVersion };
        }

        // ---- ResolveVersionToLoad helpers ---------------------------------------------

        /// <summary>The staged zip: none without a manifest; the manifest's fileName when several are staged, else the newest.</summary>
        private static string FindStagedZip(IPathProvider paths, ILogger logger)
        {
            if (!File.Exists(paths.ManifestFile)) return null;
            var zips = Directory.GetFiles(paths.ManifestDirectory, "*.zip");
            if (zips.Length == 0) return null;
            if (zips.Length == 1) return zips[0];

            string named = new VersionParser(logger).ParseVersionFile(paths.ManifestFile).FileName;
            string match = named == null ? null : zips.FirstOrDefault(z => string.Equals(Path.GetFileName(z), named, StringComparison.OrdinalIgnoreCase));
            if (match != null) return match;

            string newest = zips.OrderByDescending(File.GetLastWriteTimeUtc).First();
            logger?.LogWarn($"UpdateBootstrapper: {zips.Length} zips are staged in '{paths.ManifestDirectory}'; using the newest, '{Path.GetFileName(newest)}'.");
            return newest;
        }

        private ResolvedRelease ExtractStaged(IXLEdgeContext context, string zipPath, string source)
        {
            ILogger logger = context.Logger;
            ResolvedRelease release;
            try
            {
                logger?.LogInfo($"UpdateBootstrapper: extracting '{zipPath}' (source={source}).");
                release = ExtractAndCatalog(context, File.ReadAllText(context.Paths.ManifestFile), File.ReadAllBytes(zipPath), source);
            }
            catch (Exception ex)
            {
                logger?.LogError($"UpdateBootstrapper: could not install '{zipPath}'.", ex.ToString());
                return null;
            }
            if (release != null)
            {
                // Only after the catalog write: a zip left behind is retried, a lost one is not.
                try { File.Delete(zipPath); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { logger?.LogWarn($"UpdateBootstrapper: could not delete '{zipPath}': {ex.Message}"); }
            }
            return release;
        }

        /// <summary>
        /// The catalog, or — when it is unreadable, or empty although releases are installed — a catalog rebuilt from each
        /// release folder's manifest.json snapshot. A broken catalog must never look like a first run (that clears Versions).
        /// </summary>
        private static List<ReleaseEntry> ReadOrRecoverCatalog(IPathProvider paths, ILogger logger)
        {
            if (ReleaseHistoryStore.TryReadAll(paths.ReleaseHistoryFile, out List<ReleaseEntry> entries))
            {
                if (entries.Count > 0) return entries;
            }
            else
            {
                logger?.LogWarn($"UpdateBootstrapper: '{paths.ReleaseHistoryFile}' is unreadable.");
            }

            var recovered = SnapshotEntries(paths.VersionsPath, logger);
            if (entries != null && recovered.Count == 0) return entries; // genuinely empty: nothing installed
            logger?.LogWarn($"UpdateBootstrapper: rebuilding the release catalog from {recovered.Count} installed release(s).");
            ReleaseHistoryStore.RebuildFrom(paths.ReleaseHistoryFile, recovered);
            return ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
        }

        private static List<ReleaseEntry> SnapshotEntries(string versionsPath, ILogger logger)
        {
            var result = new List<ReleaseEntry>();
            if (!Directory.Exists(versionsPath)) return result;
            foreach (string folder in Directory.GetDirectories(versionsPath))
            {
                string name = Path.GetFileName(folder);
                string snapshot = Path.Combine(folder, "manifest.json");
                if (name.Contains(PartialMarker) || !HasDlls(folder) || !File.Exists(snapshot)) continue;
                VersionInfo info = new VersionParser(logger).ParseVersionFile(snapshot).AllVersions.FirstOrDefault();
                if (info == null || string.IsNullOrWhiteSpace(info.Version)) continue;
                result.Add(new ReleaseEntry
                {
                    Version = info.Version,
                    ReleaseDate = info.ReleaseDate,
                    FolderName = name,
                    Checksum = info.Checksum,
                    Notes = info.Notes,
                    Source = "Recovered",
                    ContractVersion = info.ContractVersion,
                });
            }
            return result;
        }

        private static void ClearVersions(string versionsPath, ILogger logger)
        {
            if (!Directory.Exists(versionsPath)) return;
            foreach (string folder in Directory.GetDirectories(versionsPath))
            {
                logger?.LogInfo($"UpdateBootstrapper: first run - removing leftover release folder '{Path.GetFileName(folder)}'.");
                TryDeleteDirectory(folder, logger);
            }
        }

        private static void RemoveStalePartialFolders(string versionsPath, ILogger logger)
        {
            if (!Directory.Exists(versionsPath)) return;
            foreach (string folder in Directory.GetDirectories(versionsPath, "*" + PartialMarker + "*"))
                TryDeleteDirectory(folder, logger);
        }

        private static DateTime ParseReleaseDate(string value) =>
            DateTime.TryParseExact(value, ReleaseDateFormats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var parsed)
                ? parsed
                : DateTime.MinValue;

        // ---- helpers -------------------------------------------------------------------

        private static ResolvedRelease Refuse(ILogger logger, string reason)
        {
            logger?.LogError("UpdateBootstrapper: release refused - " + reason + ".");
            return null;
        }

        internal static string Sha256(byte[] bytes)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }

        internal static bool HasDlls(string folder) =>
            Directory.Exists(folder) && Directory.EnumerateFiles(folder, "*.dll", SearchOption.TopDirectoryOnly).Any();

        /// <summary>Extracts into "{folder}.partial-{guid}" and renames it into place, so a failure never leaves a half-extracted release.</summary>
        private static void ExtractAtomically(string versionsPath, string versionFolder, string folderName, string manifestJson, byte[] zipBytes, ILogger logger)
        {
            Directory.CreateDirectory(versionsPath);
            string partial = Path.Combine(versionsPath, folderName + PartialMarker + Guid.NewGuid().ToString("N"));
            string tempZip = Path.Combine(Path.GetTempPath(), $"XLEdgeRelease_{Guid.NewGuid():N}.zip");
            try
            {
                File.WriteAllBytes(tempZip, zipBytes);
                ZipFile.ExtractToDirectory(tempZip, partial);
                File.WriteAllText(Path.Combine(partial, "manifest.json"), manifestJson);

                // A leftover folder without DLLs (e.g. emptied by hand) is in the way of the rename.
                if (Directory.Exists(versionFolder)) Directory.Delete(versionFolder, recursive: true);
                Directory.Move(partial, versionFolder);
            }
            catch (Exception ex)
            {
                logger?.LogError($"UpdateBootstrapper: extracting '{folderName}' failed.", ex.ToString());
                TryDeleteDirectory(partial, logger);
                throw;
            }
            finally
            {
                try { File.Delete(tempZip); } catch (Exception) { }
            }
        }

        /// <summary>
        /// Moves known third-party files to SharedDependencies (first release that has them), deletes release copies that are
        /// identical to the shared one, and keeps a release copy that differs (that release keeps exactly what it shipped).
        /// </summary>
        private static void MaterializeSharedDependencies(string versionFolder, string sharedFolder, ILogger logger)
        {
            foreach (string file in Directory.GetFiles(versionFolder, "*", SearchOption.AllDirectories))
            {
                if (!SharedFileNames.Contains(Path.GetFileName(file))) continue;
                string relative = file.Substring(versionFolder.Length).TrimStart(Path.DirectorySeparatorChar);
                string target = Path.Combine(sharedFolder, relative);
                try
                {
                    if (!File.Exists(target))
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(target));
                        File.Copy(file, target);
                        File.Delete(file);
                        logger?.LogDebug($"UpdateBootstrapper: cached shared dependency '{relative}'.");
                    }
                    else if (FilesHaveSameContent(file, target))
                    {
                        File.Delete(file);
                    }
                    else
                    {
                        logger?.LogWarn($"UpdateBootstrapper: shared dependency '{relative}' differs from the cached copy; this release keeps its own.");
                    }
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    logger?.LogWarn($"UpdateBootstrapper: could not share '{relative}': {ex.Message}");
                }
            }
            RemoveEmptyDirectories(versionFolder, logger);
        }

        private static bool FilesHaveSameContent(string a, string b)
        {
            if (new FileInfo(a).Length != new FileInfo(b).Length) return false;
            return Sha256(File.ReadAllBytes(a)) == Sha256(File.ReadAllBytes(b));
        }

        private static void RemoveEmptyDirectories(string root, ILogger logger)
        {
            foreach (string dir in Directory.GetDirectories(root, "*", SearchOption.AllDirectories).OrderByDescending(d => d.Length))
            {
                try
                {
                    if (!Directory.EnumerateFileSystemEntries(dir).Any()) Directory.Delete(dir);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    logger?.LogWarn($"UpdateBootstrapper: could not remove empty folder '{dir}': {ex.Message}");
                }
            }
        }

        private static void TryDeleteDirectory(string folder, ILogger logger)
        {
            try
            {
                if (Directory.Exists(folder)) Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                logger?.LogWarn($"UpdateBootstrapper: could not delete '{folder}': {ex.Message}");
            }
        }
    }
}
