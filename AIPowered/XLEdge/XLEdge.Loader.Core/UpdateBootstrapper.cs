using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Decides which core release to load and turns release zips into release folders:
    /// AddinCore\Versions\{folder}\ (the release's DLLs + a manifest.json snapshot), catalogued in ReleaseHistory.json.
    /// Third-party DLLs that are identical across releases live once in AddinCore\SharedDependencies\.
    /// <para>
    /// Several Excel processes can start at once; everything that changes Versions\, SharedDependencies\, Manifest\ or the
    /// catalog runs under one install-wide named mutex, and every decision is re-checked after it is acquired.
    /// </para>
    /// </summary>
    public class UpdateBootstrapper
    {
        private const string PartialMarker = ".partial-";
        private const string RejectedMarker = ".rejected-";
        private const string SelectedReleaseFileName = "SelectedRelease.txt";
        private static readonly TimeSpan InstallLockTimeout = TimeSpan.FromMinutes(2);
        private static readonly TimeSpan AbandonedPartialAge = TimeSpan.FromMinutes(10);
        private const int MoveAttempts = 10;
        private static readonly TimeSpan MoveRetryDelay = TimeSpan.FromMilliseconds(200);
        private static readonly string[] ReleaseDateFormats = { "yyyy-MM-ddTHH:mm:ss", "dd-MMM-yyyy", "yyyy-MM-dd" };

        // WebView2 finds WebView2Loader.dll next to Microsoft.Web.WebView2.Core.dll, so these are shared all together or not at all.
        private static readonly HashSet<string> WebView2FileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.Wpf.dll", "Microsoft.Web.WebView2.WinForms.dll", "WebView2Loader.dll",
        };

        // Other third-party files shared between releases (matched by file name anywhere in the release, relative path kept).
        // XLEdge's own assemblies are never shared: each release carries the ones it was built with.
        private static readonly HashSet<string> SharedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "MahApps.Metro.IconPacks.Core.dll", "MahApps.Metro.IconPacks.FontAwesome.dll",
            "Wpf.Ui.dll", "Wpf.Ui.Abstractions.dll", "Microsoft.Xaml.Behaviors.dll",
            "NLog.dll",
            "System.Text.Json.dll", "System.Text.Encodings.Web.dll", "System.Buffers.dll", "System.Memory.dll", "System.IO.Pipelines.dll",
            "System.Numerics.Vectors.dll", "System.Runtime.CompilerServices.Unsafe.dll", "System.Threading.Tasks.Extensions.dll",
            "Microsoft.Bcl.AsyncInterfaces.dll", "System.ValueTuple.dll",
        };

        /// <summary>
        /// Chooses the release to load at startup or after a reload request. Never throws; null means nothing can be loaded.
        /// <list type="number">
        /// <item>No catalog and nothing installed — first run: clear leftovers, install the staged seed (manifest.json + zip).</item>
        /// <item>Catalog missing, unreadable, or empty while releases are installed — rebuilt from the releases' manifest.json snapshots.</item>
        /// <item>A manifest + zip staged (offline/online update, rebuilt dev core) — install it; the zip is deleted once catalogued,
        /// or parked as "*.rejected-*" when refused. A refused update falls back to the steps below.</item>
        /// <item>The release last loaded (see <see cref="RememberLoaded"/>), so a Release History rollback survives a restart.</item>
        /// <item>The newest catalogued release whose folder holds DLLs and whose contract this shell can load.</item>
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
                return WithInstallLock(paths, () => ResolveUnderLock(context, paths, logger, source));
            }
            catch (Exception ex)
            {
                try { logger?.LogException(ex.ToString(), "UpdateBootstrapper.ResolveVersionToLoad"); } catch (Exception) { }
                return null;
            }
        }

        private ResolvedRelease ResolveUnderLock(IXLEdgeContext context, IPathProvider paths, ILogger logger, string source)
        {
            const string Log = "UpdateBootstrapper.ResolveVersionToLoad: ";
            paths.Refresh();
            RemoveAbandonedPartialFolders(paths.VersionsPath, logger);
            string stagedZip = FindStagedZip(paths, logger);

            if (!File.Exists(paths.ReleaseHistoryFile) && SnapshotEntries(paths.VersionsPath, logger).Count == 0)
            {
                if (stagedZip == null)
                {
                    logger?.LogError(Log + $"no staged release in '{paths.ManifestDirectory}' and no installed release - nothing to load.");
                    return null;
                }
                logger?.LogDebug(Log + "first run - installing the staged core release.");
                ClearVersions(paths.VersionsPath, logger);
                ResolvedRelease seeded = ExtractStaged(context, stagedZip, "Install");
                if (seeded != null) TryDeleteDirectory(paths.ManifestDirectory, logger);
                return seeded;
            }

            List<ReleaseEntry> catalog = ReadOrRecoverCatalog(paths, logger);

            if (stagedZip != null)
            {
                VersionInfo staged = FirstManifestEntry(paths.ManifestFile, logger);
                if (staged != null && catalog.Any(e => string.Equals(e.Version, staged.Version, StringComparison.OrdinalIgnoreCase) &&
                                                       string.Equals(e.ReleaseDate, staged.ReleaseDate, StringComparison.OrdinalIgnoreCase)))
                {
                    logger?.LogDebug(Log + $"release {VersionText.Display(staged.Version)} is already catalogued (reinstall); reconciling the catalog.");
                    ReleaseHistoryStore.Reconcile(paths.ReleaseHistoryFile, paths.VersionsPath);
                }
                ResolvedRelease extracted = ExtractStaged(context, stagedZip, source);
                if (extracted != null) return extracted;
                logger?.LogWarn(Log + "the staged release was not installed; using an installed release.");
                catalog = ReleaseHistoryStore.ReadAll(paths.ReleaseHistoryFile);
            }

            var loadable = catalog
                .Where(e => IsUsableFolderName(e.FolderName) && HasDlls(Path.Combine(paths.VersionsPath, e.FolderName)))
                .Where(e =>
                {
                    string refusal = ContractCompatibility.RefusalReason(e.ContractVersion);
                    if (refusal != null) logger?.LogWarn(Log + $"skipping catalogued release '{e.FolderName}': {refusal}.");
                    return refusal == null;
                })
                .ToList();

            string remembered = ReadRemembered(paths);
            ReleaseEntry chosen = loadable.FirstOrDefault(e => string.Equals(e.FolderName, remembered, StringComparison.OrdinalIgnoreCase))
                ?? loadable.OrderByDescending(e => ParseReleaseDate(e.ReleaseDate)).FirstOrDefault();
            if (chosen != null)
            {
                logger?.LogDebug(Log + $"using installed release '{chosen.FolderName}'" + (string.Equals(chosen.FolderName, remembered, StringComparison.OrdinalIgnoreCase) ? " (last loaded)." : "."));
                return ToResolved(chosen);
            }

            logger?.LogError(Log + $"no staged release in '{paths.ManifestDirectory}' and no usable installed release - nothing to load.");
            return null;
        }

        /// <summary>
        /// Records the release the shell has just loaded, so the next start loads it again (until a new release is staged).
        /// Call after every successful load, including a Release History rollback. Never throws.
        /// </summary>
        public void RememberLoaded(IXLEdgeContext context, string folderName)
        {
            try
            {
                string file = SelectedReleaseFile(context.Paths);
                string tmp = file + ".tmp";
                Directory.CreateDirectory(Path.GetDirectoryName(file));
                File.WriteAllText(tmp, folderName ?? string.Empty);
                if (File.Exists(file)) File.Replace(tmp, file, null);
                else File.Move(tmp, file);
            }
            catch (Exception ex)
            {
                try { context?.Logger?.LogWarn($"UpdateBootstrapper.RememberLoaded: could not record '{folderName}': {ex.Message}"); } catch (Exception) { }
            }
        }

        /// <summary>
        /// Extracts a release zip described by <paramref name="manifestJson"/> (its first entry describes the release) into its
        /// release folder and catalogues it. Returns null — logged, nothing changed — when the release is refused: no usable
        /// entry, unsafe folder name, missing or wrong checksum, a contract this shell cannot load, or a folder that already
        /// holds a different build. An unreadable zip throws and leaves no folder behind.
        /// </summary>
        public ResolvedRelease ExtractAndCatalog(IXLEdgeContext context, string manifestJson, byte[] zipBytes, string source)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            return WithInstallLock(context.Paths, () => ExtractUnderLock(context, manifestJson, zipBytes, source));
        }

        private ResolvedRelease ExtractUnderLock(IXLEdgeContext context, string manifestJson, byte[] zipBytes, string source)
        {
            const string Log = "UpdateBootstrapper.ExtractAndCatalog: ";
            ILogger logger = context.Logger;
            IPathProvider paths = context.Paths;

            var parsed = new VersionParser(logger).ParseVersionJson(manifestJson);
            VersionInfo info = parsed.Success ? parsed.AllVersions.FirstOrDefault() : null;
            if (info == null || string.IsNullOrWhiteSpace(info.Version))
                return Refuse(logger, "the manifest has no usable release entry" + (parsed.ErrorMessage == null ? "" : $" ({parsed.ErrorMessage})"));

            string folderName = string.IsNullOrWhiteSpace(info.FolderName) ? ReleaseHistoryStore.BuildFolderName(info.Version, info.ReleaseDate) : info.FolderName;
            if (!IsUsableFolderName(folderName))
                return Refuse(logger, $"invalid release folder name '{folderName}'");

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
                // Never re-extract into a populated folder (its DLLs may be loaded, and native ones stay locked for the life of
                // Excel) — but only accept it when it holds this very build.
                string snapshot = Path.Combine(versionFolder, "manifest.json");
                VersionInfo existing = File.Exists(snapshot) ? FirstManifestEntry(snapshot, logger) : null;
                if (existing != null && !string.IsNullOrWhiteSpace(existing.Checksum) &&
                    !string.Equals(existing.Checksum, info.Checksum, StringComparison.OrdinalIgnoreCase))
                    return Refuse(logger, $"'{folderName}' already holds a different build (checksum {existing.Checksum})");
                if (existing == null) File.WriteAllText(snapshot, manifestJson);
                logger?.LogDebug(Log + $"'{folderName}' is already extracted; not extracting again.");
            }
            else if (string.Equals(folderName, context.ActiveFolderName, StringComparison.OrdinalIgnoreCase))
            {
                return Refuse(logger, $"'{folderName}' is the release currently loaded");
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
            logger?.LogDebug(Log + $"extracted and catalogued '{folderName}' (source={source}).");
            return new ResolvedRelease { Version = info.Version, ReleaseDate = info.ReleaseDate, FolderName = folderName, ContractVersion = info.ContractVersion };
        }

        // ---- install-wide lock ---------------------------------------------------------

        /// <summary>One named mutex per install (per-user install folder), so two Windows users never share it.</summary>
        internal static string InstallLockName(IPathProvider paths)
        {
            string key = Path.GetFullPath(paths.VersionsPath).ToLowerInvariant();
            using (var sha = SHA256.Create())
                return @"Global\XLEdge_Install_" + string.Concat(sha.ComputeHash(Encoding.UTF8.GetBytes(key)).Take(8).Select(b => b.ToString("x2")));
        }

        private static T WithInstallLock<T>(IPathProvider paths, Func<T> work)
        {
            string key = Path.GetFullPath(paths.VersionsPath);
            string name = InstallLockName(paths);

            Mutex mutex;
            try { mutex = new Mutex(false, name); }
            catch (UnauthorizedAccessException ex) { throw new TimeoutException("Could not open the XLEdge install lock for '" + key + "'.", ex); }

            using (mutex)
            {
                bool acquired;
                try { acquired = mutex.WaitOne(InstallLockTimeout); }
                catch (AbandonedMutexException) { acquired = true; } // an Excel process died while installing; every step re-checks
                if (!acquired) throw new TimeoutException($"Another Excel process held the XLEdge install lock for more than {InstallLockTimeout.TotalMinutes:0} minutes.");
                try { return work(); }
                finally { mutex.ReleaseMutex(); }
            }
        }

        // ---- ResolveVersionToLoad helpers ---------------------------------------------

        private static string SelectedReleaseFile(IPathProvider paths) =>
            Path.Combine(Path.GetDirectoryName(paths.ReleaseHistoryFile), SelectedReleaseFileName);

        private static string ReadRemembered(IPathProvider paths)
        {
            try
            {
                string file = SelectedReleaseFile(paths);
                return File.Exists(file) ? File.ReadAllText(file).Trim() : null;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                return null;
            }
        }

        private static VersionInfo FirstManifestEntry(string manifestFile, ILogger logger) =>
            new VersionParser(logger).ParseVersionFile(manifestFile).AllVersions.FirstOrDefault();

        /// <summary>The staged zip: none without a manifest; the first manifest entry's fileName when several are staged, else the newest.</summary>
        private static string FindStagedZip(IPathProvider paths, ILogger logger)
        {
            if (!File.Exists(paths.ManifestFile)) return null;
            var zips = Directory.GetFiles(paths.ManifestDirectory, "*.zip");
            if (zips.Length == 0) return null;
            if (zips.Length == 1) return zips[0];

            string named = FirstManifestEntry(paths.ManifestFile, logger)?.FileName;
            string match = named == null ? null : zips.FirstOrDefault(z => string.Equals(Path.GetFileName(z), named, StringComparison.OrdinalIgnoreCase));
            if (match != null) return match;

            string newest = zips.OrderByDescending(File.GetLastWriteTimeUtc).First();
            logger?.LogWarn($"UpdateBootstrapper.ResolveVersionToLoad: {zips.Length} zips are staged in '{paths.ManifestDirectory}'; using the newest, '{Path.GetFileName(newest)}'.");
            return newest;
        }

        private ResolvedRelease ExtractStaged(IXLEdgeContext context, string zipPath, string source)
        {
            const string Log = "UpdateBootstrapper.ResolveVersionToLoad: ";
            ILogger logger = context.Logger;
            ResolvedRelease release;
            try
            {
                logger?.LogDebug(Log + $"extracting '{zipPath}' (source={source}).");
                release = ExtractAndCatalog(context, File.ReadAllText(context.Paths.ManifestFile), File.ReadAllBytes(zipPath), source);
            }
            catch (InvalidDataException ex)
            {
                logger?.LogError(Log + $"'{zipPath}' is not a valid zip.", ex.ToString());
                release = null;
            }
            catch (Exception ex)
            {
                // Possibly temporary (file in use, disk full): keep the zip for the next start.
                logger?.LogError(Log + $"could not install '{zipPath}'; it is kept for the next start.", ex.ToString());
                return null;
            }

            if (release != null)
            {
                // Only after the catalog write: a zip left behind is retried, a lost one is not.
                try { File.Delete(zipPath); }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) { logger?.LogWarn(Log + $"could not delete '{zipPath}': {ex.Message}"); }
            }
            else
            {
                // Refused for good (checksum, contract, invalid zip...): park it so it is not retried at every start.
                string parked = zipPath + RejectedMarker + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture);
                try
                {
                    File.Move(zipPath, parked);
                    logger?.LogWarn(Log + $"the staged release was refused and parked as '{Path.GetFileName(parked)}'.");
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    logger?.LogWarn(Log + $"could not park the refused '{zipPath}': {ex.Message}");
                }
            }
            return release;
        }

        /// <summary>
        /// The catalog; when it is missing, unreadable, or empty although releases are installed, it is rebuilt from each release
        /// folder's manifest.json snapshot. A broken catalog must never look like a first run (that clears Versions).
        /// </summary>
        private static List<ReleaseEntry> ReadOrRecoverCatalog(IPathProvider paths, ILogger logger)
        {
            const string Log = "UpdateBootstrapper.ResolveVersionToLoad: ";
            bool exists = File.Exists(paths.ReleaseHistoryFile);
            bool readable = ReleaseHistoryStore.TryReadAll(paths.ReleaseHistoryFile, out List<ReleaseEntry> entries);
            if (exists && readable && entries.Count > 0) return entries;

            var installed = SnapshotEntries(paths.VersionsPath, logger);
            if (readable && installed.Count == 0) return entries;

            logger?.LogWarn(Log + (!exists ? "the release catalog is missing" : readable ? "the release catalog is empty" : "the release catalog is unreadable") +
                $"; rebuilding it from {installed.Count} installed release(s).");
            return ReleaseHistoryStore.RebuildIfUnusable(paths.ReleaseHistoryFile, () => installed);
        }

        private static List<ReleaseEntry> SnapshotEntries(string versionsPath, ILogger logger)
        {
            var result = new List<ReleaseEntry>();
            if (!Directory.Exists(versionsPath)) return result;
            foreach (string folder in Directory.GetDirectories(versionsPath))
            {
                string name = Path.GetFileName(folder);
                string snapshot = Path.Combine(folder, "manifest.json");
                if (!IsUsableFolderName(name) || !HasDlls(folder) || !File.Exists(snapshot)) continue;
                VersionInfo info = FirstManifestEntry(snapshot, logger);
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
                logger?.LogDebug($"UpdateBootstrapper.ResolveVersionToLoad: first run - removing leftover folder '{Path.GetFileName(folder)}'.");
                TryDeleteDirectory(folder, logger);
            }
        }

        /// <summary>Removes "*.partial-*" folders left by a crashed extraction (only old ones: another Excel may be extracting).</summary>
        private static void RemoveAbandonedPartialFolders(string versionsPath, ILogger logger)
        {
            if (!Directory.Exists(versionsPath)) return;
            DateTime cutoff = DateTime.UtcNow - AbandonedPartialAge;
            foreach (string folder in Directory.GetDirectories(versionsPath, "*" + PartialMarker + "*"))
            {
                if (Directory.GetLastWriteTimeUtc(folder) < cutoff && Directory.GetCreationTimeUtc(folder) < cutoff)
                    TryDeleteDirectory(folder, logger);
            }
        }

        private static DateTime ParseReleaseDate(string value) =>
            DateTime.TryParseExact(value, ReleaseDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed) ? parsed : DateTime.MinValue;

        private static ResolvedRelease ToResolved(ReleaseEntry e) =>
            new ResolvedRelease { Version = e.Version, ReleaseDate = e.ReleaseDate, FolderName = e.FolderName, ContractVersion = e.ContractVersion };

        // ---- extraction helpers --------------------------------------------------------

        /// <summary>A safe folder name that is not an in-progress extraction folder.</summary>
        private static bool IsUsableFolderName(string folderName) =>
            ReleaseHistoryStore.IsSafeFolderName(folderName) && folderName.IndexOf(PartialMarker, StringComparison.OrdinalIgnoreCase) < 0;

        private static ResolvedRelease Refuse(ILogger logger, string reason)
        {
            logger?.LogError("UpdateBootstrapper.ExtractAndCatalog: release refused - " + reason + ".");
            return null;
        }

        internal static string Sha256(byte[] bytes)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }

        internal static bool HasDlls(string folder) =>
            Directory.Exists(folder) && Directory.EnumerateFiles(folder, "*.dll", SearchOption.TopDirectoryOnly).Any();

        /// <summary>
        /// Extracts (in memory, no temp file) into "{folder}.partial-{guid}" and renames it into place, so a failure never leaves a
        /// half-extracted release.
        /// </summary>
        private static void ExtractAtomically(string versionsPath, string versionFolder, string folderName, string manifestJson, byte[] zipBytes, ILogger logger)
        {
            Directory.CreateDirectory(versionsPath);
            string partial = Path.Combine(versionsPath, folderName + PartialMarker + Guid.NewGuid().ToString("N"));
            try
            {
                using (var stream = new MemoryStream(zipBytes, writable: false))
                using (var archive = new ZipArchive(stream, ZipArchiveMode.Read))
                    archive.ExtractToDirectory(partial);
                File.WriteAllText(Path.Combine(partial, "manifest.json"), manifestJson);

                // A leftover folder without DLLs (emptied by hand) is in the way of the rename.
                if (Directory.Exists(versionFolder) && !HasDlls(versionFolder)) Directory.Delete(versionFolder, recursive: true);
                if (!MoveDirectoryWithRetry(partial, versionFolder, logger))
                    logger?.LogWarn($"UpdateBootstrapper.ExtractAndCatalog: '{folderName}' appeared while extracting; keeping the existing folder.");
            }
            catch (Exception ex)
            {
                logger?.LogError($"UpdateBootstrapper.ExtractAndCatalog: extracting '{folderName}' failed.", ex.ToString());
                throw;
            }
            finally
            {
                TryDeleteDirectory(partial, logger);
            }
        }

        /// <summary>
        /// Renames <paramref name="source"/> to <paramref name="target"/>, retrying while antivirus or the indexer briefly holds a
        /// file inside it open. Returns false (and leaves the source) when the target already holds a release.
        /// </summary>
        internal static bool MoveDirectoryWithRetry(string source, string target, ILogger logger)
        {
            for (int attempt = 1; ; attempt++)
            {
                if (HasDlls(target)) return false;
                try
                {
                    Directory.Move(source, target);
                    return true;
                }
                catch (Exception ex) when ((ex is IOException || ex is UnauthorizedAccessException) && attempt < MoveAttempts)
                {
                    logger?.LogDebug($"UpdateBootstrapper.ExtractAndCatalog: rename of '{Path.GetFileName(source)}' failed (attempt {attempt}): {ex.Message}");
                    Thread.Sleep(MoveRetryDelay);
                }
            }
        }

        /// <summary>
        /// Moves known third-party files to SharedDependencies (first release that has them), deletes release copies identical to
        /// the shared one, and keeps a release copy that differs (that release keeps exactly what it shipped). The WebView2 files
        /// are treated as one unit.
        /// </summary>
        private static void MaterializeSharedDependencies(string versionFolder, string sharedFolder, ILogger logger)
        {
            var files = Directory.GetFiles(versionFolder, "*", SearchOption.AllDirectories);
            string Relative(string file) => file.Substring(versionFolder.Length).TrimStart(Path.DirectorySeparatorChar);

            foreach (string file in files.Where(f => SharedFileNames.Contains(Path.GetFileName(f))))
                ShareGroup(new[] { file }, Relative, sharedFolder, logger);

            var webView2 = files.Where(f => WebView2FileNames.Contains(Path.GetFileName(f))).ToArray();
            if (webView2.Length > 0) ShareGroup(webView2, Relative, sharedFolder, logger);

            RemoveEmptyDirectories(versionFolder, logger);
        }

        private static void ShareGroup(string[] files, Func<string, string> relative, string sharedFolder, ILogger logger)
        {
            const string Log = "UpdateBootstrapper.ExtractAndCatalog: ";
            try
            {
                var targets = files.Select(f => Path.Combine(sharedFolder, relative(f))).ToArray();
                bool noneShared = targets.All(t => !File.Exists(t));
                bool allIdentical = files.Zip(targets, (f, t) => File.Exists(t) && FilesHaveSameContent(f, t)).All(same => same);

                if (noneShared)
                {
                    for (int i = 0; i < files.Length; i++)
                    {
                        Directory.CreateDirectory(Path.GetDirectoryName(targets[i]));
                        string tmp = targets[i] + ".tmp";
                        File.Copy(files[i], tmp, overwrite: true);
                        File.Move(tmp, targets[i]);
                        logger?.LogDebug(Log + $"cached shared dependency '{relative(files[i])}'.");
                    }
                    foreach (string f in files) File.Delete(f);
                }
                else if (allIdentical)
                {
                    foreach (string f in files) File.Delete(f);
                }
                else
                {
                    logger?.LogWarn(Log + $"shared dependency '{string.Join("', '", files.Select(relative))}' differs from the cached copy; this release keeps its own.");
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                logger?.LogWarn(Log + $"could not share '{string.Join("', '", files.Select(relative))}': {ex.Message}");
            }
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
                    logger?.LogWarn($"UpdateBootstrapper.ExtractAndCatalog: could not remove empty folder '{dir}': {ex.Message}");
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
