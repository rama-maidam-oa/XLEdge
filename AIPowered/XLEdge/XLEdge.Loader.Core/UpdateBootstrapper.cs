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
