using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace XLEdge.Shared
{
    public enum ReleaseDeleteResult
    {
        Deleted,
        NotFound,
        ActiveRelease,
        InvalidFolderName,
        FolderDeleteFailed,
        CatalogWriteFailed,
    }

    /// <summary>
    /// ReleaseHistory.json: a JSON array of <see cref="ReleaseEntry"/>. Every read and write runs under a
    /// named mutex (several Excel processes can start at once) and writes are atomic (temp file, then
    /// File.Replace / File.Move), so a crash never leaves a half-written catalog.
    /// </summary>
    public static class ReleaseHistoryStore
    {
        private const string LockNamePrefix = @"Global\XLEdge_ReleaseHistory_";
        private static readonly TimeSpan LockTimeout = TimeSpan.FromSeconds(10);
        private const int ReplaceAttempts = 10;
        private static readonly TimeSpan ReplaceRetryDelay = TimeSpan.FromMilliseconds(100);
        private static readonly JsonSerializerOptions ReadOptions = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        private static readonly JsonSerializerOptions WriteOptions = new JsonSerializerOptions { WriteIndented = true };

        /// <summary>All entries; empty when the file is missing, empty or "null". A corrupt file throws JsonException.</summary>
        public static List<ReleaseEntry> ReadAll(string releaseHistoryFile) => ReadAll(releaseHistoryFile, LockTimeout);

        internal static List<ReleaseEntry> ReadAll(string releaseHistoryFile, TimeSpan lockTimeout) =>
            WithLock(releaseHistoryFile, lockTimeout, () => ReadAllUnlocked(releaseHistoryFile));

        /// <summary>Adds the entry unless one with the same Version, ReleaseDate and FolderName already exists.</summary>
        public static void Append(string releaseHistoryFile, ReleaseEntry entry)
        {
            if (entry == null) throw new ArgumentNullException(nameof(entry));
            WithLock(releaseHistoryFile, LockTimeout, () =>
            {
                var entries = ReadAllUnlocked(releaseHistoryFile);
                bool exists = entries.Any(e =>
                    string.Equals(e.Version, entry.Version, StringComparison.Ordinal) &&
                    string.Equals(e.ReleaseDate, entry.ReleaseDate, StringComparison.Ordinal) &&
                    string.Equals(e.FolderName, entry.FolderName, StringComparison.Ordinal));
                if (!exists)
                {
                    entries.Add(entry);
                    WriteAllUnlocked(releaseHistoryFile, entries);
                }
                return 0;
            });
        }

        /// <summary>
        /// Drops entries whose release folder is gone or holds no DLL (deleted by hand, or a failed extraction),
        /// rewrites the catalog only when something was dropped, and returns the remaining entries.
        /// </summary>
        public static List<ReleaseEntry> Reconcile(string releaseHistoryFile, string versionsPath) =>
            WithLock(releaseHistoryFile, LockTimeout, () =>
            {
                var entries = ReadAllUnlocked(releaseHistoryFile);
                var kept = entries.Where(e => HasReleaseDlls(versionsPath, e.FolderName)).ToList();
                if (kept.Count != entries.Count) WriteAllUnlocked(releaseHistoryFile, kept);
                return kept;
            });

        /// <summary>Deletes a release folder and its catalog entry. The active release is never deleted.</summary>
        public static ReleaseDeleteResult Delete(string releaseHistoryFile, string versionsPath, string folderName, string activeFolderName, out Exception error)
        {
            Exception failure = null;
            var result = WithLock(releaseHistoryFile, LockTimeout, () =>
            {
                var entries = ReadAllUnlocked(releaseHistoryFile);
                if (!entries.Any(e => string.Equals(e.FolderName, folderName, StringComparison.OrdinalIgnoreCase)))
                    return ReleaseDeleteResult.NotFound;
                if (!string.IsNullOrWhiteSpace(activeFolderName) && string.Equals(folderName, activeFolderName, StringComparison.OrdinalIgnoreCase))
                    return ReleaseDeleteResult.ActiveRelease;
                if (!IsSafeFolderName(folderName))
                    return ReleaseDeleteResult.InvalidFolderName;

                string target = ResolveReleaseFolder(versionsPath, folderName);
                if (target == null)
                    return ReleaseDeleteResult.InvalidFolderName;
                string active = ResolveReleaseFolder(versionsPath, activeFolderName);
                if (active != null && string.Equals(target, active, StringComparison.OrdinalIgnoreCase))
                    return ReleaseDeleteResult.ActiveRelease;

                try
                {
                    if (Directory.Exists(target)) Directory.Delete(target, recursive: true);
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    failure = ex;
                    return ReleaseDeleteResult.FolderDeleteFailed;
                }

                try
                {
                    WriteAllUnlocked(releaseHistoryFile, entries.Where(e => !string.Equals(e.FolderName, folderName, StringComparison.OrdinalIgnoreCase)).ToList());
                }
                catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                {
                    // The folder is already gone; the next Reconcile drops the stale entry.
                    failure = ex;
                    return ReleaseDeleteResult.CatalogWriteFailed;
                }
                return ReleaseDeleteResult.Deleted;
            });
            error = failure;
            return result;
        }

        /// <summary>"v{version}_{releaseDate}" with characters that are invalid in file names replaced by '-'.</summary>
        public static string BuildFolderName(string version, string releaseDate)
        {
            var invalid = Path.GetInvalidFileNameChars();
            string safeDate = new string((releaseDate ?? string.Empty).Select(c => invalid.Contains(c) ? '-' : c).ToArray());
            return $"v{version}_{safeDate}";
        }

        /// <summary>
        /// A bare folder name: no path separators or invalid characters, not rooted, and not ending in '.' or ' '
        /// (Windows drops those, so "v1." would resolve to the folder "v1").
        /// </summary>
        public static bool IsSafeFolderName(string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName) || folderName.EndsWith(".", StringComparison.Ordinal) || folderName.EndsWith(" ", StringComparison.Ordinal)) return false;
            if (folderName.IndexOf('\\') >= 0 || folderName.IndexOf('/') >= 0) return false;
            if (folderName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return false;
            return !Path.IsPathRooted(folderName) && Path.GetFileName(folderName) == folderName;
        }

        /// <summary>
        /// Full path of a release folder directly inside <paramref name="versionsPath"/>, or null when the name
        /// is unusable or resolves anywhere else (outside the versions folder, or the versions folder itself).
        /// </summary>
        internal static string ResolveReleaseFolder(string versionsPath, string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName)) return null;
            try
            {
                string root = Path.GetFullPath(versionsPath).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                string target = Path.GetFullPath(Path.Combine(versionsPath, folderName)).TrimEnd(Path.DirectorySeparatorChar);
                if (!target.StartsWith(root, StringComparison.OrdinalIgnoreCase) || target.Length <= root.Length) return null;
                if (target.IndexOf(Path.DirectorySeparatorChar, root.Length) >= 0) return null; // nested, not a release folder
                return target;
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                return null;
            }
        }

        /// <summary>
        /// One mutex per catalog file (per-user install folder), so two Windows users never share it; a single
        /// machine-wide name would make one user's lock unopenable (access denied) for another.
        /// </summary>
        internal static string LockName(string releaseHistoryFile)
        {
            string key = Path.GetFullPath(releaseHistoryFile).ToLowerInvariant();
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(key));
                return LockNamePrefix + string.Concat(hash.Take(8).Select(b => b.ToString("x2")));
            }
        }

        private static bool HasReleaseDlls(string versionsPath, string folderName)
        {
            if (string.IsNullOrWhiteSpace(folderName)) return false;
            string folder = Path.Combine(versionsPath, folderName);
            return Directory.Exists(folder) && Directory.EnumerateFiles(folder, "*.dll", SearchOption.TopDirectoryOnly).Any();
        }

        private static T WithLock<T>(string releaseHistoryFile, TimeSpan timeout, Func<T> work)
        {
            Mutex mutex;
            try
            {
                mutex = new Mutex(false, LockName(releaseHistoryFile));
            }
            catch (UnauthorizedAccessException ex)
            {
                throw LockTimeoutException(releaseHistoryFile, timeout, ex);
            }

            using (mutex)
            {
                bool acquired;
                try
                {
                    acquired = mutex.WaitOne(timeout);
                }
                catch (AbandonedMutexException)
                {
                    // A process died holding the lock; writes are atomic, so the file is consistent.
                    acquired = true;
                }
                if (!acquired) throw LockTimeoutException(releaseHistoryFile, timeout, null);

                try
                {
                    return work();
                }
                finally
                {
                    mutex.ReleaseMutex();
                }
            }
        }

        private static TimeoutException LockTimeoutException(string file, TimeSpan timeout, Exception inner) =>
            new TimeoutException($"Could not acquire the release history lock for '{file}' within {timeout} - refusing to read/write it unprotected.", inner);

        private static List<ReleaseEntry> ReadAllUnlocked(string file)
        {
            if (!File.Exists(file)) return new List<ReleaseEntry>();
            string json = File.ReadAllText(file);
            if (string.IsNullOrWhiteSpace(json)) return new List<ReleaseEntry>();
            return JsonSerializer.Deserialize<List<ReleaseEntry>>(json, ReadOptions) ?? new List<ReleaseEntry>();
        }

        private static void WriteAllUnlocked(string file, List<ReleaseEntry> entries)
        {
            string folder = Path.GetDirectoryName(file);
            if (!string.IsNullOrEmpty(folder)) Directory.CreateDirectory(folder);

            string tmp = file + ".tmp";
            byte[] content = new UTF8Encoding(false).GetBytes(JsonSerializer.Serialize(entries, WriteOptions));
            using (var stream = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                stream.Write(content, 0, content.Length);
                // On disk before the swap: after a power cut the catalog is the old or the new one, never empty.
                stream.Flush(flushToDisk: true);
            }

            // Antivirus and indexers open a freshly written file for a moment; File.Replace then fails with a
            // sharing violation or access denied. Retry briefly while the (complete) temp file is still there.
            for (int attempt = 1; ; attempt++)
            {
                try
                {
                    if (File.Exists(file)) File.Replace(tmp, file, null);
                    else File.Move(tmp, file);
                    return;
                }
                catch (Exception ex) when ((ex is IOException || ex is UnauthorizedAccessException) && attempt < ReplaceAttempts && File.Exists(tmp))
                {
                    Thread.Sleep(ReplaceRetryDelay);
                }
            }
        }
    }
}
