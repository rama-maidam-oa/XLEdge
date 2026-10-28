using System;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;

namespace XLEdge.Tests
{
    /// <summary>Builds release folders, zips and manifests for loader and bootstrapper tests.</summary>
    internal static class ReleaseFolders
    {
        /// <summary>XLEdge.Addin.Core\bin\&lt;same configuration as the tests&gt;\ — the real core build output.</summary>
        public static string CoreBuildFolder
        {
            get
            {
                string testBin = Path.GetDirectoryName(typeof(ReleaseFolders).Assembly.Location);
                string config = new DirectoryInfo(testBin).Name;
                string folder = Path.GetFullPath(Path.Combine(testBin, @"..\..\..\XLEdge.Addin.Core\bin", config));
                if (!File.Exists(Path.Combine(folder, "XLEdge.Addin.Core.dll")))
                    throw new InvalidOperationException("XLEdge.Addin.Core is not built: " + folder);
                return folder;
            }
        }

        /// <summary>Copies the real core build (DLLs + config) into versionsPath\folderName and returns that folder.</summary>
        public static string Stage(string versionsPath, string folderName)
        {
            string target = Path.Combine(versionsPath, folderName);
            Directory.CreateDirectory(target);
            foreach (var file in Directory.GetFiles(CoreBuildFolder))
            {
                string ext = Path.GetExtension(file).ToLowerInvariant();
                if (ext == ".dll" || ext == ".config")
                    File.Copy(file, Path.Combine(target, Path.GetFileName(file)), overwrite: true);
            }
            return target;
        }

        /// <summary>Zips a folder's files (keeping relative paths) and returns the bytes.</summary>
        public static byte[] Zip(string folder)
        {
            string zipPath = Path.Combine(Path.GetTempPath(), "XLEdge.Tests", Guid.NewGuid().ToString("N") + ".zip");
            Directory.CreateDirectory(Path.GetDirectoryName(zipPath));
            try
            {
                ZipFile.CreateFromDirectory(folder, zipPath);
                return File.ReadAllBytes(zipPath);
            }
            finally
            {
                File.Delete(zipPath);
            }
        }

        /// <summary>A zip with fake files: name → content.</summary>
        public static byte[] FakeZip(params (string Path, string Content)[] files)
        {
            string folder = Path.Combine(Path.GetTempPath(), "XLEdge.Tests", Guid.NewGuid().ToString("N"));
            try
            {
                foreach (var f in files)
                {
                    string p = Path.Combine(folder, f.Path);
                    Directory.CreateDirectory(Path.GetDirectoryName(p));
                    File.WriteAllText(p, f.Content);
                }
                return Zip(folder);
            }
            finally
            {
                Directory.Delete(folder, recursive: true);
            }
        }

        public static string Sha256(byte[] bytes)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }

        /// <summary>One-entry manifest JSON (the array the build writes).</summary>
        public static string Manifest(string version, string releaseDate, string folderName, string checksum, int contractVersion = 1, string fileName = null) =>
            "[{" +
            $"\"version\":\"{version}\",\"releaseDate\":\"{releaseDate}\"," +
            (folderName == null ? "" : $"\"folderName\":\"{folderName}\",") +
            (fileName == null ? "" : $"\"fileName\":\"{fileName}\",") +
            (checksum == null ? "" : $"\"checksum\":\"{checksum}\",") +
            $"\"notes\":\"test\",\"contractVersion\":{contractVersion}" +
            "}]";
    }
}
