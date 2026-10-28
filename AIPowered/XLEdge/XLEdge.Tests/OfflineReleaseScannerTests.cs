using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Shared;
using XLEdge.Views;

namespace XLEdge.Tests
{
    [TestClass]
    public class OfflineReleaseScannerTests
    {
        private const string Loaded = "2026-10-20T10:00:00";

        private static (string Manifest, string Zip) Write(string folder, string date = "2026-10-28T10:00:00", string checksum = "auto",
            int contract = 1, string name = "v11.1.3.0_x", string manifestName = "manifest.json")
        {
            Directory.CreateDirectory(folder);
            byte[] zip = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", name));
            string zipPath = Path.Combine(folder, name + ".zip");
            File.WriteAllBytes(zipPath, zip);
            string manifestPath = Path.Combine(folder, manifestName);
            File.WriteAllText(manifestPath, ReleaseFolders.Manifest("11.1.3.0", date, name, checksum == "auto" ? ReleaseFolders.Sha256(zip) : checksum, contract, name + ".zip"));
            return (manifestPath, zipPath);
        }

        [TestMethod]
        public void Scan_ValidNewerRelease_IsReady()
        {
            using (var t = new TempFolder())
            {
                var files = Write(t.Path);
                var r = OfflineReleaseScanner.Scan(t.Path, "11.1.2.0", Loaded);
                Assert.IsTrue(r.Ok, r.Message);
                Assert.IsFalse(r.NotNewer);
                Assert.AreEqual(files.Manifest, r.ManifestPath);
                Assert.AreEqual(files.Zip, r.ZipPath);
                Assert.AreEqual("11.1.3.0", r.Release.Version);
                long kb = (long)Math.Round(new FileInfo(files.Zip).Length / 1024.0);
                Assert.AreEqual($"Ready to reload: version 11.1.3.0, released 2026-10-28T10:00:00 ({kb} KB).", r.Message);
            }
        }

        [TestMethod]
        public void Scan_Refusals()
        {
            using (var t = new TempFolder())
            {
                string missing = t.Combine("nope");
                Assert.AreEqual("Folder not found: " + missing, OfflineReleaseScanner.Scan(missing, null, null).Message);

                string empty = t.Combine("empty");
                Directory.CreateDirectory(empty);
                Assert.AreEqual("No manifest.json + zip pair found in this folder.", OfflineReleaseScanner.Scan(empty, null, null).Message);

                string bad = t.Combine("bad");
                Write(bad);
                File.WriteAllText(Path.Combine(bad, "manifest.json"), "{garbage");
                StringAssert.StartsWith(OfflineReleaseScanner.Scan(bad, null, null).Message, "Could not parse manifest.json: ");

                string noSum = t.Combine("nosum");
                Write(noSum, checksum: null);
                Assert.AreEqual("manifest.json has no checksum recorded - cannot verify, refusing to reload.", OfflineReleaseScanner.Scan(noSum, null, null).Message);

                string wrong = t.Combine("wrong");
                var w = Write(wrong, checksum: "00");
                Assert.AreEqual($"Checksum mismatch - the zip may be corrupt or incomplete. Expected 00, got {ReleaseFolders.Sha256(File.ReadAllBytes(w.Zip))}.",
                    OfflineReleaseScanner.Scan(wrong, null, null).Message);

                string newer = t.Combine("newer");
                Write(newer, contract: ContractInfo.Version + 1);
                StringAssert.Contains(OfflineReleaseScanner.Scan(newer, null, null).Message, "needs a newer XLEdge installer");

                foreach (var dir in new[] { missing, empty, bad, noSum, wrong, newer })
                    Assert.IsFalse(OfflineReleaseScanner.Scan(dir, null, null).Ok, dir);
            }
        }

        [TestMethod]
        public void Scan_NotNewerThanTheLoadedRelease()
        {
            using (var t = new TempFolder())
            {
                Write(t.Path, date: Loaded);
                var r = OfflineReleaseScanner.Scan(t.Path, "11.1.2.0", Loaded);
                Assert.IsFalse(r.Ok);
                Assert.IsTrue(r.NotNewer);
                Assert.AreEqual($"No updates available.\n\nYou already have the latest version: 11.1.2.0 (released {Loaded}).\n\nChecked release: 11.1.3.0 ({Loaded}).", r.Message);
                Assert.IsTrue(OfflineReleaseScanner.Scan(t.Path, "11.1.2.0", "not a date").Ok, "unparseable loaded date counts as older");
                Assert.IsTrue(OfflineReleaseScanner.Scan(t.Path, null, null).Ok, "nothing loaded");
            }
        }

        [TestMethod]
        public void Scan_SeveralCandidates_UsesTheNewestPair()
        {
            using (var t = new TempFolder())
            {
                var old = Write(t.Path, name: "v11.1.1.0_old", manifestName: "manifest (1).json");
                File.SetLastWriteTimeUtc(old.Zip, DateTime.UtcNow.AddHours(-2));
                File.SetLastWriteTimeUtc(old.Manifest, DateTime.UtcNow.AddHours(-2));
                var newest = Write(t.Path, name: "v11.1.3.0_new");
                var r = OfflineReleaseScanner.Scan(t.Path, null, null);
                Assert.AreEqual(newest.Zip, r.ZipPath);
                Assert.AreEqual(newest.Manifest, r.ManifestPath);
            }
        }

        [TestMethod]
        public void Scan_PrefersTheZipTheManifestNames()
        {
            using (var t = new TempFolder())
            {
                var named = Write(t.Path);
                File.WriteAllBytes(Path.Combine(t.Path, "v99.9_unrelated.zip"), new byte[] { 1, 2, 3 });
                File.SetLastWriteTimeUtc(named.Zip, DateTime.UtcNow.AddHours(-1));
                var r = OfflineReleaseScanner.Scan(t.Path, null, null);
                Assert.IsTrue(r.Ok, r.Message);
                Assert.AreEqual(named.Zip, r.ZipPath);
            }
        }

        [TestMethod]
        public void Stage_CopiesThePairAndRemovesOtherStagedZips()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                File.WriteAllText(Path.Combine(paths.ManifestDirectory, "stale.zip"), "x");
                Write(t.Combine("offline"));
                var r = OfflineReleaseScanner.Scan(t.Combine("offline"), null, null);

                OfflineReleaseScanner.Stage(r, paths);

                CollectionAssert.AreEqual(new[] { "v11.1.3.0_x.zip" }, Directory.GetFiles(paths.ManifestDirectory, "*.zip").Select(Path.GetFileName).ToArray());
                Assert.AreEqual(File.ReadAllText(r.ManifestPath), File.ReadAllText(paths.ManifestFile));
            }
        }

        [TestMethod]
        public void Stage_FromTheManifestFolderItself_KeepsTheCandidate()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                Write(paths.ManifestDirectory);
                var r = OfflineReleaseScanner.Scan(paths.ManifestDirectory, null, null);
                OfflineReleaseScanner.Stage(r, paths);
                Assert.IsTrue(File.Exists(r.ZipPath));
                Assert.IsTrue(File.Exists(paths.ManifestFile));
            }
        }

        [TestMethod]
        public void Stage_RefusedScan_Throws()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                var r = OfflineReleaseScanner.Scan(t.Combine("missing"), null, null);
                Assert.ThrowsExactly<InvalidOperationException>(() => OfflineReleaseScanner.Stage(r, paths));
            }
        }
    }
}
