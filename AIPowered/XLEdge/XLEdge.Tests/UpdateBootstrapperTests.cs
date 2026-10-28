using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class UpdateBootstrapperTests
    {
        private const string Date = "2026-10-28T10:00:00";

        private sealed class Fixture : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();

            public Fixture()
            {
                Paths = new PathProvider(_temp.Combine("inst"), _temp.Combine("data"));
                Paths.Ensure();
                Log = new RecordingLogger();
                Context = new FakeContext(Paths) { Logger = Log };
            }

            public PathProvider Paths { get; }
            public RecordingLogger Log { get; }
            public FakeContext Context { get; }
            public UpdateBootstrapper Bootstrapper { get; } = new UpdateBootstrapper();

            public string[] ReleaseFolders() => Directory.GetDirectories(Paths.VersionsPath).Select(Path.GetFileName).OrderBy(n => n).ToArray();

            public ResolvedRelease Extract(byte[] zip, string folder, string version = "11.1.2.0", int contract = 1, string checksum = "auto", string source = "Offline", string date = Date) =>
                Bootstrapper.ExtractAndCatalog(Context, XLEdge.Tests.ReleaseFolders.Manifest(version, date, folder, checksum == "auto" ? XLEdge.Tests.ReleaseFolders.Sha256(zip) : checksum, contract), zip, source);

            public void Dispose() => _temp.Dispose();
        }

        private static byte[] CoreZip(string marker = "core") =>
            ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", marker), ("XLEdge.Contracts.dll", "contracts"));

        // ---- ExtractAndCatalog ---------------------------------------------------------

        [TestMethod]
        public void Extract_CataloguesAndSnapshotsManifest()
        {
            using (var f = new Fixture())
            {
                byte[] zip = CoreZip();
                string manifest = ReleaseFolders.Manifest("11.1.2.0", Date, "v11.1.2.0_a", ReleaseFolders.Sha256(zip));
                var release = f.Bootstrapper.ExtractAndCatalog(f.Context, manifest, zip, "Offline");

                Assert.IsNotNull(release, f.Log.Dump());
                Assert.AreEqual("v11.1.2.0_a", release.FolderName);
                Assert.AreEqual("11.1.2.0", release.Version);
                Assert.AreEqual(1, release.ContractVersion);
                string folder = Path.Combine(f.Paths.VersionsPath, "v11.1.2.0_a");
                Assert.AreEqual("core", File.ReadAllText(Path.Combine(folder, "XLEdge.Addin.Core.dll")));
                Assert.AreEqual(manifest, File.ReadAllText(Path.Combine(folder, "manifest.json")));
                CollectionAssert.AreEqual(new[] { "v11.1.2.0_a" }, f.ReleaseFolders(), "no partial folder may remain");

                var entry = ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Single();
                Assert.AreEqual("Offline", entry.Source);
                Assert.AreEqual(1, entry.ContractVersion);
                Assert.AreEqual(ReleaseFolders.Sha256(zip), entry.Checksum);
            }
        }

        [TestMethod]
        public void Extract_WithoutFolderName_UsesBuildFolderName()
        {
            using (var f = new Fixture())
            {
                var release = f.Extract(CoreZip(), folder: null);
                Assert.AreEqual("v11.1.2.0_2026-10-28T10-00-00", release?.FolderName, f.Log.Dump());
            }
        }

        [TestMethod]
        public void Extract_Refusals_LeaveVersionsAndCatalogUntouched()
        {
            using (var f = new Fixture())
            {
                byte[] zip = CoreZip();
                f.Context.ActiveFolderName = "vactive";
                var cases = new (string Name, Func<ResolvedRelease> Run)[]
                {
                    ("parent", () => f.Extract(zip, "..")),
                    ("trailing dot", () => f.Extract(zip, "v1.")),
                    ("active folder", () => f.Extract(zip, "VACTIVE")),
                    ("no checksum", () => f.Extract(zip, "v1", checksum: null)),
                    ("wrong checksum", () => f.Extract(zip, "v1", checksum: "00")),
                    ("contract 0", () => f.Extract(zip, "v1", contract: 0)),
                    ("contract too new", () => f.Extract(zip, "v1", contract: ContractInfo.Version + 1)),
                    ("no version", () => f.Extract(zip, "v1", version: "")),
                    ("not a manifest", () => f.Bootstrapper.ExtractAndCatalog(f.Context, "{}", zip, "Offline")),
                };
                foreach (var c in cases)
                {
                    Assert.IsNull(c.Run(), c.Name);
                    Assert.AreEqual(0, f.ReleaseFolders().Length, c.Name);
                    Assert.IsFalse(File.Exists(f.Paths.ReleaseHistoryFile), c.Name);
                }
                Assert.IsTrue(f.Log.Contains("needs a newer XLEdge installer"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Extract_AlreadyExtracted_IsIdempotent()
        {
            using (var f = new Fixture())
            {
                byte[] zip = CoreZip();
                var first = f.Extract(zip, "v1");
                var second = f.Extract(zip, "v1");
                Assert.AreEqual(first.FolderName, second?.FolderName, f.Log.Dump());
                Assert.AreEqual(1, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Count);
                CollectionAssert.AreEqual(new[] { "v1" }, f.ReleaseFolders());
            }
        }

        [TestMethod]
        public void Extract_BadZip_LeavesNoPartialFolder()
        {
            using (var f = new Fixture())
            {
                byte[] garbage = { 1, 2, 3, 4, 5 };
                Assert.ThrowsExactly<InvalidDataException>(() => f.Extract(garbage, "v1"));
                Assert.AreEqual(0, f.ReleaseFolders().Length);
                Assert.IsFalse(File.Exists(f.Paths.ReleaseHistoryFile));
            }
        }

        // ---- ResolveVersionToLoad -------------------------------------------------------

        private static void StageInManifestFolder(Fixture f, byte[] zip, string folder, string version = "11.1.2.0", string date = Date,
            string zipName = "OrbitXLEdge.zip", string checksum = "auto", string fileName = null)
        {
            Directory.CreateDirectory(f.Paths.ManifestDirectory);
            File.WriteAllText(f.Paths.ManifestFile, ReleaseFolders.Manifest(version, date, folder, checksum == "auto" ? ReleaseFolders.Sha256(zip) : checksum, 1, fileName));
            File.WriteAllBytes(Path.Combine(f.Paths.ManifestDirectory, zipName), zip);
        }

        [TestMethod]
        public void FreshInstall_Seed_ExtractsCataloguesAndRemovesManifest()
        {
            using (var f = new Fixture())
            {
                StageInManifestFolder(f, CoreZip(), "v1");
                var release = f.Bootstrapper.ResolveVersionToLoad(f.Context);
                Assert.AreEqual("v1", release?.FolderName, f.Log.Dump());
                Assert.AreEqual("Install", ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Single().Source);
                Assert.IsFalse(Directory.Exists(f.Paths.ManifestDirectory), "seed consumed");
            }
        }

        [TestMethod]
        public void FreshInstall_BadSeed_KeepsManifestAndReturnsNull()
        {
            using (var f = new Fixture())
            {
                StageInManifestFolder(f, CoreZip(), "v1", checksum: "00");
                Assert.IsNull(f.Bootstrapper.ResolveVersionToLoad(f.Context));
                Assert.IsTrue(File.Exists(f.Paths.ManifestFile), "seed kept for the next start");
                Assert.IsTrue(File.Exists(Path.Combine(f.Paths.ManifestDirectory, "OrbitXLEdge.zip")));
            }
        }

        [TestMethod]
        public void FreshInstall_WipesStrayVersions()
        {
            using (var f = new Fixture())
            {
                Directory.CreateDirectory(Path.Combine(f.Paths.VersionsPath, "stray"));
                File.WriteAllText(Path.Combine(f.Paths.VersionsPath, "stray", "x.dll"), "x");
                StageInManifestFolder(f, CoreZip(), "v1");
                f.Bootstrapper.ResolveVersionToLoad(f.Context);
                CollectionAssert.AreEqual(new[] { "v1" }, f.ReleaseFolders());
            }
        }

        [TestMethod]
        public void Staged_ZipAndManifest_AddsSecondRelease_DeletesZipKeepsManifest()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("one"), "v1");
                StageInManifestFolder(f, CoreZip("two"), "v2", version: "11.1.3.0", date: "2026-10-29T10:00:00");
                var release = f.Bootstrapper.ResolveVersionToLoad(f.Context, "Offline");
                Assert.AreEqual("v2", release?.FolderName, f.Log.Dump());
                Assert.AreEqual(2, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Count);
                Assert.AreEqual("Offline", ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Last().Source);
                Assert.AreEqual(0, Directory.GetFiles(f.Paths.ManifestDirectory, "*.zip").Length, "zip consumed");
                Assert.IsTrue(File.Exists(f.Paths.ManifestFile));
            }
        }

        [TestMethod]
        public void Staged_SeveralZips_UsesManifestFileName()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("one"), "v1");
                StageInManifestFolder(f, CoreZip("two"), "v2", version: "11.1.3.0", zipName: "v2.zip", fileName: "v2.zip");
                File.WriteAllBytes(Path.Combine(f.Paths.ManifestDirectory, "zzz-newer.zip"), new byte[] { 1, 2, 3 });
                Assert.AreEqual("v2", f.Bootstrapper.ResolveVersionToLoad(f.Context)?.FolderName, f.Log.Dump());
            }
        }

        [TestMethod]
        public void Staged_ReleaseRefused_FallsBackToTheCataloguedRelease()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("one"), "v1");
                StageInManifestFolder(f, CoreZip("two"), "v2", version: "11.1.3.0", checksum: "00");
                Assert.AreEqual("v1", f.Bootstrapper.ResolveVersionToLoad(f.Context)?.FolderName, f.Log.Dump());
            }
        }

        [TestMethod]
        public void KnownReinstall_Reconciles()
        {
            using (var f = new Fixture())
            {
                byte[] zip = CoreZip();
                f.Extract(zip, "v1");
                ReleaseHistoryStore.Append(f.Paths.ReleaseHistoryFile, new ReleaseEntry { Version = "9.9", ReleaseDate = "x", FolderName = "gone", ContractVersion = 1 });
                StageInManifestFolder(f, zip, "v1");
                Assert.AreEqual("v1", f.Bootstrapper.ResolveVersionToLoad(f.Context)?.FolderName, f.Log.Dump());
                CollectionAssert.AreEqual(new[] { "v1" }, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Select(e => e.FolderName).ToArray());
            }
        }

        [TestMethod]
        public void NoZip_NewestLoadableCataloguedRelease()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("old"), "vOld", date: "2026-10-01T10:00:00");
                f.Extract(CoreZip("new"), "vNew", date: "2026-10-15T10:00:00");
                Directory.CreateDirectory(Path.Combine(f.Paths.VersionsPath, "vNoDll"));
                ReleaseHistoryStore.Append(f.Paths.ReleaseHistoryFile, new ReleaseEntry { Version = "11.1.4.0", ReleaseDate = "2026-10-20T10:00:00", FolderName = "vNoDll", ContractVersion = 1 });
                Directory.CreateDirectory(Path.Combine(f.Paths.VersionsPath, "vTooNew"));
                File.WriteAllText(Path.Combine(f.Paths.VersionsPath, "vTooNew", "XLEdge.Addin.Core.dll"), "x");
                ReleaseHistoryStore.Append(f.Paths.ReleaseHistoryFile, new ReleaseEntry { Version = "12.0.0.0", ReleaseDate = "2026-10-25T10:00:00", FolderName = "vTooNew", ContractVersion = ContractInfo.Version + 1 });

                var release = f.Bootstrapper.ResolveVersionToLoad(f.Context);
                Assert.AreEqual("vNew", release?.FolderName, f.Log.Dump());
                Assert.AreEqual(1, release.ContractVersion);
            }
        }

        [TestMethod]
        public void CorruptOrEmptyCatalog_RecoveredFromSnapshots_VersionsKept()
        {
            foreach (var content in new[] { "[{\"Version\":", "", "[]" })
            {
                using (var f = new Fixture())
                {
                    f.Extract(CoreZip("old"), "vOld", date: "2026-10-01T10:00:00");
                    f.Extract(CoreZip("new"), "vNew", version: "11.1.3.0", date: "2026-10-15T10:00:00");
                    File.WriteAllText(f.Paths.ReleaseHistoryFile, content);

                    var release = f.Bootstrapper.ResolveVersionToLoad(f.Context);

                    Assert.AreEqual("vNew", release?.FolderName, "content '" + content + "':\n" + f.Log.Dump());
                    CollectionAssert.AreEqual(new[] { "vNew", "vOld" }, f.ReleaseFolders(), "releases must survive");
                    var rebuilt = ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile);
                    CollectionAssert.AreEquivalent(new[] { "vOld", "vNew" }, rebuilt.Select(e => e.FolderName).ToArray());
                    Assert.IsTrue(rebuilt.All(e => e.Source == "Recovered" && e.ContractVersion == 1));
                    Assert.AreEqual(1, Directory.GetFiles(Path.GetDirectoryName(f.Paths.ReleaseHistoryFile), "ReleaseHistory.corrupt-*.json").Length);
                }
            }
        }

        [TestMethod]
        public void StalePartialFolders_AreRemoved()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip(), "v1");
                Directory.CreateDirectory(Path.Combine(f.Paths.VersionsPath, "v2.partial-0123456789abcdef0123456789abcdef"));
                f.Bootstrapper.ResolveVersionToLoad(f.Context);
                CollectionAssert.AreEqual(new[] { "v1" }, f.ReleaseFolders());
            }
        }

        [TestMethod]
        public void Nothing_ReturnsNull()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(f.Bootstrapper.ResolveVersionToLoad(f.Context));
                Assert.IsTrue(f.Log.Contains("nothing to load"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void SharedDependencies_MovedDedupedOrKept()
        {
            using (var f = new Fixture())
            {
                byte[] a = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", "a"), ("XLEdge.Shared.dll", "s"), ("NLog.dll", "nlog-1"), (@"x64\WebView2Loader.dll", "native"));
                byte[] b = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", "b"), ("NLog.dll", "nlog-1"));
                byte[] c = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", "c"), ("NLog.dll", "nlog-2"));
                f.Extract(a, "vA");
                f.Extract(b, "vB");
                f.Extract(c, "vC");

                string shared = f.Paths.SharedDependenciesPath, versions = f.Paths.VersionsPath;
                Assert.AreEqual("nlog-1", File.ReadAllText(Path.Combine(shared, "NLog.dll")));
                Assert.AreEqual("native", File.ReadAllText(Path.Combine(shared, "x64", "WebView2Loader.dll")));
                Assert.IsFalse(File.Exists(Path.Combine(versions, "vA", "NLog.dll")), "moved to shared");
                Assert.IsFalse(Directory.Exists(Path.Combine(versions, "vA", "x64")), "empty folder removed");
                Assert.IsTrue(File.Exists(Path.Combine(versions, "vA", "XLEdge.Shared.dll")), "XLEdge assemblies are never shared");
                Assert.IsFalse(File.Exists(Path.Combine(versions, "vB", "NLog.dll")), "identical copy removed");
                Assert.AreEqual("nlog-2", File.ReadAllText(Path.Combine(versions, "vC", "NLog.dll")), "different copy kept for rollback isolation");
                Assert.IsTrue(f.Log.Contains("shared dependency 'NLog.dll' differs"), f.Log.Dump());
            }
        }
    }
}
