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
        public void FreshInstall_RefusedSeed_IsParkedAndNotRetried()
        {
            using (var f = new Fixture())
            {
                StageInManifestFolder(f, CoreZip(), "v1", checksum: "00");
                Assert.IsNull(f.Bootstrapper.ResolveVersionToLoad(f.Context));
                Assert.IsTrue(File.Exists(f.Paths.ManifestFile), "manifest kept for diagnosis");
                Assert.AreEqual(0, Directory.GetFiles(f.Paths.ManifestDirectory, "*.zip").Length);
                Assert.AreEqual(1, Directory.GetFiles(f.Paths.ManifestDirectory, "OrbitXLEdge.zip.rejected-*").Length);

                int refusals = f.Log.Lines.Count(l => l.Contains("release refused"));
                Assert.IsNull(f.Bootstrapper.ResolveVersionToLoad(f.Context));
                Assert.AreEqual(refusals, f.Log.Lines.Count(l => l.Contains("release refused")), "a parked zip is not retried");
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
                Assert.AreEqual(0, Directory.GetFiles(f.Paths.ManifestDirectory, "*.zip").Length, "refused zip parked");
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
                string old = Path.Combine(f.Paths.VersionsPath, "v2.partial-0123456789abcdef0123456789abcdef");
                string young = Path.Combine(f.Paths.VersionsPath, "v3.partial-fedcba9876543210fedcba9876543210");
                Directory.CreateDirectory(old);
                Directory.CreateDirectory(young);
                Directory.SetCreationTimeUtc(old, DateTime.UtcNow.AddHours(-1));
                Directory.SetLastWriteTimeUtc(old, DateTime.UtcNow.AddHours(-1));
                f.Bootstrapper.ResolveVersionToLoad(f.Context);
                CollectionAssert.AreEqual(new[] { "v1", Path.GetFileName(young) }, f.ReleaseFolders(), "only abandoned partial folders are removed");
            }
        }

        [TestMethod]
        public void Nothing_ReturnsNull()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(f.Bootstrapper.ResolveVersionToLoad(f.Context));
                Assert.IsTrue(f.Log.Contains("UpdateBootstrapper.ResolveVersionToLoad: no staged release"), f.Log.Dump());
                Assert.IsTrue(f.Log.Contains("nothing to load"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void MissingCatalog_WithInstalledReleases_IsRecoveredNotTreatedAsFirstRun()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip(), "v1");
                File.Delete(f.Paths.ReleaseHistoryFile);
                Assert.AreEqual("v1", f.Bootstrapper.ResolveVersionToLoad(f.Context)?.FolderName, f.Log.Dump());
                Assert.AreEqual("v1", ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Single().FolderName);
            }
        }

        [TestMethod]
        public void MissingCatalog_WithInstalledReleasesAndAStagedUpdate_KeepsTheReleases()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("one"), "v1");
                File.Delete(f.Paths.ReleaseHistoryFile);
                StageInManifestFolder(f, CoreZip("two"), "v2", version: "11.1.3.0", date: "2026-10-29T10:00:00");
                Assert.AreEqual("v2", f.Bootstrapper.ResolveVersionToLoad(f.Context, "Offline")?.FolderName, f.Log.Dump());
                CollectionAssert.AreEqual(new[] { "v1", "v2" }, f.ReleaseFolders(), "installed releases must never be wiped");
                Assert.AreEqual(2, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Count);
            }
        }

        [TestMethod]
        public void TwoExcelProcesses_FirstRunAtTheSameTime_InstallOnce()
        {
            for (int round = 0; round < 5; round++)
            {
                using (var f = new Fixture())
                {
                    StageInManifestFolder(f, CoreZip(), "v1");
                    var second = new FakeContext(new PathProvider(Path.GetDirectoryName(Path.GetDirectoryName(f.Paths.VersionsPath)),
                        Path.GetDirectoryName(f.Paths.Root))) { Logger = f.Log };
                    Assert.AreEqual(f.Paths.VersionsPath, second.Paths.VersionsPath);

                    var results = new ResolvedRelease[2];
                    using (var start = new System.Threading.Barrier(2))
                    {
                        var a = new System.Threading.Thread(() => { start.SignalAndWait(); results[0] = new UpdateBootstrapper().ResolveVersionToLoad(f.Context); });
                        var b = new System.Threading.Thread(() => { start.SignalAndWait(); results[1] = new UpdateBootstrapper().ResolveVersionToLoad(second); });
                        a.Start(); b.Start(); a.Join(); b.Join();
                    }

                    Assert.AreEqual("v1", results[0]?.FolderName, "round " + round + ":\n" + f.Log.Dump());
                    Assert.AreEqual("v1", results[1]?.FolderName, "round " + round + ":\n" + f.Log.Dump());
                    Assert.AreEqual("core", File.ReadAllText(Path.Combine(f.Paths.VersionsPath, "v1", "XLEdge.Addin.Core.dll")));
                    CollectionAssert.AreEqual(new[] { "v1" }, f.ReleaseFolders());
                    Assert.AreEqual(1, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Count);
                }
            }
        }

        [TestMethod]
        public void InstallLock_HeldByAnotherProcess_ResolveAndExtractWaitForIt()
        {
            // Deterministic version of the two-Excel race: whoever holds the install lock, the other waits.
            using (var f = new Fixture())
            {
                StageInManifestFolder(f, CoreZip(), "v1");
                string stray = Path.Combine(f.Paths.VersionsPath, "stray");
                Directory.CreateDirectory(stray);
                File.WriteAllText(Path.Combine(stray, "x.dll"), "x");
                using (var held = new System.Threading.ManualResetEventSlim())
                using (var release = new System.Threading.ManualResetEventSlim())
                {
                    var holder = new System.Threading.Thread(() =>
                    {
                        using (var m = new System.Threading.Mutex(false, UpdateBootstrapper.InstallLockName(f.Paths)))
                        {
                            m.WaitOne();
                            held.Set();
                            release.Wait();
                            m.ReleaseMutex();
                        }
                    });
                    holder.Start();
                    held.Wait();

                    var resolve = System.Threading.Tasks.Task.Run(() => new UpdateBootstrapper().ResolveVersionToLoad(f.Context));
                    var extract = System.Threading.Tasks.Task.Run(() => f.Extract(CoreZip("other"), "v2", version: "11.1.3.0"));
                    Assert.IsFalse(resolve.Wait(500), "ResolveVersionToLoad must wait for the install lock");
                    Assert.IsFalse(extract.Wait(100), "ExtractAndCatalog must wait for the install lock");
                    CollectionAssert.AreEqual(new[] { "stray" }, f.ReleaseFolders(), "nothing may change while another process holds the lock (first-run clean-up included)");

                    release.Set();
                    holder.Join();
                    Assert.IsTrue(System.Threading.Tasks.Task.WaitAll(new System.Threading.Tasks.Task[] { resolve, extract }, 30000));
                }
                var folders = f.ReleaseFolders();
                CollectionAssert.IsSubsetOf(new[] { "v1", "v2" }, folders, f.Log.Dump());
            }
        }

        [TestMethod]
        public void MoveDirectory_RetriesWhileAFileInsideIsBrieflyOpen()
        {
            // Antivirus opens freshly extracted DLLs; renaming a folder fails while any file in it is open.
            using (var t = new TempFolder())
            {
                string src = t.Combine("src"), dst = t.Combine("dst");
                Directory.CreateDirectory(src);
                File.WriteAllText(Path.Combine(src, "a.dll"), "x");
                var handle = File.Open(Path.Combine(src, "a.dll"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
                var closer = new System.Threading.Thread(() => { System.Threading.Thread.Sleep(300); handle.Dispose(); });
                closer.Start();
                try { Assert.IsTrue(UpdateBootstrapper.MoveDirectoryWithRetry(src, dst, new RecordingLogger())); }
                finally { closer.Join(); }
                Assert.IsTrue(File.Exists(Path.Combine(dst, "a.dll")));
                Assert.IsFalse(Directory.Exists(src));
            }
        }

        [TestMethod]
        public void RememberedRelease_SurvivesRestart_UntilANewReleaseIsStaged()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("old"), "vOld", date: "2026-10-01T10:00:00");
                f.Extract(CoreZip("new"), "vNew", date: "2026-10-15T10:00:00");
                f.Bootstrapper.RememberLoaded(f.Context, "vOld");
                Assert.AreEqual("vOld", new UpdateBootstrapper().ResolveVersionToLoad(f.Context)?.FolderName, "rollback must stick:\n" + f.Log.Dump());

                StageInManifestFolder(f, CoreZip("newest"), "vNewest", version: "11.1.4.0", date: "2026-10-20T10:00:00");
                Assert.AreEqual("vNewest", new UpdateBootstrapper().ResolveVersionToLoad(f.Context)?.FolderName, "a staged update overrides it");

                Directory.Delete(Path.Combine(f.Paths.VersionsPath, "vOld"), recursive: true);
                Assert.AreEqual("vNewest", new UpdateBootstrapper().ResolveVersionToLoad(f.Context)?.FolderName, "a remembered release that is gone is ignored");
            }
        }

        [TestMethod]
        public void Manifest_FirstEntryIsUsedEverywhere()
        {
            using (var f = new Fixture())
            {
                byte[] first = CoreZip("first"), second = CoreZip("second");
                Directory.CreateDirectory(f.Paths.ManifestDirectory);
                File.WriteAllText(f.Paths.ManifestFile,
                    "[{\"version\":\"11.1.2.0\",\"releaseDate\":\"" + Date + "\",\"folderName\":\"vFirst\",\"fileName\":\"first.zip\",\"checksum\":\"" + ReleaseFolders.Sha256(first) + "\",\"contractVersion\":1}," +
                    " {\"version\":\"11.1.9.0\",\"releaseDate\":\"" + Date + "\",\"folderName\":\"vSecond\",\"fileName\":\"second.zip\",\"checksum\":\"" + ReleaseFolders.Sha256(second) + "\",\"contractVersion\":1}]");
                File.WriteAllBytes(Path.Combine(f.Paths.ManifestDirectory, "first.zip"), first);
                File.WriteAllBytes(Path.Combine(f.Paths.ManifestDirectory, "second.zip"), second);
                Assert.AreEqual("vFirst", f.Bootstrapper.ResolveVersionToLoad(f.Context)?.FolderName, f.Log.Dump());
            }
        }

        [TestMethod]
        public void Extract_PartialMarkerInFolderName_Refused()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(f.Extract(CoreZip(), "v1.partial-1"));
                Assert.AreEqual(0, f.ReleaseFolders().Length);
            }
        }

        [TestMethod]
        public void Extract_PopulatedFolder_IsCheckedAgainstItsSnapshot()
        {
            using (var f = new Fixture())
            {
                f.Extract(CoreZip("one"), "v1");
                Assert.IsNull(f.Extract(CoreZip("different build"), "v1"), "a different build must not be catalogued over an existing folder");
                Assert.AreEqual(1, ReleaseHistoryStore.ReadAll(f.Paths.ReleaseHistoryFile).Count);

                File.Delete(Path.Combine(f.Paths.VersionsPath, "v1", "manifest.json"));
                Assert.IsNotNull(f.Extract(CoreZip("one"), "v1"));
                Assert.IsTrue(File.Exists(Path.Combine(f.Paths.VersionsPath, "v1", "manifest.json")), "missing snapshot written");
            }
        }

        [TestMethod]
        public void Extract_ReinstallOfTheLoadedRelease_IsAcceptedWithoutError()
        {
            using (var f = new Fixture())
            {
                byte[] zip = CoreZip();
                f.Extract(zip, "v1");
                f.Context.ActiveFolderName = "v1";
                Assert.AreEqual("v1", f.Extract(zip, "v1")?.FolderName, f.Log.Dump());
                Assert.IsFalse(f.Log.Lines.Any(l => l.StartsWith("ERROR")), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Extract_WritesNoTempFilesAndUsesMethodLogPrefixes()
        {
            using (var f = new Fixture())
            {
                string temp = Path.GetTempPath();
                int before = Directory.GetFiles(temp, "XLEdgeRelease_*").Length;
                f.Extract(CoreZip(), "v1");
                Assert.AreEqual(before, Directory.GetFiles(temp, "XLEdgeRelease_*").Length);
                Assert.IsTrue(f.Log.Contains("UpdateBootstrapper.ExtractAndCatalog: extracted and catalogued 'v1'"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void SharedDependencies_WebView2FilesStayTogether()
        {
            using (var f = new Fixture())
            {
                byte[] a = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", "a"), ("Microsoft.Web.WebView2.Core.dll", "core-1"), (@"x64\WebView2Loader.dll", "loader"));
                byte[] b = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", "b"), ("Microsoft.Web.WebView2.Core.dll", "core-2"), (@"x64\WebView2Loader.dll", "loader"));
                f.Extract(a, "vA");
                f.Extract(b, "vB");
                string vB = Path.Combine(f.Paths.VersionsPath, "vB");
                Assert.AreEqual("core-2", File.ReadAllText(Path.Combine(vB, "Microsoft.Web.WebView2.Core.dll")));
                Assert.IsTrue(File.Exists(Path.Combine(vB, "x64", "WebView2Loader.dll")),
                    "a release that keeps its own WebView2.Core must keep the native loader next to it");
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
                Assert.AreEqual(0, Directory.GetFiles(shared, "*.tmp", SearchOption.AllDirectories).Length, "shared copies are written atomically");
            }
        }
    }
}
