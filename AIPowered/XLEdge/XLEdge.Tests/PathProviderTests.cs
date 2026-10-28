using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class PathProviderTests
    {
        [TestMethod]
        public void Paths_FollowXLEdgeLayout()
        {
            using (var t = new TempFolder())
            {
                string inst = t.Combine("inst"), data = t.Combine("data");
                var p = new PathProvider(inst, data);

                Assert.AreEqual(Path.Combine(data, "XLEdge_Logs"), p.Root);
                Assert.AreEqual(Path.Combine(data, "XLEdge_Logs", "Logs"), p.Logs);
                Assert.AreEqual(Path.Combine(data, "XLEdge_Logs", "Temp"), p.Temp);
                Assert.AreEqual(Path.Combine(data, "XLEdge_Logs", "BrowserLogs"), p.BrowserDataPath);
                Assert.AreEqual(Path.Combine(data, "ORBIT_URLS.xml"), p.UrlsFile);
                Assert.AreEqual(Path.Combine(inst, "AddinCore", "Versions"), p.VersionsPath);
                Assert.AreEqual(Path.Combine(inst, "AddinCore", "SharedDependencies"), p.SharedDependenciesPath);
                Assert.AreEqual(Path.Combine(inst, "AddinCore", "Manifest"), p.ManifestDirectory);
                Assert.AreEqual(Path.Combine(inst, "AddinCore", "Manifest", "manifest.json"), p.ManifestFile);
                Assert.AreEqual(Path.Combine(inst, "AddinCore", "ReleaseHistory.json"), p.ReleaseHistoryFile);
            }
        }

        [TestMethod]
        public void DefaultDataBase_IsLocalAppDataOrbitExcelLogs()
        {
            var p = new PathProvider(@"C:\inst");
            var expected = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ORBIT", "Excel_Logs");
            Assert.AreEqual(Path.Combine(expected, "XLEdge_Logs"), p.Root);
            Assert.AreEqual(Path.Combine(expected, "ORBIT_URLS.xml"), p.UrlsFile);
        }

        [TestMethod]
        public void Ctor_DoesNoIo()
        {
            using (var t = new TempFolder())
            {
                new PathProvider(t.Combine("inst"), t.Combine("data"));
                Assert.IsFalse(Directory.Exists(t.Combine("inst")));
                Assert.IsFalse(Directory.Exists(t.Combine("data")));
            }
        }

        [TestMethod]
        public void Ensure_CreatesFolders()
        {
            using (var t = new TempFolder())
            {
                var p = new PathProvider(t.Combine("inst"), t.Combine("data"));
                p.Ensure();
                foreach (var dir in new[] { p.Logs, p.Temp, p.BrowserDataPath, p.VersionsPath, p.SharedDependenciesPath, p.ManifestDirectory })
                    Assert.IsTrue(Directory.Exists(dir), dir);
                Assert.IsFalse(File.Exists(p.ManifestFile), "Ensure must not create manifest.json");
            }
        }

        [TestMethod]
        public void Refresh_PicksHighestVersion()
        {
            using (var t = new TempFolder())
            {
                var p = new PathProvider(t.Combine("inst"), t.Combine("data"));
                p.Ensure();
                File.WriteAllText(p.ManifestFile,
                    "[{\"version\":\"11.1.1\",\"releaseDate\":\"a\"},{\"version\":\"11.1.10\",\"releaseDate\":\"b\"},{\"version\":\"bad\",\"releaseDate\":\"c\"}]");
                p.Refresh();
                Assert.AreEqual("11.1.10", p.LatestVersion);
                Assert.AreEqual("b", p.LatestReleaseDate);
            }
        }

        [TestMethod]
        public void Refresh_MissingManifest_NullsAndDoesNotCreate()
        {
            using (var t = new TempFolder())
            {
                var p = new PathProvider(t.Combine("inst"), t.Combine("data"));
                p.Ensure();
                p.Refresh();
                Assert.IsNull(p.LatestVersion);
                Assert.IsNull(p.LatestReleaseDate);
                Assert.IsFalse(File.Exists(p.ManifestFile));
            }
        }

        [TestMethod]
        public void Refresh_CorruptOrObjectOrEmpty_Nulls()
        {
            using (var t = new TempFolder())
            {
                var p = new PathProvider(t.Combine("inst"), t.Combine("data"));
                p.Ensure();
                foreach (var content in new[] { "{not json", "{\"version\":\"11.1.2\"}", "", "[]" })
                {
                    File.WriteAllText(p.ManifestFile, "[{\"version\":\"11.1.1\",\"releaseDate\":\"a\"}]");
                    p.Refresh();
                    Assert.AreEqual("11.1.1", p.LatestVersion);

                    File.WriteAllText(p.ManifestFile, content);
                    p.Refresh();
                    Assert.IsNull(p.LatestVersion, "content: " + content);
                    Assert.IsNull(p.LatestReleaseDate, "content: " + content);
                }
            }
        }

        [TestMethod]
        public void Ctor_BlankInstallRoot_Throws()
        {
            Assert.ThrowsExactly<ArgumentException>(() => new PathProvider(null));
            Assert.ThrowsExactly<ArgumentException>(() => new PathProvider("  "));
        }

        [TestMethod]
        public void Lease_IsInfinite()
        {
            Assert.IsInstanceOfType(new PathProvider(@"C:\inst"), typeof(MarshalByRefObject));
            Assert.IsNull(new PathProvider(@"C:\inst").InitializeLifetimeService());
        }
    }
}
