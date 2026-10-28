using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>The core build publishes a release (zip + manifest.json) that the bootstrapper accepts and the loader runs.</summary>
    [TestClass]
    public class CorePackageTests
    {
        internal static string PackageFolder
        {
            get
            {
                string coreBin = ReleaseFolders.CoreBuildFolder;
                string config = new DirectoryInfo(coreBin).Name;
                return Path.GetFullPath(Path.Combine(coreBin, @"..\..\SetupFiles", config, "Manifest"));
            }
        }

        [TestMethod]
        public void BuiltPackage_IsAValidRelease()
        {
            string folder = PackageFolder;
            Assert.IsTrue(Directory.Exists(folder), "no package folder " + folder);
            var zips = Directory.GetFiles(folder, "*.zip");
            Assert.AreEqual(1, zips.Length, "exactly one zip");
            string manifestPath = Path.Combine(folder, "manifest.json");
            Assert.IsTrue(File.Exists(manifestPath));

            var parsed = new VersionParser().ParseVersionFile(manifestPath);
            Assert.IsTrue(parsed.Success, parsed.ErrorMessage);
            VersionInfo release = parsed.AllVersions.Single();
            byte[] zip = File.ReadAllBytes(zips[0]);
            Assert.AreEqual(ReleaseFolders.Sha256(zip), release.Checksum);
            Assert.AreEqual(ContractInfo.Version, release.ContractVersion);
            Assert.IsTrue(ReleaseHistoryStore.IsSafeFolderName(release.FolderName), release.FolderName);
            Assert.AreEqual(release.FolderName + ".zip", release.FileName);
            Assert.AreEqual(Path.GetFileName(zips[0]), release.FileName);
            Assert.IsFalse(File.ReadAllBytes(manifestPath).Take(3).SequenceEqual(new byte[] { 0xEF, 0xBB, 0xBF }), "manifest.json without BOM");

            using (var archive = new ZipArchive(new MemoryStream(zip)))
            {
                var names = archive.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.IsSubsetOf(new[] { "XLEdge.Addin.Core.dll", "XLEdge.Contracts.dll", "XLEdge.Shared.dll" }, names);
                Assert.IsFalse(names.Any(n => n.EndsWith(".pdb") || n.EndsWith(".xml") || n.EndsWith(".zip")), string.Join(", ", names));
                CollectionAssert.DoesNotContain(names, "XLEdge.Loader.Core.dll", "the shell supplies the loader");
            }
        }

        [TestMethod]
        public void BuiltPackage_CarriesWebView2()
        {
            // The pane content is a WinForms WebView2 in the core; WebView2 finds its native loader under runtimes\ next to Core.dll.
            using (var archive = new ZipArchive(File.OpenRead(Directory.GetFiles(PackageFolder, "*.zip").Single())))
            {
                var names = archive.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.IsSubsetOf(new[]
                {
                    "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.WinForms.dll",
                    "runtimes/win-x86/native/WebView2Loader.dll", "runtimes/win-x64/native/WebView2Loader.dll",
                }, names, string.Join(", ", names));
                Assert.AreEqual(new System.Version(1, 0, 4022, 49), System.Reflection.AssemblyName.GetAssemblyName(
                    Path.Combine(ReleaseFolders.CoreBuildFolder, "Microsoft.Web.WebView2.Core.dll")).Version, "today's XLEdge WebView2 version");
            }
        }

        [TestMethod]
        public void ShellBuild_StagesExactlyTheCorePackage_AndNoCoreNextToTheShell()
        {
            string coreBin = ReleaseFolders.CoreBuildFolder;
            string config = new DirectoryInfo(coreBin).Name;
            string shellBin = Path.GetFullPath(Path.Combine(coreBin, @"..\..\..\XLEdge\bin", config));
            string staged = Path.Combine(shellBin, "AddinCore", "Manifest");
            var package = Directory.GetFiles(PackageFolder).Select(Path.GetFileName).OrderBy(n => n).ToArray();
            CollectionAssert.AreEqual(package, Directory.GetFiles(staged).Select(Path.GetFileName).OrderBy(n => n).ToArray());
            foreach (string name in package)
                CollectionAssert.AreEqual(File.ReadAllBytes(Path.Combine(PackageFolder, name)), File.ReadAllBytes(Path.Combine(staged, name)), name);
            Assert.IsFalse(File.Exists(Path.Combine(shellBin, "XLEdge.Addin.Core.dll")));
            foreach (string loader in new[] { "adxloader.dll", "adxloader64.dll", "adxloader.dll.manifest" })
                Assert.IsTrue(File.Exists(Path.Combine(shellBin, loader)), loader);
        }

        [TestMethod]
        public void BuiltPackage_RunsWithTheShellsRealLoggerAndPaths()
        {
            // The shell passes XLEdge.Shared's Logger/PathProvider (NLog, System.Text.Json) into the core's AppDomain.
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                var context = new FakeContext(paths) { DebugMode = true };
                using (var logger = new Logger(context))
                {
                    context.Logger = logger;
                    string folder = PackageFolder;
                    var release = new UpdateBootstrapper().ExtractAndCatalog(context,
                        File.ReadAllText(Path.Combine(folder, "manifest.json")), File.ReadAllBytes(Directory.GetFiles(folder, "*.zip").Single()), "Install");
                    Assert.IsNotNull(release);
                    var loader = new AddinDomainLoader();
                    try
                    {
                        loader.Load(context, release).OnRibbonAction("RibEdgeRefresh", null);
                    }
                    finally
                    {
                        loader.Unload(context);
                    }
                }
                string log = File.ReadAllText(Directory.GetFiles(paths.Logs, "XLEdge_Logs_*.log").Single());
                StringAssert.Contains(log, "AddinEntry.Initialize: XLEdge core initialized in AppDomain 'XLEdgeCore_");
                StringAssert.Contains(log, "AddinEntry.OnRibbonAction: 'RibEdgeRefresh'.");
            }
        }

        [TestMethod]
        public void BuiltPackage_InstallsAndLoads()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                var context = new FakeContext(paths) { Logger = new RecordingLogger() };
                string folder = PackageFolder;
                var release = new UpdateBootstrapper().ExtractAndCatalog(context,
                    File.ReadAllText(Path.Combine(folder, "manifest.json")), File.ReadAllBytes(Directory.GetFiles(folder, "*.zip").Single()), "Install");
                Assert.IsNotNull(release, ((RecordingLogger)context.Logger).Dump());

                var loader = new AddinDomainLoader();
                try
                {
                    var addin = loader.Load(context, release);
                    Assert.IsFalse(addin.IsBusy());
                }
                finally
                {
                    loader.Unload(context);
                }
            }
        }
    }
}
