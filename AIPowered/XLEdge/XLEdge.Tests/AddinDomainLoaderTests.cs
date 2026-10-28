using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class AddinDomainLoaderTests
    {
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
            public AddinDomainLoader Loader { get; } = new AddinDomainLoader();

            public ResolvedRelease Stage(string folder, int contractVersion = 1)
            {
                ReleaseFolders.Stage(Paths.VersionsPath, folder);
                return new ResolvedRelease { Version = "11.1.2.0", ReleaseDate = "2026-10-28T10:00:00", FolderName = folder, ContractVersion = contractVersion };
            }

            public void Dispose()
            {
                Loader.Unload(Context);
                _temp.Dispose();
            }
        }

        [TestMethod]
        public void Load_RunsCoreInItsOwnDomain_ShellNeverLoadsTheCoreAssembly()
        {
            using (var f = new Fixture())
            {
                var loadedHere = new List<string>();
                AssemblyLoadEventHandler onLoad = (s, e) => { lock (loadedHere) loadedHere.Add(e.LoadedAssembly.GetName().Name); };
                AppDomain.CurrentDomain.AssemblyLoad += onLoad;
                try
                {
                    var addin = f.Loader.Load(f.Context, f.Stage("v1"));
                    Assert.IsNotNull(addin.GetLoginInfo());
                    Assert.IsFalse(addin.IsBusy());
                    Assert.AreEqual(0, addin.GetOpenWindowTitles().Length);
                }
                finally
                {
                    AppDomain.CurrentDomain.AssemblyLoad -= onLoad;
                }

                Assert.AreEqual("XLEdgeCore_v1", f.Loader.CurrentDomainName);
                Assert.IsTrue(f.Log.Contains("XLEdge core initialized in AppDomain 'XLEdgeCore_v1' from '" + Path.Combine(f.Paths.VersionsPath, "v1")), f.Log.Dump());
                CollectionAssert.DoesNotContain(loadedHere, "XLEdge.Addin.Core", "the shell domain must never load the core assembly");
            }
        }

        [TestMethod]
        public void Load_NewerOrUnknownContract_RefusedWithoutCreatingDomain()
        {
            using (var f = new Fixture())
            {
                foreach (int contract in new[] { 0, ContractInfo.Version + 1 })
                {
                    var ex = Assert.ThrowsExactly<InvalidOperationException>(() => f.Loader.Load(f.Context, f.Stage("v" + contract, contract)));
                    StringAssert.Contains(ex.Message, ContractCompatibility.RefusalReason(contract));
                    Assert.IsNull(f.Loader.CurrentDomainName);
                }
                Assert.IsFalse(f.Log.Contains("created AppDomain"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Load_MissingCoreDll_ThrowsAndLeavesNothingLoaded()
        {
            using (var f = new Fixture())
            {
                Directory.CreateDirectory(Path.Combine(f.Paths.VersionsPath, "empty"));
                var release = new ResolvedRelease { Version = "1.0", FolderName = "empty", ContractVersion = 1 };
                Assert.ThrowsExactly<FileNotFoundException>(() => f.Loader.Load(f.Context, release));
                Assert.IsNull(f.Loader.CurrentDomainName);
            }
        }

        [TestMethod]
        public void Load_CoreFailsInsideItsDomain_UnloadsTheHalfCreatedDomain()
        {
            using (var f = new Fixture())
            {
                var release = f.Stage("broken");
                File.WriteAllText(Path.Combine(f.Paths.VersionsPath, "broken", "XLEdge.Addin.Core.dll"), "not an assembly");
                Assert.ThrowsExactly<BadImageFormatException>(() => f.Loader.Load(f.Context, release));
                Assert.IsNull(f.Loader.CurrentDomainName);
                Assert.IsTrue(f.Log.Contains("AppDomain unloaded"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Unload_ReleasesTheReleaseFolder()
        {
            using (var f = new Fixture())
            {
                var addin = f.Loader.Load(f.Context, f.Stage("v1"));
                addin.OnRibbonAction("RibEdgeRefresh", null);
                addin.Shutdown();
                f.Loader.Unload(f.Context);
                Assert.IsNull(f.Loader.CurrentDomainName);
                Directory.Delete(Path.Combine(f.Paths.VersionsPath, "v1"), recursive: true);
            }
        }

        [TestMethod]
        public void Reload_AThenBThenA()
        {
            using (var f = new Fixture())
            {
                var a = f.Stage("vA");
                var b = f.Stage("vB");
                foreach (var release in new[] { a, b, a })
                {
                    var addin = f.Loader.Load(f.Context, release);
                    Assert.AreEqual("XLEdgeCore_" + release.FolderName, f.Loader.CurrentDomainName);
                    Assert.IsFalse(addin.IsBusy());
                    addin.Shutdown();
                    f.Loader.Unload(f.Context);
                }
                Assert.AreEqual(2, f.Log.Lines.Count(l => l.Contains("initialized in AppDomain 'XLEdgeCore_vA'")), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Unload_WhenNothingLoaded_IsSafe()
        {
            var loader = new AddinDomainLoader();
            loader.Unload(null);
            using (var f = new Fixture()) f.Loader.Unload(f.Context);
        }

        [TestMethod]
        public void Proxy_HasInfiniteLease()
        {
            using (var f = new Fixture())
            {
                var addin = f.Loader.Load(f.Context, f.Stage("v1"));
                Assert.IsTrue(RemotingServices.IsTransparentProxy(addin));
                Assert.IsNull(RemotingServices.GetLifetimeService((MarshalByRefObject)addin));
            }
        }

        // ---- resolver --------------------------------------------------------------

        private static string PackageDll(string package, string dll)
        {
            string testBin = Path.GetDirectoryName(typeof(AddinDomainLoaderTests).Assembly.Location);
            return Path.GetFullPath(Path.Combine(testBin, @"..\..\..\packages", package, @"lib\net462", dll));
        }

        [TestMethod]
        public void Resolver_ProbesReleaseFirstThenShared_AndTriesTheNextCandidateWhenOneFails()
        {
            const string name = "System.IO.Pipelines";
            Assert.IsFalse(AppDomain.CurrentDomain.GetAssemblies().Any(a => a.GetName().Name == name), "precondition: not loaded yet");
            using (var t = new TempFolder())
            {
                string release = t.Combine("release"), shared = t.Combine("shared");
                Directory.CreateDirectory(release);
                Directory.CreateDirectory(shared);
                File.WriteAllText(Path.Combine(release, name + ".dll"), "corrupt");
                File.Copy(PackageDll("System.IO.Pipelines.10.0.9", name + ".dll"), Path.Combine(shared, name + ".dll"));
                var log = new RecordingLogger();

                Assembly resolved = RemoteLoader.ResolveFromFolders(name + ", Version=10.0.0.9, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51", release, shared, log);

                Assert.IsNotNull(resolved, log.Dump());
                Assert.AreEqual(name, resolved.GetName().Name);
                StringAssert.StartsWith(resolved.Location, shared);
                Assert.IsTrue(log.Contains("AssemblyResolve: could not load"), log.Dump());
            }
        }

        [TestMethod]
        public void Resolver_SkipsResourcesAndReturnsAlreadyLoaded()
        {
            var log = new RecordingLogger();
            Assert.IsNull(RemoteLoader.ResolveFromFolders("XLEdge.Addin.Core.resources, Culture=en-US", @"C:\none", @"C:\none", log));
            var contracts = typeof(ContractInfo).Assembly;
            Assert.AreSame(contracts, RemoteLoader.ResolveFromFolders(contracts.FullName, @"C:\none", @"C:\none", log));
            Assert.IsNull(RemoteLoader.ResolveFromFolders("Does.Not.Exist", @"C:\none", null, log));
            Assert.IsTrue(log.Contains("AssemblyResolve: could not resolve 'Does.Not.Exist'"), log.Dump());
        }
    }
}
