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
                var resolveRequests = new List<string>();
                AssemblyLoadEventHandler onLoad = (s, e) => { lock (loadedHere) loadedHere.Add(e.LoadedAssembly.GetName().Name); };
                ResolveEventHandler onResolve = (s, e) => { lock (resolveRequests) resolveRequests.Add(e.Name); return null; };
                AppDomain.CurrentDomain.AssemblyLoad += onLoad;
                AppDomain.CurrentDomain.AssemblyResolve += onResolve;
                try
                {
                    var addin = f.Loader.Load(f.Context, f.Stage("v1"));
                    // The shell holds a proxy of the loader's own type, never of a core type.
                    Assert.AreEqual(typeof(RemoteLoader.RemoteAddinProxy), addin.GetType());
                    Assert.IsNotNull(addin.GetLoginInfo());
                    Assert.IsFalse(addin.IsBusy());
                    Assert.AreEqual(0, addin.GetOpenWindowTitles().Length);
                }
                finally
                {
                    AppDomain.CurrentDomain.AssemblyLoad -= onLoad;
                    AppDomain.CurrentDomain.AssemblyResolve -= onResolve;
                }

                Assert.AreEqual("XLEdgeCore_v1", f.Loader.CurrentDomainName);
                Assert.IsTrue(f.Log.Contains("XLEdge core initialized in AppDomain 'XLEdgeCore_v1' from '" + Path.Combine(f.Paths.VersionsPath, "v1")), f.Log.Dump());
                CollectionAssert.DoesNotContain(loadedHere, "XLEdge.Addin.Core", "the shell domain must never load the core assembly");
                Assert.IsFalse(resolveRequests.Any(r => r.StartsWith("XLEdge.Addin.Core", StringComparison.OrdinalIgnoreCase)),
                    "the shell domain must never even ask for the core assembly: " + string.Join("; ", resolveRequests));
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
                Assert.IsFalse(f.Log.Contains("created AppDomain"), f.Log.Dump());
            }
        }

        [TestMethod]
        public void Load_UnsafeFolderName_RefusedAndLogged()
        {
            using (var f = new Fixture())
            {
                var release = new ResolvedRelease { Version = "1.0", FolderName = "..", ContractVersion = 1 };
                Assert.ThrowsExactly<InvalidOperationException>(() => f.Loader.Load(f.Context, release));
                Assert.IsTrue(f.Log.Contains("AddinDomainLoader.Load: invalid release folder name '..'"), f.Log.Dump());
            }
        }

        /// <summary>Created in the core's domain: a thread there that sits in native code (an unload cannot abort it meanwhile).</summary>
        public sealed class StuckThread : MarshalByRefObject
        {
            [System.Runtime.InteropServices.DllImport("kernel32.dll")]
            private static extern void Sleep(uint milliseconds);

            public void Start(int milliseconds)
            {
                var thread = new System.Threading.Thread(() => Sleep((uint)milliseconds)) { IsBackground = true };
                thread.Start();
                System.Threading.Thread.Sleep(200);
            }
        }

        [TestMethod]
        public void Unload_Normal_ReturnsTrue()
        {
            using (var f = new Fixture())
            {
                var release = f.Stage("v1");
                f.Loader.Load(f.Context, release);
                Assert.IsTrue(f.Loader.Unload(f.Context));
                Assert.IsTrue(f.Loader.Unload(f.Context), "nothing loaded: nothing left to unload");
            }
        }

        [TestMethod]
        public void Unload_DomainThreadStuckInNativeCode_ReturnsFalseWithinTheCap()
        {
            using (var f = new Fixture())
            {
                var release = f.Stage("v1");
                f.Loader.Load(f.Context, release);
                var domain = (AppDomain)typeof(AddinDomainLoader).GetField("_domain", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(f.Loader);
                var stuck = (StuckThread)domain.CreateInstanceFromAndUnwrap(typeof(StuckThread).Assembly.Location, typeof(StuckThread).FullName);
                stuck.Start(9000);

                var watch = System.Diagnostics.Stopwatch.StartNew();
                Assert.IsFalse(f.Loader.Unload(f.Context));
                Assert.IsTrue(watch.Elapsed < TimeSpan.FromSeconds(6.5), "returned at the cap: " + watch.Elapsed);
                Assert.IsTrue(f.Log.Contains("did not unload within 5 s"), f.Log.Dump());

                f.Loader.Load(f.Context, release);
                Assert.IsNotNull(f.Loader.CurrentDomainName, "the next core still loads");
            }
        }

        [TestMethod]
        public void Load_CoreStartThrows_LoadThrowsAndDomainIsUnloaded()
        {
            // A core whose start fails must not count as loaded (every click would then do nothing, with no message).
            using (var f = new Fixture())
            {
                var release = f.Stage("v1");
                f.Context.RibbonControllerThrows = "no ribbon";
                var ex = Assert.ThrowsExactly<InvalidOperationException>(() => f.Loader.Load(f.Context, release));
                StringAssert.Contains(ex.Message, "no ribbon");
                Assert.IsNull(f.Loader.CurrentDomainName);
                Assert.IsTrue(f.Log.Contains("AddinDomainLoader.Load: failed to load"), f.Log.Dump());
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

        // ---- resolver (run in a throwaway AppDomain: it loads files and must not pollute the test domain) ----

        private static void InProbeDomain(Action<ResolverProbe> test)
        {
            var domain = AppDomain.CreateDomain("ResolverProbe", null, new AppDomainSetup { ApplicationBase = Path.GetDirectoryName(typeof(ResolverProbe).Assembly.Location) });
            try
            {
                test((ResolverProbe)domain.CreateInstanceAndUnwrap(typeof(ResolverProbe).Assembly.FullName, typeof(ResolverProbe).FullName));
            }
            finally
            {
                AppDomain.Unload(domain);
            }
        }

        private static string PackageDll(string packageId, string dll)
        {
            string testBin = Path.GetDirectoryName(typeof(AddinDomainLoaderTests).Assembly.Location);
            string packages = Path.GetFullPath(Path.Combine(testBin, @"..\..\..\packages"));
            string package = Directory.GetDirectories(packages, packageId + ".*").OrderByDescending(d => d).First();
            return Path.Combine(package, @"lib\net462", dll);
        }

        [TestMethod]
        public void Resolver_ProbesReleaseFirstThenShared_AndTriesTheNextCandidateWhenOneFails()
        {
            const string name = "System.IO.Pipelines";
            using (var t = new TempFolder())
            {
                string release = t.Combine("release"), shared = t.Combine("shared");
                Directory.CreateDirectory(release);
                Directory.CreateDirectory(shared);
                File.WriteAllText(Path.Combine(release, name + ".dll"), "corrupt");
                File.Copy(PackageDll(name, name + ".dll"), Path.Combine(shared, name + ".dll"));
                var log = new RecordingLogger();

                InProbeDomain(probe =>
                {
                    Assert.IsFalse(probe.IsLoaded(name), "precondition: not loaded in the probe domain");
                    string location = probe.Resolve(name + ", Culture=neutral", release, shared, log);
                    Assert.IsNotNull(location, log.Dump());
                    StringAssert.StartsWith(location, shared);
                });
                Assert.IsTrue(log.Contains("RemoteLoader.ResolveFromFolders: could not load"), log.Dump());
            }
        }

        [TestMethod]
        public void Resolver_SkipsResourcesAndSerializersAndReturnsAlreadyLoaded()
        {
            var log = new RecordingLogger();
            InProbeDomain(probe =>
            {
                Assert.IsNull(probe.Resolve("XLEdge.Addin.Core.resources, Culture=en-US", @"C:\none", @"C:\none", log));
                Assert.IsNull(probe.Resolve("XLEdge.Addin.Core.XmlSerializers, Culture=neutral", @"C:\none", @"C:\none", log));
                Assert.IsNotNull(probe.Resolve(typeof(ContractInfo).Assembly.FullName, @"C:\none", @"C:\none", log));
                Assert.IsNull(probe.Resolve("Does.Not.Exist", @"C:\none", null, log));
            });
            Assert.IsTrue(log.Contains("RemoteLoader.ResolveFromFolders: could not resolve 'Does.Not.Exist'"), log.Dump());
            Assert.IsFalse(log.Contains("XmlSerializers"), "serializer probes are expected and not worth a warning: " + log.Dump());
        }
    }
}
