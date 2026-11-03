using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Hosting;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class CoreHostTests
    {
        // ---- fakes -------------------------------------------------------------------------

        private sealed class Recorder
        {
            public readonly List<string> Events = new List<string>();
            public void Add(string e) { lock (Events) Events.Add(e); }
        }

        private sealed class FakeAddin : IXLEdgeAddin
        {
            private readonly Recorder _r;
            public FakeAddin(Recorder r, string name) { _r = r; Name = name; }
            public string Name { get; }
            public string[] Titles = new string[0];
            public bool ThrowOnTitles, ThrowOnShutdown, Busy;
            public void Initialize(IXLEdgeContext context) { }
            public void OnRibbonAction(string action, object parameter) { }
            public bool OnExcelEvent(string eventName, object[] args) => true;
            public void Shutdown() { _r.Add("Shutdown:" + Name); if (ThrowOnShutdown) throw new InvalidOperationException("shutdown failed"); }
            public IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle) => IntPtr.Zero;
            public void CloseTaskPaneContent(IntPtr excelWindowHandle) { }
            public void OnTaskPaneShown(IntPtr excelWindowHandle) { }
            public void OnLoginFromGLSense(string i, string u, string t, string n, bool p) { }
            public void OnLogoutFromGLSense() { }
            public LoginInfo GetLoginInfo() => new LoginInfo();
            public string[] GetOpenWindowTitles() { if (ThrowOnTitles) throw new InvalidOperationException("old core"); return Titles; }
            public bool IsBusy() => Busy;
        }

        private sealed class FakeLoader : ICoreLoader
        {
            private readonly Recorder _r;
            public FakeLoader(Recorder r) { _r = r; }
            public Func<ResolvedRelease, IXLEdgeAddin> OnLoad;
            public bool ThrowOnUnload;
            public bool UnloadResult = true;
            public IXLEdgeAddin Load(IXLEdgeContext context, ResolvedRelease release) { _r.Add("Load:" + release.FolderName); return OnLoad(release); }
            public bool Unload(IXLEdgeContext context) { _r.Add("Unload"); if (ThrowOnUnload) throw new InvalidOperationException("unload failed"); return UnloadResult; }
        }

        private sealed class FakeCatalog : IReleaseCatalog
        {
            private readonly Recorder _r;
            public FakeCatalog(Recorder r) { _r = r; }
            public ResolvedRelease StartRelease;
            public ResolvedRelease ResolveVersionToLoad(IXLEdgeContext context, string source) { _r.Add("Resolve:" + source); return StartRelease; }
            public void RememberLoaded(IXLEdgeContext context, string folderName) => _r.Add("Remember:" + folderName);
        }

        private sealed class FakeUi : IHostUi
        {
            private readonly Recorder _r;
            public FakeUi(Recorder r) { _r = r; }
            public bool ConfirmAnswer = true;
            public Action<int> OnStage;
            public readonly List<string> Errors = new List<string>();
            public bool ConfirmCloseOpenWindows(string[] titles) { _r.Add("Confirm:" + string.Join("|", titles)); return ConfirmAnswer; }
            public IReloadProgress ShowProgress(string currentVersion, string currentReleaseDate) { _r.Add("ProgressShown:" + currentVersion); return new Progress(_r, this); }
            public void ShowError(string message, string title) { _r.Add("Error"); Errors.Add(title + ": " + message); }

            private sealed class Progress : IReloadProgress
            {
                private readonly Recorder _r;
                private readonly FakeUi _ui;
                public Progress(Recorder r, FakeUi ui) { _r = r; _ui = ui; }
                public void ShowStage(string message, int step) { _r.Add($"Stage{step}:{message}"); _ui.OnStage?.Invoke(step); }
                public void SetTarget(string version, string releaseDate) => _r.Add("Target:" + version);
                public void Close() => _r.Add("ProgressClosed");
            }
        }

        private sealed class FakePanes : IPaneHost
        {
            private readonly Recorder _r;
            public FakePanes(Recorder r) { _r = r; }
            public void DetachAll(IXLEdgeAddin core) => _r.Add("Detach:" + ((FakeAddin)core).Name);
            public void AttachAll() => _r.Add("Attach");
        }

        private sealed class Setup
        {
            public readonly Recorder R = new Recorder();
            public readonly RecordingLogger Log = new RecordingLogger();
            public readonly FakeContext Context;
            public readonly FakeLoader Loader;
            public readonly FakeCatalog Catalog;
            public readonly FakeUi Ui;
            public readonly CoreHost Host;
            public readonly Dictionary<string, FakeAddin> Addins = new Dictionary<string, FakeAddin>();

            public Setup()
            {
                Context = new FakeContext(new PathProvider(@"C:\inst", @"C:\data")) { Logger = Log };
                Loader = new FakeLoader(R) { OnLoad = rel => Addins[rel.FolderName] = new FakeAddin(R, rel.FolderName) };
                Catalog = new FakeCatalog(R) { StartRelease = Release("vA") };
                Ui = new FakeUi(R);
                Host = new CoreHost(Context, Loader, Catalog, Ui, new FakePanes(R));
            }

            public static ResolvedRelease Release(string folder) => new ResolvedRelease { Version = "ver-" + folder, ReleaseDate = "date-" + folder, FolderName = folder, ContractVersion = 1 };
        }

        // ---- Start -------------------------------------------------------------------------

        [TestMethod]
        public void Start_LoadsSetsContextAndRemembers()
        {
            var s = new Setup();
            Assert.IsTrue(s.Host.Start());
            Assert.AreSame(s.Addins["vA"], s.Host.Addin);
            Assert.AreEqual("vA", s.Context.ActiveFolderName);
            Assert.AreEqual("ver-vA", s.Context.Version);
            Assert.AreEqual("date-vA", s.Context.ReleaseDate);
            CollectionAssert.AreEqual(new[] { "Resolve:Install", "Load:vA", "Remember:vA", "Attach" }, s.R.Events);
            Assert.IsTrue(s.Log.Contains("SESSION | begin ver-vA"), s.Log.Dump());
            Assert.AreEqual(0, s.Ui.Errors.Count);
        }

        [TestMethod]
        public void StartAndReload_RunInsideALogScope_SoTheirDebugLinesFollowTheMarker()
        {
            var s = new Setup();
            Assert.IsTrue(s.Host.Start());
            CollectionAssert.AreEqual(new[] { "begin:XLEdge start", "end:XLEdge start" }, s.Log.Events.Where(e => e.Contains("XLEdge start")).ToArray());
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vB")));
            CollectionAssert.AreEqual(new[] { "begin:XLEdge reload", "end:XLEdge reload" }, s.Log.Events.Where(e => e.Contains("XLEdge reload")).ToArray());
        }

        [TestMethod]
        public void LoadingACoreThatPredatesTheSessionBag_ClearsIt()
        {
            // Such a core (Release History rollback) neither reads nor keeps session values; whatever is in the bag is stale by
            // the time a newer core loads again (e.g. session-only options would come back over options saved meanwhile).
            var s = new Setup();
            s.Context.SetSessionValue("XLEdge.SessionOptions", "{}");
            s.Catalog.StartRelease = new ResolvedRelease { Version = "v2", ReleaseDate = "d", FolderName = "vOld", ContractVersion = 2 };
            Assert.IsTrue(s.Host.Start());
            Assert.IsNull(s.Context.GetSessionValue("XLEdge.SessionOptions"));
        }

        [TestMethod]
        public void LoadingACoreThatPredatesTheSessionLogin_DropsOnlyTheLogin()
        {
            // A contract-3 core (2c/2d, Release History) shows a logged-out ribbon and does not keep the login: back on a newer
            // core, the old token must not log the user in again silently.
            var s = new Setup();
            s.Context.SetSessionValue("XLEdge.SessionOptions", "{}");
            s.Context.SetSessionValue("XLEdge.Login", "{\"LoginToken\":\"t\"}");
            s.Catalog.StartRelease = new ResolvedRelease { Version = "v3", ReleaseDate = "d", FolderName = "vOld", ContractVersion = 3 };
            Assert.IsTrue(s.Host.Start());
            Assert.IsNull(s.Context.GetSessionValue("XLEdge.Login"));
            Assert.AreEqual("{}", s.Context.GetSessionValue("XLEdge.SessionOptions"), "the options it does keep");
        }

        [TestMethod]
        public void LoadingACoreWithTheSessionLogin_KeepsIt()
        {
            var s = new Setup();
            s.Context.SetSessionValue("XLEdge.Login", "{}");
            s.Catalog.StartRelease = new ResolvedRelease { Version = "v4", ReleaseDate = "d", FolderName = "vNew", ContractVersion = 4 };
            Assert.IsTrue(s.Host.Start());
            Assert.AreEqual("{}", s.Context.GetSessionValue("XLEdge.Login"));
        }

        [TestMethod]
        public void LoadingACoreWithTheSessionBag_KeepsIt()
        {
            var s = new Setup();
            s.Context.SetSessionValue("XLEdge.SessionOptions", "{}");
            s.Catalog.StartRelease = new ResolvedRelease { Version = "v3", ReleaseDate = "d", FolderName = "vNew", ContractVersion = 3 };
            Assert.IsTrue(s.Host.Start());
            Assert.AreEqual("{}", s.Context.GetSessionValue("XLEdge.SessionOptions"));
        }

        [TestMethod]
        public void Start_NothingToLoad_ShowsErrorAndStaysUsable()
        {
            var s = new Setup();
            s.Catalog.StartRelease = null;
            Assert.IsFalse(s.Host.Start());
            Assert.IsNull(s.Host.Addin);
            CollectionAssert.AreEqual(new[] { "XLEdge: XLEdge could not find a usable core version. Use Reload Add-in to load one, or reinstall XLEdge. Check the XLEdge log for details." }, s.Ui.Errors);
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vB")), "Reload must still work");
        }

        [TestMethod]
        public void Start_LoadThrows_ShowsErrorClearsContextAndDoesNotRemember()
        {
            var s = new Setup();
            s.Loader.OnLoad = rel => throw new BadImageFormatException("bad core");
            Assert.IsFalse(s.Host.Start());
            Assert.IsNull(s.Host.Addin);
            Assert.IsNull(s.Context.ActiveFolderName);
            Assert.IsNull(s.Context.Version);
            CollectionAssert.DoesNotContain(s.R.Events, "Remember:vA");
            Assert.AreEqual("XLEdge: XLEdge could not load core version ver-vA: bad core", s.Ui.Errors.Single());
        }

        // ---- Reload ------------------------------------------------------------------------

        [TestMethod]
        public void Reload_SwitchesRelease_InOrder()
        {
            var s = new Setup();
            s.Host.Start();
            s.R.Events.Clear();
            Assert.IsTrue(s.Host.Reload(() => { s.R.Add("Resolve:new"); return Setup.Release("vB"); }));

            CollectionAssert.AreEqual(new[]
            {
                "ProgressShown:ver-vA",
                "Stage1:Closing the current XLEdge session...",
                "Detach:vA", "Shutdown:vA", "Unload",
                "Stage2:Finding the selected XLEdge version...",
                "Resolve:new", "Target:ver-vB",
                "Stage3:Loading the selected XLEdge version...",
                "Load:vB", "Remember:vB", "Attach",
                "Stage4:XLEdge reload complete.",
                "ProgressClosed",
            }, s.R.Events);
            Assert.AreSame(s.Addins["vB"], s.Host.Addin);
            Assert.AreEqual("vB", s.Context.ActiveFolderName);
            var sessions = s.Log.Lines.Where(l => l.StartsWith("SESSION")).ToArray();
            CollectionAssert.AreEqual(new[] { "SESSION | begin ver-vA", "SESSION | end ver-vA", "SESSION | begin ver-vB" }, sessions);
        }

        [TestMethod]
        public void Reload_PreviousCoreStuck_NextCoreLoadsAndTheUserIsTold()
        {
            var s = new Setup();
            s.Host.Start();
            s.Loader.UnloadResult = false;
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vB")), "the next core still loads");
            Assert.AreSame(s.Addins["vB"], s.Host.Addin);
            CollectionAssert.AreEqual(new[] { "Reload XLEdge Add-in: The previous XLEdge core version did not close within 5 seconds. XLEdge keeps working; restart Excel before installing or switching versions again." }, s.Ui.Errors);

            s.Loader.UnloadResult = true;
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vA")));
            Assert.AreEqual(1, s.Ui.Errors.Count, "told once - this unload finished");
        }

        [TestMethod]
        public void Reload_WhileTheCoreIsBusy_IsRefused()
        {
            var s = new Setup();
            s.Host.Start();
            s.Addins["vA"].Busy = true;
            s.R.Events.Clear();
            Assert.IsFalse(s.Host.Reload(() => Setup.Release("vB")));
            CollectionAssert.AreEqual(new[] { "Error" }, s.R.Events);
            Assert.AreEqual("Reload XLEdge Add-in: XLEdge is busy (a report or refresh is running, or an XLEdge window is open). Use Reload Add-in again when it has finished.", s.Ui.Errors.Single());
            Assert.AreSame(s.Addins["vA"], s.Host.Addin);
        }

        [TestMethod]
        public void Reload_NoCoreIsReachableWhileReloading()
        {
            // Ribbon clicks / Excel events / GLSense calls processed while the progress window pumps messages must not reach
            // the core that is being shut down.
            var s = new Setup();
            s.Host.Start();
            var seen = new List<object>();
            s.Ui.OnStage = step => seen.Add(s.Host.Addin);
            s.Host.Reload(() => Setup.Release("vB"));
            CollectionAssert.AreEqual(new object[] { null, null, null, s.Addins["vB"] }, seen);
        }

        [TestMethod]
        public void Shutdown_DuringReload_NoNewCoreIsLoaded()
        {
            var s = new Setup();
            s.Host.Start();
            s.Ui.OnStage = step => { if (step == 2) s.Host.Shutdown(); };
            Assert.IsFalse(s.Host.Reload(() => Setup.Release("vB")));
            CollectionAssert.DoesNotContain(s.R.Events, "Load:vB");
            Assert.IsNull(s.Host.Addin);
            Assert.AreEqual(0, s.Ui.Errors.Count, "Excel is closing - no message");
        }

        [TestMethod]
        public void Reload_ResolveThrows_ReportsItsMessage()
        {
            var s = new Setup();
            s.Host.Start();
            Assert.IsFalse(s.Host.Reload(() => throw new System.IO.IOException("disk full")));
            Assert.AreEqual("Reload XLEdge Add-in: Reload failed: disk full. Use Reload Add-in again, or restart Excel.", s.Ui.Errors.Single());
        }

        [TestMethod]
        public void Shutdown_DetachesPanesBeforeTheCoreShutsDown()
        {
            var s = new Setup();
            s.Host.Start();
            s.R.Events.Clear();
            s.Host.Shutdown();
            CollectionAssert.AreEqual(new[] { "Detach:vA", "Shutdown:vA", "Unload" }, s.R.Events);
        }

        [TestMethod]
        public void Reload_OpenWindowsDeclined_ChangesNothing()
        {
            var s = new Setup();
            s.Host.Start();
            s.Addins["vA"].Titles = new[] { "Options", "About XLEdge" };
            s.Ui.ConfirmAnswer = false;
            s.R.Events.Clear();
            Assert.IsFalse(s.Host.Reload(() => Setup.Release("vB")));
            CollectionAssert.AreEqual(new[] { "Confirm:Options|About XLEdge" }, s.R.Events);
            Assert.AreSame(s.Addins["vA"], s.Host.Addin);
        }

        [TestMethod]
        public void Reload_OpenWindowsAccepted_OrTitlesUnavailable_Proceeds()
        {
            var s = new Setup();
            s.Host.Start();
            s.Addins["vA"].Titles = new[] { "Options" };
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vB")));
            s.Addins["vB"].ThrowOnTitles = true;
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vC")), "an older core without the member counts as no open windows");
        }

        [TestMethod]
        public void Reload_ResolveReturnsNull_ReportsAfterProgressClosed()
        {
            var s = new Setup();
            s.Host.Start();
            s.R.Events.Clear();
            Assert.IsFalse(s.Host.Reload(() => null));
            int closed = s.R.Events.IndexOf("ProgressClosed"), error = s.R.Events.IndexOf("Error");
            Assert.IsTrue(closed >= 0 && error > closed, string.Join(", ", s.R.Events));
            Assert.AreEqual("Reload XLEdge Add-in: Reload failed - no usable XLEdge core version was found. The previous version has been closed; use Reload Add-in again.", s.Ui.Errors.Single());
            Assert.IsNull(s.Host.Addin);
            Assert.IsNull(s.Context.ActiveFolderName);
        }

        [TestMethod]
        public void Reload_NewCoreFails_AddinNullAndUsableAgain()
        {
            var s = new Setup();
            s.Host.Start();
            s.Loader.OnLoad = rel => throw new InvalidOperationException("contract too new");
            Assert.IsFalse(s.Host.Reload(() => Setup.Release("vB")));
            Assert.IsNull(s.Host.Addin);
            Assert.AreEqual("Reload XLEdge Add-in: Reload failed: contract too new. Use Reload Add-in again, or restart Excel.", s.Ui.Errors.Single());
            CollectionAssert.DoesNotContain(s.R.Events, "Remember:vB");
            Assert.AreEqual(1, s.R.Events.Count(e => e == "Attach"), "panes are attached only after a successful load");

            s.Loader.OnLoad = rel => s.Addins[rel.FolderName] = new FakeAddin(s.R, rel.FolderName);
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vA")));
            Assert.IsFalse(s.R.Events.Contains("Shutdown:vB"), "nothing was loaded, so nothing is shut down");
        }

        [TestMethod]
        public void Reload_ReentrantCall_Refused()
        {
            var s = new Setup();
            s.Host.Start();
            bool inner = true;
            Assert.IsTrue(s.Host.Reload(() => { inner = s.Host.Reload(() => Setup.Release("vX")); return Setup.Release("vB"); }));
            Assert.IsFalse(inner);
        }

        [TestMethod]
        public void Reload_ShutdownAndUnloadFailures_DoNotStopTheReload()
        {
            var s = new Setup();
            s.Host.Start();
            s.Addins["vA"].ThrowOnShutdown = true;
            s.Loader.ThrowOnUnload = true;
            Assert.IsTrue(s.Host.Reload(() => Setup.Release("vB")));
            Assert.AreSame(s.Addins["vB"], s.Host.Addin);
        }

        // ---- Shutdown ----------------------------------------------------------------------

        [TestMethod]
        public void Shutdown_NeverThrows()
        {
            var s = new Setup();
            s.Host.Start();
            s.Addins["vA"].ThrowOnShutdown = true;
            s.Loader.ThrowOnUnload = true;
            s.Host.Shutdown();
            Assert.IsNull(s.Host.Addin);
            s.Host.Shutdown();
        }

        // ---- real loader + bootstrapper (integration) --------------------------------------

        [TestMethod]
        public void Start_WithTheRealLoaderAndTheBuiltPackage_LoadsAndRemembers()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                foreach (var file in Directory.GetFiles(CorePackageTests.PackageFolder))
                    File.Copy(file, Path.Combine(paths.ManifestDirectory, Path.GetFileName(file)));
                var context = new FakeContext(paths) { Logger = new RecordingLogger() };
                var loader = new AddinDomainLoader();
                var host = new CoreHost(context, new DomainCoreLoader(loader), new BootstrapperReleaseCatalog(new UpdateBootstrapper()), new FakeUi(new Recorder()), new FakePanes(new Recorder()));
                try
                {
                    Assert.IsTrue(host.Start(), ((RecordingLogger)context.Logger).Dump());
                    Assert.AreEqual("XLEdgeCore_" + context.ActiveFolderName, loader.CurrentDomainName);
                    Assert.AreEqual(context.ActiveFolderName, File.ReadAllText(Path.Combine(t.Combine("inst"), "AddinCore", "SelectedRelease.txt")));
                }
                finally
                {
                    host.Shutdown();
                }
                Assert.IsNull(loader.CurrentDomainName);
            }
        }
    }
}
