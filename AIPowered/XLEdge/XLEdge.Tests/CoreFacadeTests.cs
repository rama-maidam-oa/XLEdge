using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// The core's facades keep today's XLEdge names and members (LogUtility, XLEdgeAppPaths, XLEdgeAppState) so ported code
    /// moves in unchanged; inside the core they are backed by the shell's context. Run inside a loader-like AppDomain.
    /// </summary>
    [TestClass]
    public class CoreFacadeTests
    {
        private const string CoreServices = "XLEdge.Addin.Core.Infrastructure.CoreServices";
        private const string Log = "XLEdge.Utilities.LogUtility";
        private const string Paths = "XLEdge.XLEdgeAppPaths";
        private const string State = "XLEdge.XLEdgeAppState";

        private sealed class Core : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();
            private readonly AppDomain _domain;

            public Core(bool initialize = true)
            {
                PathProvider = new PathProvider(_temp.Combine("inst"), _temp.Combine("data"));
                Logger = new RecordingLogger();
                Context = new FakeContext(PathProvider) { Logger = Logger };
                (_domain, Probe) = CoreProbe.Create("CoreFacade");
                if (initialize) Probe.CallStatic(CoreServices, "Initialize", Context);
            }

            public PathProvider PathProvider { get; }
            public RecordingLogger Logger { get; }
            public FakeContext Context { get; }
            public CoreProbe Probe { get; }

            public void Dispose()
            {
                try { Probe.CallStatic(CoreServices, "Shutdown"); }
                finally
                {
                    AppDomain.Unload(_domain);
                    _temp.Dispose();
                }
            }
        }

        [TestMethod]
        public void Paths_AreTheShellsPathProvider()
        {
            using (var core = new Core())
            {
                var p = core.PathProvider;
                string Get(string name) => (string)core.Probe.CallStatic(Paths, name);
                Assert.AreEqual(Path.GetDirectoryName(p.UrlsFile), Get("BaseFolder"));
                Assert.AreEqual(p.Root, Get("XLEdgeLogsFolder"));
                Assert.AreEqual(p.Logs, Get("LogFolder"));
                Assert.AreEqual(p.Temp, Get("TempFolder"));
                Assert.AreEqual(p.UrlsFile, Get("TempUrlsPath"));
                Assert.AreEqual(p.BrowserDataPath, Get("BrowserLogsFolder"));
                Assert.AreEqual(Path.Combine(p.BrowserDataPath, "Login"), Get("LoginBrowserLogsPath"));
                Assert.IsTrue(Directory.Exists(Get("LoginBrowserLogsPath")), "folders are created on first use, as today");
            }
        }

        [TestMethod]
        public void Paths_WithoutAContext_AreTodaysFolders()
        {
            using (var core = new Core(initialize: false))
            {
                string today = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"ORBIT\Excel_Logs");
                Assert.AreEqual(Path.Combine(today, @"XLEdge_Logs\Logs"), (string)core.Probe.CallStatic(Paths, "LogFolder"));
                Assert.AreEqual(Path.Combine(today, "ORBIT_URLS.xml"), (string)core.Probe.CallStatic(Paths, "TempUrlsPath"));
            }
        }

        /// <summary>Folders that cannot be created (a character Windows does not allow in a path).</summary>
        private sealed class UncreatablePaths : MarshalByRefObject, XLEdge.Contracts.IPathProvider
        {
            public override object InitializeLifetimeService() => null;
            public string Root => @"C:\bad|folder";
            public string Logs => Root + @"\Logs";
            public string Temp => Root + @"\Temp";
            public string BrowserDataPath => Root + @"\BrowserLogs";
            public string UrlsFile => @"C:\bad|folder\ORBIT_URLS.xml";
            public string VersionsPath => Root;
            public string SharedDependenciesPath => Root;
            public string ManifestFile => Root;
            public string ManifestDirectory => Root;
            public string ReleaseHistoryFile => Root;
            public string LatestVersion => null;
            public string LatestReleaseDate => null;
            public void Ensure() { }
            public void Refresh() { }
        }

        [TestMethod]
        public void Paths_ThatCannotBeCreated_AreReturnedAndLogged_AsToday()
        {
            using (var t = new TempFolder())
            {
                var logger = new RecordingLogger();
                var context = new FakeContext(new UncreatablePaths()) { Logger = logger };
                var (domain, probe) = CoreProbe.Create("CorePathsBad");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    Assert.AreEqual(@"C:\bad|folder\Logs", (string)probe.CallStatic(Paths, "LogFolder"));
                    Assert.IsTrue(logger.Lines.Any(l => l.StartsWith("WARN | EnsureDirectoryExists: failed to create directory 'C:\\bad|folder\\Logs'")), logger.Dump());
                    probe.CallStatic(CoreServices, "Shutdown");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void LogUtility_WritesThroughTheShellLogger()
        {
            using (var core = new Core())
            {
                core.Context.DebugMode = true;
                core.Probe.CallStatic(Log, "LogInfo", "an info");
                core.Probe.CallStatic(Log, "LogWarn", "a warning");
                core.Probe.CallStatic(Log, "LogError", "an error");
                core.Probe.CallStatic(Log, "LogDebug", "a debug line");
                CollectionAssert.AreEqual(new[] { "INFO | an info", "WARN | a warning", "ERROR | an error", "DEBUG | a debug line" }, core.Logger.Lines);
            }
        }

        [TestMethod]
        public void LogUtility_Exception_KeepsTodaysBlock()
        {
            using (var core = new Core())
            {
                core.Probe.LogThrownException("boom", "Refresh");
                string line = core.Logger.Lines.Single();
                StringAssert.StartsWith(line, "EXCEPTION | Refresh|");
                foreach (string part in new[] { "Type: System.InvalidOperationException", "Message: boom", "Source: XLEdge.Tests", "TargetSite: ", "StackTrace:",
                    "LogThrownException", "----- Inner Exception -----", "Type: System.ArgumentException", "Message: inner boom" })
                    StringAssert.Contains(line, part);
            }
        }

        [TestMethod]
        public void LogUtility_RawJson_FullPayloadOnlyWithDebugAndOutputData()
        {
            using (var core = new Core())
            {
                string payload = "{\"rows\":[" + string.Join(",", Enumerable.Repeat("1", 200)) + "]}";
                core.Probe.CallStatic(Log, "LogRawJson", "fetch", payload);
                core.Context.DebugMode = true;
                core.Probe.OnSingleton(State, "Instance", "DebugOutputData", true);
                core.Probe.CallStatic(Log, "LogRawJson", "fetch", payload);
                var lines = core.Logger.Lines;
                Assert.AreEqual(2, lines.Length, string.Join("\n", lines));
                StringAssert.Contains(lines[0], "enable Debug mode + 'Include Output Data' to log the full payload");
                Assert.IsFalse(lines[0].Contains(payload));
                StringAssert.Contains(lines[1], payload);
            }
        }

        [TestMethod]
        public void LogUtility_ScopesAndFlushes_ReachTheShellLogger()
        {
            using (var core = new Core())
            {
                core.Context.DebugMode = true;
                core.Probe.RunInLogScope("Refresh", "inside");
                core.Probe.CallStatic(Log, "FlushAllOpenBuffers", "Debug session ended");
                CollectionAssert.AreEqual(new[] { "begin:Refresh", "end:Refresh", "flush:Debug session ended" }, core.Logger.Events);
                CollectionAssert.Contains(core.Logger.Lines, "DEBUG | inside");
            }
        }

        [TestMethod]
        public void LogUtility_BeforeInitializeAndAfterShutdown_DoesNotThrow()
        {
            using (var core = new Core(initialize: false))
            {
                core.Probe.CallStatic(Log, "LogInfo", "no context yet");
                core.Probe.CallStatic(Log, "LogDebug", "no context yet");
                core.Probe.LogThrownException("no context", "early");
                core.Probe.RunInLogScope("early", "x");
                core.Probe.CallStatic(Log, "FlushAllOpenBuffers", "early");
                core.Probe.CallStatic(CoreServices, "Initialize", core.Context);
                core.Probe.CallStatic(CoreServices, "Shutdown");
                core.Probe.CallStatic(Log, "LogWarn", "after shutdown");
                Assert.AreEqual(0, core.Logger.Lines.Length, "nothing reaches a logger the core is not attached to");
            }
        }

        [TestMethod]
        public void DebugMode_IsTheShellsDebugToggle()
        {
            using (var core = new Core())
            {
                core.Probe.CallStatic(Log, "LogDebug", "hidden");
                Assert.IsFalse((bool)core.Probe.CallStatic(Log, "DebugMode"));
                core.Context.DebugMode = true;
                Assert.IsTrue((bool)core.Probe.CallStatic(Log, "DebugMode"));
                Assert.IsTrue((bool)core.Probe.OnSingleton(State, "Instance", "DebugLogs"));
                core.Probe.OnSingleton(State, "Instance", "DebugLogs", false);
                Assert.IsFalse(core.Context.DebugMode, "the core turning debug off turns the shell's toggle off");
                Assert.AreEqual(0, core.Logger.Lines.Length);
            }
        }

        [TestMethod]
        public void AppState_KeepsTodaysDefaults_AndHasNoExcelWithoutOne()
        {
            using (var core = new Core())
            {
                object Get(string name) => core.Probe.OnSingleton(State, "Instance", name);
                Assert.IsNull(Get("ExcelApp"));
                Assert.AreEqual(IntPtr.Zero, Get("ExcelHandle"));
                Assert.AreEqual(true, Get("ParamDataSameSheet"));
                Assert.AreEqual(true, Get("ShowSegmentSelectionWindow"));
                Assert.AreEqual(false, Get("PreserveFormats"));
                Assert.AreEqual(false, Get("LoginFromGLSense"));
                Assert.AreEqual(string.Empty, Get("DrillPostData"));
                core.Probe.OnSingleton(State, "Instance", "LoginToken", "t");
                Assert.AreEqual("t", Get("LoginToken"));
            }
        }
    }
}
