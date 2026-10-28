using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// The core reads and writes the same preferences file as XLEdge 11.1.2 (XLEdge_Logs\xledgeuserpreferences.json), in the
    /// same format (GoldenValues PREFS lines come from 11.1.2's own serializer), with today's rules for missing keys and damaged
    /// files - and loads it at start-up, so preferences also come back after a hot reload.
    /// </summary>
    [TestClass]
    public class CorePreferencesTests
    {
        private const string CoreServices = "XLEdge.Addin.Core.Infrastructure.CoreServices";
        private const string Manager = "XLEdge.Helpers.XLEdgePreferencesManager";
        private const string State = "XLEdge.XLEdgeAppState";

        private static string Golden(string flags)
        {
            string line = GoldenValues.Lines.Single(l => l.StartsWith($"PREFS.Serialize(\"{flags}\") = s:", StringComparison.Ordinal));
            return line.Substring(line.IndexOf("= s:", StringComparison.Ordinal) + 4);
        }

        private sealed class Core : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();
            private readonly AppDomain _domain;

            public Core()
            {
                Paths = new PathProvider(_temp.Combine("inst"), _temp.Combine("data"));
                Paths.Ensure();
                Logger = new RecordingLogger();
                Context = new FakeContext(Paths) { Logger = Logger };
                Context.SetRibbonController(new CoreRibbonFacadeTests.FakeRibbon());
                (_domain, Probe) = CoreProbe.Create("CorePreferences");
            }

            public PathProvider Paths { get; }
            public RecordingLogger Logger { get; }
            public FakeContext Context { get; }
            public CoreProbe Probe { get; }
            public string File => Path.Combine(Paths.Root, "xledgeuserpreferences.json");

            public void LoadPreferences()
            {
                Probe.CallStatic(CoreServices, "Initialize", Context);
                Probe.CallOnSingleton(Manager, "Instance", "Initialize");
            }

            public object Get(string property) => Probe.OnSingleton(State, "Instance", property);

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
        public void NoFile_WritesTodaysDefaults_InTodaysFormat()
        {
            using (var core = new Core())
            {
                core.LoadPreferences();
                Assert.AreEqual(Golden("1000010"), System.IO.File.ReadAllText(core.File), "XLEdge 11.1.2 writes this for a new user");
                Assert.AreEqual(true, core.Get("ParamDataSameSheet"));
                Assert.AreEqual(true, core.Get("ShowSegmentSelectionWindow"));
                Assert.AreEqual(false, core.Get("PreserveFormats"));
            }
        }

        [TestMethod]
        public void TodaysFile_IsApplied_AndLeftAsItIs()
        {
            using (var core = new Core())
            {
                // Written by XLEdge 11.1.2: every flag flipped from its default.
                string today = Golden("0111101");
                System.IO.File.WriteAllText(core.File, today);
                core.LoadPreferences();
                Assert.AreEqual(false, core.Get("ParamDataSameSheet"));
                Assert.AreEqual(true, core.Get("SchOutputsToSameSheet"));
                Assert.AreEqual(true, core.Get("RefreshSync"));
                Assert.AreEqual(true, core.Get("AllowSheetNameChanges"));
                Assert.AreEqual(true, core.Get("ShowCalendarControl"));
                Assert.AreEqual(false, core.Get("ShowSegmentSelectionWindow"));
                Assert.AreEqual(true, core.Get("PreserveFormats"));
                Assert.AreEqual(today, System.IO.File.ReadAllText(core.File), "a complete file is not rewritten");
            }
        }

        [TestMethod]
        public void MissingKey_GetsItsDefault_AndTheFileIsCompleted()
        {
            using (var core = new Core())
            {
                // A file from before "Preserve formats" existed.
                System.IO.File.WriteAllText(core.File, "{\"parameterValues\":false,\"scheduledOutputs\":false,\"refreshSync\":false,\"changeSheetName\":false,\"calendarCtrlDisplay\":false,\"segmentSelectionWindowDisplay\":true}");
                core.LoadPreferences();
                Assert.AreEqual(false, core.Get("ParamDataSameSheet"), "saved values are kept");
                Assert.AreEqual(false, core.Get("PreserveFormats"), "the missing one gets today's default");
                Assert.AreEqual(Golden("0000010"), System.IO.File.ReadAllText(core.File));
            }
        }

        [TestMethod]
        public void DamagedFile_FallsBackToDefaults_AndIsRewritten()
        {
            using (var core = new Core())
            {
                System.IO.File.WriteAllText(core.File, "{ not json");
                core.LoadPreferences();
                Assert.AreEqual(true, core.Get("ShowSegmentSelectionWindow"));
                Assert.AreEqual(Golden("1000010"), System.IO.File.ReadAllText(core.File));
                Assert.IsTrue(core.Logger.Lines.Any(l => l.StartsWith("EXCEPTION | Failed reading preference file.", StringComparison.Ordinal)), core.Logger.Dump());
            }
        }

        [TestMethod]
        public void CoreStartup_LoadsThePreferences()
        {
            // 11.1.2 loads them when the ribbon loads (AddinModule.OnRibbonLoaded); the core does at Initialize - which also
            // brings them back after every hot reload.
            using (var core = new Core())
            {
                System.IO.File.WriteAllText(core.File, Golden("1000000"));
                core.Probe.InitializeTheEntry(core.Context);
                try
                {
                    Assert.AreEqual(false, core.Get("ShowSegmentSelectionWindow"));
                    Assert.AreEqual(true, core.Get("ParamDataSameSheet"));
                }
                finally
                {
                    core.Probe.ShutDownTheEntry();
                }
            }
        }
    }
}
