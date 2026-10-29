using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// XLEdge 11.1.2's ribbon helper moves into the core with its state machine (hot-reloadable); only its lowest layer -
    /// finding an Add-in Express control by reflection - becomes calls to the shell's IRibbonController. Also: the WPF
    /// Application/dispatcher the ported windows need, created inside the core's AppDomain.
    /// </summary>
    [TestClass]
    public class CoreRibbonFacadeTests
    {
        private const string CoreServices = "XLEdge.Addin.Core.Infrastructure.CoreServices";
        private const string Ribbon = "XLEdge.Helpers.XLEdgeRibbonHelper";
        private const string State = "XLEdge.XLEdgeAppState";

        /// <summary>Records every call the core makes on the ribbon (the shell's RibbonController in the add-in).</summary>
        internal sealed class FakeRibbon : MarshalByRefObject, IRibbonController
        {
            private readonly List<string> _calls = new List<string>();
            public override object InitializeLifetimeService() => null;
            public string[] Calls { get { lock (_calls) return _calls.ToArray(); } }
            public void Clear() { lock (_calls) _calls.Clear(); }
            private void Add(string call) { lock (_calls) _calls.Add(call); }
            public Dictionary<string, bool> Pressed { get; } = new Dictionary<string, bool>();
            public Dictionary<string, bool> Visible { get; } = new Dictionary<string, bool>();
            public Dictionary<string, string> Labels { get; } = new Dictionary<string, string>();

            public void ExecuteAction(string controlId) => Add("Execute:" + controlId);
            public void SetState(string stateName) => Add("SetState:" + stateName);
            public void SetControlEnabled(string controlId, bool enabled) => Add($"Enabled:{controlId}={enabled}");
            public void SetControlVisible(string controlId, bool visible) { Visible[controlId] = visible; Add($"Visible:{controlId}={visible}"); }
            public void SetControlPressed(string controlId, bool pressed) { Pressed[controlId] = pressed; Add($"Pressed:{controlId}={pressed}"); }
            public bool GetControlPressed(string controlId) => Pressed.TryGetValue(controlId, out bool b) && b;
            public bool GetControlVisible(string controlId) => Visible.TryGetValue(controlId, out bool b) && b;
            public void SetControlLabel(string controlId, string label) { Labels[controlId] = label; Add($"Label:{controlId}={label}"); }
            public string GetControlLabel(string controlId) => Labels.TryGetValue(controlId, out string s) ? s : string.Empty;
            public void EnableControls(string[] controlIds) => Add("EnableControls:" + string.Join(",", controlIds));
            public void DisableControls(string[] controlIds) => Add("DisableControls:" + string.Join(",", controlIds));
            public void Invalidate(string controlId) => Add("Invalidate:" + controlId);
            public void InvalidateAll() => Add("InvalidateAll");
            public void ShowTaskPane() => Add("ShowTaskPane");
            public void ToggleTaskPane() => Add("ToggleTaskPane");
            public void HideAllTaskPanes() => Add("HideAllTaskPanes");
            public bool IsTaskPaneVisible() => false;
            public void SetTaskPaneCaption(string caption) => Add("Caption:" + caption);

            // Contract 4: one pane per Excel window, as the shell keeps them.
            public IntPtr ActiveWindow { get; set; }
            public Dictionary<IntPtr, bool> PaneVisible { get; } = new Dictionary<IntPtr, bool>();
            public Dictionary<IntPtr, string> PaneCaption { get; } = new Dictionary<IntPtr, string>();
            public Dictionary<IntPtr, int> PaneWidth { get; } = new Dictionary<IntPtr, int>();
            public IntPtr GetActiveTaskPaneWindow() => ActiveWindow;
            public bool GetTaskPaneVisible(IntPtr excelWindow) { lock (PaneVisible) return PaneVisible.TryGetValue(excelWindow, out bool v) && v; }
            public void SetTaskPaneVisible(IntPtr excelWindow, bool visible) { lock (PaneVisible) PaneVisible[excelWindow] = visible; Add($"PaneVisible:{excelWindow}={visible}"); }
            public void SetTaskPaneCaption(IntPtr excelWindow, string caption) { lock (PaneCaption) PaneCaption[excelWindow] = caption; Add($"PaneCaption:{excelWindow}={caption}"); }
            public int GetTaskPaneWidth(IntPtr excelWindow) { lock (PaneWidth) return PaneWidth.TryGetValue(excelWindow, out int w) ? w : 400; }
            public void SetTaskPaneWidth(IntPtr excelWindow, int width) { lock (PaneWidth) PaneWidth[excelWindow] = width; Add($"PaneWidth:{excelWindow}={width}"); }
        }

        private sealed class Core : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();
            private readonly AppDomain _domain;

            public Core()
            {
                Ribbon = new FakeRibbon();
                Logger = new RecordingLogger();
                Context = new FakeContext(new PathProvider(_temp.Combine("inst"), _temp.Combine("data"))) { Logger = Logger };
                Context.SetRibbonController(Ribbon);
                (_domain, Probe) = CoreProbe.Create("CoreRibbon");
                Probe.CallStatic(CoreServices, "Initialize", Context);
                Helper = Probe.CreateInstance(CoreRibbonFacadeTests.Ribbon, Ribbon);
            }

            public FakeRibbon Ribbon { get; }
            public RecordingLogger Logger { get; }
            public FakeContext Context { get; }
            public CoreProbe Probe { get; }
            public int Helper { get; }

            public object Call(string method, params object[] args) => Probe.CallOnInstance(Helper, method, args);

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

        private static readonly string[] Feature8 =
            { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide", "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp" };

        [TestMethod]
        public void Members_ReachTheShellsRibbon()
        {
            using (var core = new Core())
            {
                core.Call("SetControlEnabled", "RibEdgeRefresh", false);
                core.Call("SetControlPressed", "RibEdgeDebug", true);
                core.Call("SetControlVisible", "RibEdgeLogout", true);
                core.Call("SetControlCaption", "RibEdgeLogout", "Prod");
                CollectionAssert.AreEqual(new[] { "Enabled:RibEdgeRefresh=False", "Pressed:RibEdgeDebug=True", "Visible:RibEdgeLogout=True", "Label:RibEdgeLogout=Prod" },
                    core.Ribbon.Calls);
                Assert.AreEqual(true, core.Call("GetControlPressed", "RibEdgeDebug"));
                Assert.AreEqual(true, core.Call("GetControlVisible", "RibEdgeLogout"));
                Assert.AreEqual("Prod", core.Call("GetControlCaption", "RibEdgeLogout"));
                Assert.AreEqual("Prod", core.Call("GetControlLabel", "RibEdgeLogout"));
            }
        }

        [TestMethod]
        public void Enabled_IsRemembered_AndRestorePreviousStateReappliesIt()
        {
            // The contract cannot read a control's Enabled state; today's helper already kept what it set.
            using (var core = new Core())
            {
                core.Call("SetControlEnabled", "RibEdgeRefresh", false);
                core.Call("SetControlEnabled", "RibEdgeHelp", true);
                Assert.AreEqual(false, core.Call("GetControlEnabled", "RibEdgeRefresh"));
                Assert.AreEqual(true, core.Call("GetControlEnabled", "RibEdgeHelp"));
                core.Ribbon.Clear();
                core.Call("RestorePreviousState");
                CollectionAssert.AreEquivalent(new[] { "Enabled:RibEdgeRefresh=False", "Enabled:RibEdgeHelp=True" }, core.Ribbon.Calls);
            }
        }

        [TestMethod]
        public void LoggedOut_IsTodaysSequence_AndTurnsDebugOff()
        {
            using (var core = new Core())
            {
                core.Context.DebugMode = true;
                core.Call("ApplyState", "LoggedOut");
                var expected = Feature8.Select(id => $"Enabled:{id}=False")
                    .Concat(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }.Select(id => $"Enabled:{id}=True"))
                    .Concat(new[] { "Visible:RibEdgeLogin=True", "Visible:RibEdgeLogout=False", "Pressed:RibEdgeDebug=False", "Pressed:RibEdgeIncludeOutputData=False",
                        "Label:RibEdgeLogin=Login", "InvalidateAll" }).ToArray();
                CollectionAssert.AreEqual(expected, core.Ribbon.Calls);
                // Fixed: today the Debug button was un-pressed while debug logging stayed on.
                Assert.IsFalse(core.Context.DebugMode, "an un-pressed Debug button means debug logging is off");
            }
        }

        [TestMethod]
        public void NoPermission_IsTodaysSequence_AndTurnsDebugOff()
        {
            // XLEdge 11.1.2's constant spelling ("NoXLEdePermission"); the shell's SetState spells it "NoXLEdgePermission".
            using (var core = new Core())
            {
                core.Context.DebugMode = true;
                core.Call("ApplyState", "NoXLEdePermission");
                var expected = new[] { "Pressed:RibEdgeIncludeOutputData=False", "Pressed:RibEdgeDebug=False" }
                    .Concat(Feature8.Concat(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }).Select(id => $"Enabled:{id}=False"))
                    .Concat(new[] { "InvalidateAll" }).ToArray();
                CollectionAssert.AreEqual(expected, core.Ribbon.Calls);
                Assert.IsFalse(core.Context.DebugMode, "an un-pressed Debug button means debug logging is off");
            }
        }

        [TestMethod]
        public void LoggedIn_ShowsLogoutWithTheInstanceName()
        {
            using (var core = new Core())
            {
                core.Probe.OnSingleton(State, "Instance", "LoginUrlName", "Production");
                core.Call("ApplyState", "LoggedIn");
                var expected = Feature8.Select(id => $"Enabled:{id}=True")
                    .Concat(new[] { "Visible:RibEdgeLogin=False", "Visible:RibEdgeLogout=True", "Pressed:RibEdgeIncludeOutputData=False",
                        "Label:RibEdgeLogout=Production", "InvalidateAll" }).ToArray();
                CollectionAssert.AreEqual(expected, core.Ribbon.Calls, "not logged in yet: no workbook-dependent state");
            }
        }

        [TestMethod]
        public void UnknownOrEmptyState_IsLoggedNotApplied()
        {
            using (var core = new Core())
            {
                core.Call("ApplyState", "NoSuchState");
                core.Call("ApplyState", "");
                CollectionAssert.AreEqual(new[] { "InvalidateAll" }, core.Ribbon.Calls, "an unknown state still refreshes, as today");
                Assert.IsTrue(core.Logger.Contains("[XLEdgeRibbonHelper] Unknown state: NoSuchState"), core.Logger.Dump());
                Assert.IsTrue(core.Logger.Contains("ApplyState called with an empty state name"), core.Logger.Dump());
            }
        }

        [TestMethod]
        public void Initialize_MakesTheRibbonHelperTalkToTheShellsController()
        {
            // Ported code reaches the ribbon through XLEdgeRibbonHelper.Current; AddinEntry.Initialize must create it.
            using (var core = new Core())
            {
                core.Ribbon.Clear();
                core.Probe.InitializeTheEntryAndDisableThroughTheRibbonHelper(core.Context, "RibEdgeRefresh");
                CollectionAssert.Contains(core.Ribbon.Calls, "Enabled:RibEdgeRefresh=False");
            }
        }

        [TestMethod]
        public void Wpf_RunsOnTheCallersThread()
        {
            Win32Test.OnSta(() =>
            {
                var (domain, probe) = CoreProbe.Create("CoreWpf");
                try { Assert.AreEqual("app=True thread=caller", probe.BootWpfAndRunOnTheDispatcher()); }
                finally { AppDomain.Unload(domain); }
            });
        }

        [TestMethod]
        public void CoreShutdown_ClosesWpfsDispatcherWindow_AndTheUnloadLeavesNothing()
        {
            // WPF's dispatcher window (a message-only HwndWrapper) has its window procedure in the core's domain.
            Win32Test.OnSta(() =>
            {
                // Pumping (Application.DoEvents) creates WinForms' hidden thread window in this test's own AppDomain:
                // do it before the snapshot so it is not mistaken for a window of the core.
                Win32Test.Pump(TimeSpan.FromMilliseconds(50));
                var before = Win32Test.ProcessWindows();
                bool Managed(string cls) => cls.StartsWith("HwndWrapper", StringComparison.Ordinal) || cls.StartsWith("WindowsForms10.", StringComparison.Ordinal)
                                            || cls.StartsWith(".NET-", StringComparison.Ordinal);
                System.Collections.Generic.List<string> Added() =>
                    Win32Test.ProcessWindows().Where(w => !before.ContainsKey(w.Key) && Managed(w.Value)).Select(w => w.Value).ToList();

                using (var t = new TempFolder())
                {
                    var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data"))) { Logger = new RecordingLogger() };
                    context.SetRibbonController(new FakeRibbon());
                    var (domain, probe) = CoreProbe.Create("CoreWpfShutdown");
                    try
                    {
                        probe.InitializeTheEntryWithWpf(context);
                        Assert.IsTrue(Added().Any(c => c.StartsWith("HwndWrapper", StringComparison.Ordinal)),
                            "precondition: the core's WPF dispatcher window exists: " + string.Join(", ", Added()));
                        probe.ShutDownTheEntry();
                        Win32Test.PumpUntil(() => !Added().Any(), TimeSpan.FromSeconds(5));
                        Assert.AreEqual(0, Added().Count, "AddinEntry.Shutdown closes WPF: " + string.Join(", ", Added()));
                    }
                    finally
                    {
                        // As the loader does it: on a pool thread while Excel's thread (this one) keeps pumping.
                        var unload = System.Threading.Tasks.Task.Run(() => AppDomain.Unload(domain));
                        Win32Test.PumpUntil(() => unload.IsCompleted, TimeSpan.FromSeconds(10));
                        unload.Wait();
                    }
                }
                Win32Test.Pump(TimeSpan.FromSeconds(1));
                Assert.AreEqual(0, Added().Count, "managed windows left after the unload: " + string.Join(", ", Added()));
            });
        }
    }
}
