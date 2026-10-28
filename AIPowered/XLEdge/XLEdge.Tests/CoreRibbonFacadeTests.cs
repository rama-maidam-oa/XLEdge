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
        public void Current_IsTheHelperTheCoreMadeLast()
        {
            using (var core = new Core())
            {
                Assert.IsTrue(core.Probe.HasStatic(Ribbon, "Current"));
            }
        }

        [TestMethod]
        public void Wpf_RunsInTheCoresDomain_AndLeavesNothingAfterUnload()
        {
            Win32Test.OnSta(() =>
            {
                // Pumping (Application.DoEvents) creates WinForms' hidden thread window in this test's own AppDomain:
                // do it before the snapshot so it is not mistaken for a window of the core.
                Win32Test.Pump(TimeSpan.FromMilliseconds(50));
                var before = Win32Test.ProcessWindows();
                var (domain, probe) = CoreProbe.Create("CoreWpf");
                try
                {
                    Assert.AreEqual("app=True thread=caller", probe.BootWpfAndRunOnTheDispatcher());
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
                bool Managed(string cls) => cls.StartsWith("HwndWrapper", StringComparison.Ordinal) || cls.StartsWith("WindowsForms10.", StringComparison.Ordinal)
                                            || cls.StartsWith(".NET-", StringComparison.Ordinal);
                Win32Test.PumpUntil(() => !Win32Test.ProcessWindows().Any(w => !before.ContainsKey(w.Key) && Managed(w.Value)), TimeSpan.FromSeconds(5));
                var left = Win32Test.ProcessWindows().Where(w => !before.ContainsKey(w.Key) && Managed(w.Value)).Select(w => w.Value).ToList();
                Assert.AreEqual(0, left.Count, "managed windows left: " + string.Join(", ", left));
                Win32Test.Pump(TimeSpan.FromSeconds(2));
            });
        }
    }
}
