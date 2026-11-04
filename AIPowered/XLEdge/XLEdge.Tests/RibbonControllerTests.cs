using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Hosting;

namespace XLEdge.Tests
{
    [TestClass]
    public class RibbonControllerTests
    {
        public sealed class FakeControl
        {
            public FakeControl(string id) { Id = id; }
            public string Id { get; }
            public bool Enabled { get; set; } = true;
            public bool Visible { get; set; } = true;
            public bool Pressed { get; set; }
            public string Caption { get; set; } = "";
        }

        /// <summary>Stands in for AddinModule: one field per ribbon control, named like the real ones.</summary>
        private sealed class FakeControls
        {
            private FakeControl RibEdgeLogin = new FakeControl("idLogin");
            private FakeControl RibEdgeLogout = new FakeControl("idLogout");
            private FakeControl RibEdgeBroadcast = new FakeControl("idBroadcast");
            private FakeControl RibEdgeDialogBoxLauncher = new FakeControl("idDbl");
            private FakeControl RibEdgeRefresh = new FakeControl("idRefresh");
            private FakeControl RibEdgeRefreshAll = new FakeControl("idRefreshAll");
            private FakeControl RibEdgeParamRefresh = new FakeControl("idParam");
            private FakeControl RibEdgeShowHide = new FakeControl("idShowHide");
            private FakeControl RibEdgeDebug = new FakeControl("idDebug");
            private FakeControl RibEdgeOptions = new FakeControl("idOptions");
            private FakeControl RibControlSheet = new FakeControl("idControlSheet");
            private FakeControl RibEdgeAbout = new FakeControl("idAbout");
            private FakeControl RibEdgeHelp = new FakeControl("idHelp");

            public FakeControl Get(string name) => (FakeControl)GetType().GetField(name, System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(this);
        }

        private sealed class FakeShell : IShellCommands
        {
            public readonly List<string> Calls = new List<string>();
            public readonly List<Action> Posted = new List<Action>();
            public bool PaneVisible;
            public void Post(Action action) { Calls.Add("Post"); Posted.Add(action); }
            public void RunRibbonAction(string controlId) => Calls.Add("Run:" + controlId);
            public void InvalidateRibbon() => Calls.Add("InvalidateAll");
            public void InvalidateControl(string ribbonXId) => Calls.Add("Invalidate:" + ribbonXId);
            public void DebugModeChanged(bool enabled) => Calls.Add("DebugMode:" + enabled);
            public void ShowTaskPane() => Calls.Add("ShowPane");
            public void ToggleTaskPane() => Calls.Add("TogglePane");
            public void HideAllTaskPanes() => Calls.Add("HideAllPanes");
            public bool IsTaskPaneVisible() => PaneVisible;
            public void SetTaskPaneCaption(string caption) => Calls.Add("Caption:" + caption);
            public IntPtr ActiveWindow;
            public IntPtr GetActiveTaskPaneWindow() => ActiveWindow;
            public bool GetTaskPaneVisible(IntPtr excelWindow) { Calls.Add("Visible?:" + excelWindow); return PaneVisible; }
            public void SetTaskPaneVisible(IntPtr excelWindow, bool visible) => Calls.Add($"Visible:{excelWindow}={visible}");
            public void SetTaskPaneCaption(IntPtr excelWindow, string caption) => Calls.Add($"Caption:{excelWindow}={caption}");
            public int GetTaskPaneWidth(IntPtr excelWindow) { Calls.Add("Width?:" + excelWindow); return 480; }
            public void SetTaskPaneWidth(IntPtr excelWindow, int width) => Calls.Add($"Width:{excelWindow}={width}");
        }

        private static (RibbonController Controller, FakeControls Controls, FakeShell Shell) Create()
        {
            var controls = new FakeControls();
            var shell = new FakeShell();
            return (new RibbonController(controls, shell, new RecordingLogger()), controls, shell);
        }

        private static readonly string[] Feature8 =
            { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide", "RibEdgeOptions", "RibEdgeHelp",
              "RibEdgeBroadcast" }; // OISR-22777: Broadcast follows the login like the others

        [TestMethod]
        public void PaneOfAnExcelWindow_GoesToThatWindowsPane()
        {
            // Contract 4: 11.1.2's pane code acts on its own pane (Visible, Text, Width) - one per Excel window.
            var (c, _, shell) = Create();
            shell.ActiveWindow = new IntPtr(0x2222);
            shell.PaneVisible = true;
            Assert.AreEqual(new IntPtr(0x2222), c.GetActiveTaskPaneWindow());
            Assert.IsTrue(c.GetTaskPaneVisible(new IntPtr(0x1111)));
            c.SetTaskPaneVisible(new IntPtr(0x1111), false);
            c.SetTaskPaneCaption(new IntPtr(0x1111), "https://instance");
            c.SetTaskPaneCaption(new IntPtr(0x1111), null);
            Assert.AreEqual(480, c.GetTaskPaneWidth(new IntPtr(0x1111)));
            c.SetTaskPaneWidth(new IntPtr(0x1111), 530);
            CollectionAssert.AreEqual(new[] { "Visible?:4369", "Visible:4369=False", "Caption:4369=https://instance", "Caption:4369=", "Width?:4369", "Width:4369=530" },
                shell.Calls.ToArray());
        }

        [TestMethod]
        public void LoggedOut_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            controls.Get("RibEdgeDebug").Pressed = true;
            c.SetState("LoggedOut");

            foreach (var id in Feature8) Assert.IsFalse(controls.Get(id).Enabled, id);
            foreach (var id in new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }) Assert.IsTrue(controls.Get(id).Enabled, id);
            Assert.IsTrue(controls.Get("RibEdgeLogin").Visible);
            Assert.IsFalse(controls.Get("RibEdgeLogout").Visible);
            Assert.IsFalse(controls.Get("RibEdgeDebug").Pressed);
            Assert.AreEqual("Login", controls.Get("RibEdgeLogin").Caption);
            CollectionAssert.Contains(shell.Calls, "DebugMode:False", "the Debug box is unchecked, so debug logging stops too");
            Assert.AreEqual("InvalidateAll", shell.Calls.Last());
        }

        [TestMethod]
        public void LoggedIn_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            c.SetState("LoggedOut");
            c.SetState("LoggedIn");
            foreach (var id in Feature8) Assert.IsTrue(controls.Get(id).Enabled, id);
            Assert.IsFalse(controls.Get("RibEdgeLogin").Visible);
            Assert.IsTrue(controls.Get("RibEdgeLogout").Visible);
        }

        [TestMethod]
        public void NoXLEdgePermission_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            controls.Get("RibEdgeDebug").Pressed = true;
            c.SetState("NoXLEdgePermission");
            foreach (var id in Feature8.Concat(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }))
                Assert.IsFalse(controls.Get(id).Enabled, id);
            Assert.IsTrue(controls.Get("RibEdgeLogout").Enabled, "a user without XLEdge permission can still log out");
            Assert.IsFalse(controls.Get("RibEdgeDebug").Pressed);
            CollectionAssert.Contains(shell.Calls, "DebugMode:False", "the shell's logger leaves debug mode with the Debug toggle");
        }

        [TestMethod]
        public void ExecuteAction_ReloadIsQueued_OthersRunNow()
        {
            // The core may call ExecuteAction; a reload run on that call would unload the core's AppDomain under its own stack.
            var (c, _, shell) = Create();
            c.ExecuteAction(RibbonControlIds.RibReload);
            c.ExecuteAction(RibbonControlIds.RibVersionHistory);
            CollectionAssert.AreEqual(new[] { "Post", "Post" }, shell.Calls);
            c.ExecuteAction(RibbonControlIds.RibEdgeRefresh);
            CollectionAssert.AreEqual(new[] { "Post", "Post", "Run:RibEdgeRefresh" }, shell.Calls);
            foreach (var a in shell.Posted) a();
            CollectionAssert.AreEqual(new[] { "Post", "Post", "Run:RibEdgeRefresh", "Run:RibReload", "Run:RibVersionHistory" }, shell.Calls);
        }

        [TestMethod]
        public void Controls_ReadWriteInvalidateAndMissing()
        {
            var (c, controls, shell) = Create();
            c.SetControlLabel("RibEdgeLogout", "orbit-prod");
            Assert.AreEqual("orbit-prod", c.GetControlLabel("RibEdgeLogout"));
            c.SetControlVisible("RibEdgeHelp", false);
            Assert.IsFalse(c.GetControlVisible("RibEdgeHelp"));
            c.SetControlPressed("RibEdgeDebug", true);
            Assert.IsTrue(c.GetControlPressed("RibEdgeDebug"));
            c.Invalidate("RibEdgeRefresh");
            CollectionAssert.Contains(shell.Calls, "Invalidate:idRefresh", "Office needs the RibbonX Id, not the field name");
            Assert.AreEqual("", c.GetControlLabel("DoesNotExist"));
            Assert.IsFalse(c.GetControlVisible("DoesNotExist"));
            c.SetControlEnabled("DoesNotExist", true);
        }

        [TestMethod]
        public void TaskPaneCallsGoToTheShell()
        {
            var (c, _, shell) = Create();
            shell.PaneVisible = true;
            c.ShowTaskPane(); c.ToggleTaskPane(); c.HideAllTaskPanes(); c.SetTaskPaneCaption("https://orbit");
            Assert.IsTrue(c.IsTaskPaneVisible());
            CollectionAssert.AreEqual(new[] { "ShowPane", "TogglePane", "HideAllPanes", "Caption:https://orbit" }, shell.Calls);
            Assert.IsNull(c.InitializeLifetimeService());
        }
    }
}
