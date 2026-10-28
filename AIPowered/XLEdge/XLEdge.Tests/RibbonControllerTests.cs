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
            private FakeControl RibEdgeDialogBoxLauncher = new FakeControl("idDbl");
            private FakeControl RibEdgeRefresh = new FakeControl("idRefresh");
            private FakeControl RibEdgeRefreshAll = new FakeControl("idRefreshAll");
            private FakeControl RibEdgeParamRefresh = new FakeControl("idParam");
            private FakeControl RibEdgeShowHide = new FakeControl("idShowHide");
            private FakeControl RibEdgeDebug = new FakeControl("idDebug");
            private FakeControl RibEdgeIncludeOutputData = new FakeControl("idOutput");
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
        }

        private static (RibbonController Controller, FakeControls Controls, FakeShell Shell) Create()
        {
            var controls = new FakeControls();
            var shell = new FakeShell();
            return (new RibbonController(controls, shell, new RecordingLogger()), controls, shell);
        }

        private static readonly string[] Feature8 =
            { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide", "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp" };

        [TestMethod]
        public void LoggedOut_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            controls.Get("RibEdgeDebug").Pressed = true;
            controls.Get("RibEdgeIncludeOutputData").Pressed = true;
            c.SetState("LoggedOut");

            foreach (var id in Feature8) Assert.IsFalse(controls.Get(id).Enabled, id);
            foreach (var id in new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }) Assert.IsTrue(controls.Get(id).Enabled, id);
            Assert.IsTrue(controls.Get("RibEdgeLogin").Visible);
            Assert.IsFalse(controls.Get("RibEdgeLogout").Visible);
            Assert.IsFalse(controls.Get("RibEdgeDebug").Pressed);
            Assert.IsFalse(controls.Get("RibEdgeIncludeOutputData").Pressed);
            Assert.AreEqual("Login", controls.Get("RibEdgeLogin").Caption);
            CollectionAssert.Contains(shell.Calls, "DebugMode:False", "the Debug box is unchecked, so debug logging stops too");
            Assert.AreEqual("InvalidateAll", shell.Calls.Last());
        }

        [TestMethod]
        public void LoggedIn_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            c.SetState("LoggedOut");
            controls.Get("RibEdgeIncludeOutputData").Pressed = true;
            c.SetState("LoggedIn");
            foreach (var id in Feature8) Assert.IsTrue(controls.Get(id).Enabled, id);
            Assert.IsFalse(controls.Get("RibEdgeLogin").Visible);
            Assert.IsTrue(controls.Get("RibEdgeLogout").Visible);
            Assert.IsFalse(controls.Get("RibEdgeIncludeOutputData").Pressed);
        }

        [TestMethod]
        public void NoXLEdgePermission_MatchesTodaysAddin()
        {
            var (c, controls, shell) = Create();
            controls.Get("RibEdgeDebug").Pressed = true;
            controls.Get("RibEdgeIncludeOutputData").Pressed = true;
            c.SetState("NoXLEdgePermission");
            foreach (var id in Feature8.Concat(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }))
                Assert.IsFalse(controls.Get(id).Enabled, id);
            Assert.IsTrue(controls.Get("RibEdgeLogout").Enabled, "a user without XLEdge permission can still log out");
            Assert.IsFalse(controls.Get("RibEdgeDebug").Pressed);
            Assert.IsFalse(controls.Get("RibEdgeIncludeOutputData").Pressed);
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
