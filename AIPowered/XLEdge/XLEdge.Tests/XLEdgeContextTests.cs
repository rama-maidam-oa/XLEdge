using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class XLEdgeContextTests
    {
        [TestMethod]
        public void Context_ShellVersionLeaseAndDefaults()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                string expected = FileVersionInfo.GetVersionInfo(typeof(XLEdgeContext).Assembly.Location).FileVersion;
                Assert.AreEqual(expected, context.ShellVersion);
                Assert.AreEqual("11.1.2.0", context.ShellVersion);
                Assert.IsNull(context.InitializeLifetimeService());
                Assert.IsFalse(context.DebugMode);
                Assert.IsTrue(Directory.Exists(context.Paths.Logs), "paths ensured");
                Assert.IsNotNull(context.Logger);
                Assert.IsNull(context.RibbonController);
            }
        }

        [TestMethod]
        public void Context_WithoutGLSense_ReturnsNullAndLogs()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                Assert.IsNull(context.GetGLSenseAddinInstance());
            }
        }

        public sealed class FakeComAddin
        {
            public string ProgId { get; set; } = "GLSense.AddinModule";
            public bool Connect { get; set; } = true;
            public object Object { get; set; } = new object();
        }

        public sealed class FakeComAddins
        {
            public FakeComAddin Only { get; } = new FakeComAddin();
            public int Count => 1;
            public object Item(object index) => Only;
        }

        public sealed class FakeExcel
        {
            public FakeComAddins COMAddIns { get; } = new FakeComAddins();
        }

        [TestMethod]
        public void Context_GLSense_IsLookedUpEveryTime()
        {
            // GLSense can be disconnected or reloaded during the session: a cached object would go stale.
            using (var t = new TempFolder())
            {
                var excel = new FakeExcel();
                using (var context = new XLEdgeContext(excel, new PathProvider(t.Combine("inst"), t.Combine("data"))))
                {
                    object first = excel.COMAddIns.Only.Object;
                    Assert.AreSame(first, context.GetGLSenseAddinInstance());
                    excel.COMAddIns.Only.Connect = false;
                    Assert.IsNull(context.GetGLSenseAddinInstance());
                    excel.COMAddIns.Only.Connect = true;
                    excel.COMAddIns.Only.Object = new object();
                    Assert.AreSame(excel.COMAddIns.Only.Object, context.GetGLSenseAddinInstance());
                }
            }
        }

        [TestMethod]
        public void Context_Dispose_ReleasesTheLogFile()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                var context = new XLEdgeContext(new object(), paths);
                context.Logger.LogInfo("hello");
                context.Dispose();
                foreach (var file in Directory.GetFiles(paths.Logs)) File.Delete(file);
            }
        }

        [TestMethod]
        public void RibbonControlIds_StateListsMatchTodaysAddin()
        {
            CollectionAssert.AreEquivalent(new[] { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide", "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp" },
                RibbonControlIds.LoggedOutDisabled.ToArray());
            CollectionAssert.AreEquivalent(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }, RibbonControlIds.LoggedOutEnabled.ToArray());
            CollectionAssert.AreEquivalent(new[] { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide",
                "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp", "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" },
                RibbonControlIds.NoPermissionDisabled.ToArray());
        }
    }
}
