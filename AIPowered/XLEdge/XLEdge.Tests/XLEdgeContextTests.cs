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
            Assert.AreEqual(12, RibbonControlIds.AllFeatureControls.Count);
            CollectionAssert.DoesNotContain(RibbonControlIds.AllFeatureControls.ToArray(), RibbonControlIds.RibReload);
            CollectionAssert.DoesNotContain(RibbonControlIds.AllFeatureControls.ToArray(), RibbonControlIds.RibVersionHistory);
        }
    }
}
