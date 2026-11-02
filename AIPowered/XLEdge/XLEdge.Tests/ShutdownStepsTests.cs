using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XLEdge.Tests.CoreWindowsTests;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 4: the core's Shutdown runs every step even when one fails - one pane's WebView2 failing to dispose must not
    /// leave the other panes, the WPF dispatcher or the Excel reference alive into the unload.
    /// </summary>
    [TestClass]
    public class ShutdownStepsTests
    {
        private static string TestBin => Path.GetDirectoryName(typeof(ShutdownStepsTests).Assembly.Location);
        private static string Core => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\XLEdge.Addin.Core"));

        [TestMethod]
        public void RunAll_AStepThrows_LaterStepsStillRun()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                {
                    FakeContext context = NewContext(t);
                    CoreOnThisThread(context, p =>
                    {
                        Assert.AreEqual("a,c", p.RunShutdownStepsWithAFailingStep());
                        var log = (RecordingLogger)context.Logger;
                        Assert.IsTrue(log.Contains("AddinEntry.Shutdown (b)"), log.Dump());
                    });
                }
            });
        }

        [TestMethod]
        public void Shutdown_UsesTheStepRunner()
        {
            // A real pane cannot be made to fail its Dispose without a production test seam: the wiring is checked in the source.
            string entry = File.ReadAllText(Path.Combine(Core, "AddinEntry.cs")).Replace("\r\n", "\n");
            string shutdown = CorePortParityTests.MethodText(entry, "void Shutdown()");
            Assert.IsNotNull(shutdown, "AddinEntry.Shutdown");
            StringAssert.Contains(shutdown, "ShutdownSteps.RunAll(\"AddinEntry.Shutdown\"");
            int last = -1;
            foreach (string step in new[] { "\"session options\"", "\"session login\"", "\"module\"", "\"panes\"", "\"wpf\"", "\"core services\"" })
            {
                int at = shutdown.IndexOf(step, StringComparison.Ordinal);
                Assert.IsTrue(at > last, "step " + step + " in order");
                last = at;
            }
            string host = File.ReadAllText(Path.Combine(Core, @"Pane\TaskPaneContentHost.cs")).Replace("\r\n", "\n");
            string closeAll = CorePortParityTests.MethodText(host, "void CloseAll()");
            Assert.IsNotNull(closeAll, "TaskPaneContentHost.CloseAll");
            StringAssert.Contains(closeAll, "catch (Exception", "each pane closes in its own try");
        }
    }
}
