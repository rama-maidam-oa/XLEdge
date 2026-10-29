using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// XLEdge 11.1.2's windows open from the core's AppDomain: their XAML, the GlobalStyles theme, WPF-UI and the images resolve
    /// there (pack URIs name XLEdge.Addin.Core), the shell can list them before a reload, and nothing of them survives the unload.
    /// </summary>
    [TestClass]
    public class CoreWindowsTests
    {
        private static bool Managed(string cls) => cls.StartsWith("HwndWrapper", StringComparison.Ordinal) ||
                                                   cls.StartsWith("WindowsForms10.", StringComparison.Ordinal) || cls.StartsWith(".NET-", StringComparison.Ordinal);

        [TestMethod]
        public void Windows_OpenFromTheCore_AreListed_AndNothingSurvivesTheUnload()
        {
            Win32Test.OnSta(() =>
            {
                // Pumping creates WinForms' hidden window in this test's own domain: before the snapshot.
                Win32Test.Pump(TimeSpan.FromMilliseconds(50));
                var before = Win32Test.ProcessWindows();
                List<string> Added() => Win32Test.ProcessWindows().Where(w => !before.ContainsKey(w.Key) && Managed(w.Value)).Select(w => w.Value).ToList();

                using (var t = new TempFolder())
                {
                    var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                    paths.Ensure();
                    var context = new FakeContext(paths) { Logger = new RecordingLogger() };
                    context.SetRibbonController(new CoreRibbonFacadeTests.FakeRibbon());
                    var (domain, probe) = CoreProbe.Create("CoreWindows");
                    try
                    {
                        probe.InitializeTheEntry(context);
                        Assert.AreEqual(true, probe.CallStatic("XLEdge.Utilities.WpfUiBootstrapper", "IsInitialized"),
                            "WPF-UI is set up at start-up, as 11.1.2 did on ribbon load");

                        var opened = new Dictionary<string, string>
                        {
                            ["XLEdge.Views.XLEdgeCalendar"] = probe.OpenWindow("XLEdge.Views.XLEdgeCalendar", new DateTime(2026, 10, 29)),
                            ["XLEdge.Views.XLEdgeDrilldownReports"] = probe.OpenWindow("XLEdge.Views.XLEdgeDrilldownReports"),
                            ["XLEdge.Views.XLEdgeLoginDetails"] = probe.OpenWindow("XLEdge.Views.XLEdgeLoginDetails"),
                            ["XLEdge.Views.XLEdgeMessageWindow"] = probe.OpenWindow("XLEdge.Views.XLEdgeMessageWindow", "a message",
                                System.Windows.Forms.MessageBoxIcon.Information, System.Windows.Forms.MessageBoxButtons.OK),
                            ["XLEdge.Views.XLEdgeOptions"] = probe.OpenWindow("XLEdge.Views.XLEdgeOptions"),
                            ["XLEdge.Views.XLEdgeWaitWindow"] = probe.OpenWindow("XLEdge.Views.XLEdgeWaitWindow", new object[] { null }),
                            ["XLEdge.Views.AppOverlay"] = probe.OpenWindow("XLEdge.Views.AppOverlay"),
                        };
                        foreach (var kv in opened)
                            StringAssert.StartsWith(kv.Value, "loaded=True", kv.Key + ": " + kv.Value);

                        string[] titles = probe.EntryOpenWindowTitles();
                        Assert.AreEqual(opened.Count, titles.Length, "every open core window is listed: " + string.Join(" | ", titles));
                        CollectionAssert.Contains(titles, "Select Date");
                        CollectionAssert.Contains(titles, "Drill Down Reports");
                        Assert.IsTrue(Added().Any(c => c.StartsWith("HwndWrapper", StringComparison.Ordinal)), "precondition: the core's windows exist");

                        probe.ShutDownTheEntry();
                        Win32Test.PumpUntil(() => !Added().Any(), TimeSpan.FromSeconds(5));
                        Assert.AreEqual(0, Added().Count, "AddinEntry.Shutdown closes the core's windows: " + string.Join(", ", Added()));
                    }
                    finally
                    {
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
