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

        /// <summary>WPF-UI's theme and XLEdge's GlobalStyles really are merged into the Application (WpfUiBootstrapper only logs failures).</summary>
        private static void AssertWpfUiResourcesLoaded(CoreProbe probe, FakeContext context)
        {
            string[] sources = probe.ResourceDictionarySources();
            foreach (string expected in new[] { "Wpf.Ui;component/Resources/Theme/Light.xaml", "Wpf.Ui;component/Resources/Wpf.Ui.xaml",
                                                "XLEdge.Addin.Core;component/Themes/GlobalStyles.xaml" })
                Assert.IsTrue(sources.Any(s => s.EndsWith(expected, StringComparison.OrdinalIgnoreCase)), expected + " not merged: " + string.Join(" | ", sources));
            var log = (RecordingLogger)context.Logger;
            Assert.IsFalse(log.Lines.Any(l => (l.StartsWith("ERROR") || l.StartsWith("EXCEPTION") || l.StartsWith("WARN")) && l.Contains("WpfUi")), log.Dump());
        }

        private static FakeContext NewContext(TempFolder t)
        {
            var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
            paths.Ensure();
            var context = new FakeContext(paths) { Logger = new RecordingLogger() };
            context.SetRibbonController(new CoreRibbonFacadeTests.FakeRibbon());
            return context;
        }

        /// <summary>One core on this STA thread: Initialize, <paramref name="use"/>, Shutdown, then an unload as the loader does it.</summary>
        private static void CoreOnThisThread(FakeContext context, Action<CoreProbe> use)
        {
            var (domain, probe) = CoreProbe.Create("CoreWindowsSession");
            try
            {
                probe.InitializeTheEntry(context);
                try { use(probe); }
                finally
                {
                    // Always: a failing check must not leave the core's windows alive for the unload below (their messages would
                    // then crash the pump with AppDomainUnloadedException and hide the real failure).
                    probe.ShutDownTheEntry();
                }
            }
            finally
            {
                var unload = System.Threading.Tasks.Task.Run(() => AppDomain.Unload(domain));
                Win32Test.PumpUntil(() => unload.IsCompleted, TimeSpan.FromSeconds(10));
                unload.Wait();
            }
        }

        [TestMethod]
        public void Reload_TheNextCoreOnTheSameThread_OpensStyledWindows()
        {
            // Hot reload: the first core's WPF is shut down and unloaded, then the next core creates its own Application on the
            // same (Excel's) thread and must open windows with WPF-UI and GlobalStyles resolved again.
            Win32Test.OnSta(() =>
            {
                Win32Test.Pump(TimeSpan.FromMilliseconds(50));
                var before = Win32Test.ProcessWindows();
                using (var t = new TempFolder())
                {
                    FakeContext context = NewContext(t);
                    CoreOnThisThread(context, p => StringAssert.StartsWith(p.OpenWindow("XLEdge.Views.XLEdgeOptions"), "loaded=True"));
                    CoreOnThisThread(context, p =>
                    {
                        Assert.AreEqual(true, p.CallStatic("XLEdge.Utilities.WpfUiBootstrapper", "IsInitialized"), "the reloaded core sets WPF-UI up again");
                        StringAssert.StartsWith(p.OpenWindow("XLEdge.Views.XLEdgeOptions"), "loaded=True");
                        StringAssert.StartsWith(p.OpenWindow("XLEdge.Views.XLEdgeWaitWindow", new object[] { null }), "loaded=True");
                        AssertWpfUiResourcesLoaded(p, context);
                    });
                }
                Win32Test.Pump(TimeSpan.FromSeconds(1));
                var left = Win32Test.ProcessWindows().Where(w => !before.ContainsKey(w.Key) && Managed(w.Value)).Select(w => w.Value).ToList();
                Assert.AreEqual(0, left.Count, "managed windows left after two cores: " + string.Join(", ", left));
            });
        }

        [TestMethod]
        public void IsBusy_WhileACoreDialogIsModal()
        {
            // A reload unloads the core's AppDomain; with a core dialog's modal loop on Excel's thread that would abort the thread.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                {
                    CoreOnThisThread(NewContext(t), p =>
                    {
                        Assert.IsFalse(p.EntryIsBusy(), "idle");
                        Assert.IsTrue(p.IsBusyInsideAModalDialog(), "a modal core dialog makes the core busy");
                        Assert.IsFalse(p.EntryIsBusy(), "idle again once it closed");
                    });
                }
            });
        }

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

                        AssertWpfUiResourcesLoaded(probe, context);

                        probe.ShutDownTheEntry();
                        CollectionAssert.AreEqual(new string[0], probe.WindowsNotClosed(), "every window really closes (Closed raised), the wait window too");
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
