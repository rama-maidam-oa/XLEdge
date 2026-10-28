using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Windows.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;
using XLEdge.Views;

namespace XLEdge.Tests
{
    /// <summary>The shell's WPF windows construct, show the XLEdge texts and behave as the reload flow needs (run on an STA thread).</summary>
    [TestClass]
    public class HostWindowsTests
    {
        private static void OnSta(Action test)
        {
            ExceptionDispatchInfo failure = null;
            var thread = new Thread(() =>
            {
                try { test(); }
                catch (Exception ex) { failure = ExceptionDispatchInfo.Capture(ex); }
                finally { Dispatcher.CurrentDispatcher.InvokeShutdown(); }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            failure?.Throw();
        }

        /// <summary>Every piece of text in the window (TextBlocks, labels, buttons, tooltips are not included).</summary>
        private static List<string> Texts(System.Windows.DependencyObject root)
        {
            var texts = new List<string>();
            void Walk(object node)
            {
                if (node is System.Windows.Controls.TextBlock tb) texts.Add(string.Concat(tb.Inlines.OfType<System.Windows.Documents.Run>().Select(r => r.Text)) is string runs && runs.Length > 0 ? runs : tb.Text);
                if (node is System.Windows.Controls.ContentControl cc && cc.Content is string s) texts.Add(s);
                if (node is System.Windows.DependencyObject d)
                    foreach (object child in System.Windows.LogicalTreeHelper.GetChildren(d)) Walk(child);
            }
            Walk(root);
            return texts;
        }

        [TestMethod]
        public void Windows_ShowTheXLEdgeTexts()
        {
            using (var t = new TempFolder())
            {
                var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data")));
                OnSta(() =>
                {
                    var progress = Texts(new XLEdgeReloadProgressWindow("1", "d"));
                    CollectionAssert.IsSubsetOf(new[] { "Reloading XLEdge", "Please keep Excel open while the new version is loaded.",
                        "This window closes automatically when reload is complete.", "Closing current", "Finding version", "Loading", "Done" }, progress, string.Join(" | ", progress));

                    var picker = Texts(new XLEdgeReloadSourcePicker(context));
                    CollectionAssert.IsSubsetOf(new[] { "Reload XLEdge", "Choose a source to safely load an available version.",
                        "Reloading safely restarts the XLEdge add-in without restarting Excel." }, picker, string.Join(" | ", picker));

                    var browser = Texts(new XLEdgeReleaseHistoryBrowser(context));
                    CollectionAssert.IsSubsetOf(new[] { "XLEdge Version History", "Versions adopted on this machine.",
                        "Loading a version safely restarts the XLEdge add-in without restarting Excel." }, browser, string.Join(" | ", browser));
                });
            }
        }

        [TestMethod]
        public void ProgressWindow_ShowsStagesAndOnlyClosesWhenAllowed()
        {
            OnSta(() =>
            {
                var window = new XLEdgeReloadProgressWindow("11.1.2.0", "2026-10-28T10:00:00");
                Assert.AreEqual("Reload XLEdge Add-in", window.Title);
                Assert.IsTrue(window.Topmost);
                window.ShowAndRender();
                window.ShowStage("Closing the current XLEdge session...", 1);
                Assert.AreEqual("Closing the current XLEdge session...", window.CurrentMessage);
                Assert.AreEqual(1, window.CurrentStep);
                window.SetTarget("11.1.3.0", "2026-10-29T10:00:00");
                window.ShowStage("XLEdge reload complete.", 4);
                Assert.AreEqual(4, window.CurrentStep);

                window.Close();
                Assert.IsTrue(window.IsVisible, "the user cannot close it while the reload runs");
                window.AllowCloseAndClose();
                Assert.IsFalse(window.IsVisible);
            });
        }

        [TestMethod]
        public void SourcePicker_OfflineOnly_ForNow()
        {
            using (var t = new TempFolder())
            {
                var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data")));
                OnSta(() =>
                {
                    var picker = new XLEdgeReloadSourcePicker(context);
                    Assert.AreEqual("Reload XLEdge Add-in", picker.Title);
                    Assert.IsFalse(picker.IsOnlineEnabled, "online updates arrive with the orbitbootapp endpoints (Stage 6)");
                    Assert.IsTrue(picker.IsOfflineSelected);
                    Assert.IsNull(picker.SelectedSource);
                    Assert.IsNull(picker.ScanResult, "the picker only verifies; the module stages after the reload is confirmed");
                    picker.Close();
                });
            }
        }

        [TestMethod]
        public void ReleaseHistoryBrowser_ListsCataloguedReleasesNewestFirst()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                foreach (var (folder, date) in new[] { ("vOld", "2026-10-01T10:00:00"), ("vNew", "2026-10-20T10:00:00") })
                {
                    System.IO.Directory.CreateDirectory(System.IO.Path.Combine(paths.VersionsPath, folder));
                    System.IO.File.WriteAllText(System.IO.Path.Combine(paths.VersionsPath, folder, "XLEdge.Addin.Core.dll"), "x");
                    ReleaseHistoryStore.Append(paths.ReleaseHistoryFile, new ReleaseEntry { Version = folder, ReleaseDate = date, FolderName = folder, ContractVersion = 1 });
                }
                var context = new FakeContext(paths) { ActiveFolderName = "vOld" };
                OnSta(() =>
                {
                    var browser = new XLEdgeReleaseHistoryBrowser(context);
                    Assert.AreEqual("XLEdge Version History", browser.Title);
                    CollectionAssert.AreEqual(new[] { "vNew", "vOld" }, browser.Rows.Select(r => r.FolderName).ToArray());
                    Assert.IsTrue(browser.Rows.Single(r => r.FolderName == "vOld").IsCurrentlyLoaded);
                    Assert.IsNull(browser.Chosen);
                    browser.Close();
                });
            }
        }

        [TestMethod]
        public void WpfHostUi_ProgressAdapterDrivesTheWindow()
        {
            OnSta(() =>
            {
                // The owner handle is read on every window; a read that throws shows the window unowned.
                var ui = new XLEdge.Hosting.WpfHostUi(() => throw new System.Runtime.InteropServices.COMException("Excel is busy"));
                var progress = ui.ShowProgress("11.1.2.0", "2026-10-28T10:00:00");
                Assert.IsNotNull(progress);
                progress.ShowStage("Loading the selected XLEdge version...", 3);
                progress.SetTarget("11.1.3.0", "2026-10-29T10:00:00");
                progress.Close();
            });
        }
    }
}
