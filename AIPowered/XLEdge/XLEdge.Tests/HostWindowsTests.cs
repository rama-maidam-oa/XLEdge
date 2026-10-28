using System;
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
                var ui = new XLEdge.Hosting.WpfHostUi(IntPtr.Zero);
                var progress = ui.ShowProgress("11.1.2.0", "2026-10-28T10:00:00");
                Assert.IsNotNull(progress);
                progress.ShowStage("Loading the selected XLEdge version...", 3);
                progress.SetTarget("11.1.3.0", "2026-10-29T10:00:00");
                progress.Close();
            });
        }
    }
}
