using System;
using XLEdge.Contracts;
using System.Threading.Tasks;
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
        public void SourcePicker_LoggedOut_OnlineDisabledWithTheReason()
        {
            using (var t = new TempFolder())
            {
                var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data")));
                OnSta(() =>
                {
                    foreach (LoginInfo login in new[] { null, new LoginInfo(), new LoginInfo { LoginUrl = "https://h/reporting", LoginToken = "t", IsLoggedIn = false } })
                    {
                        var picker = new XLEdgeReloadSourcePicker(context, login);
                        Assert.AreEqual("Reload XLEdge Add-in", picker.Title);
                        Assert.IsFalse(picker.IsOnlineEnabled);
                        Assert.AreEqual("Log in to XLEdge to check for online updates.", picker.OnlineToolTip);
                        Assert.IsTrue(picker.IsOfflineSelected);
                        Assert.IsNull(picker.SelectedSource);
                        Assert.IsNull(picker.ScanResult, "the picker only verifies; the module stages after the reload is confirmed");
                        picker.Close();
                    }
                });
            }
        }

        // ---- Stage 6: Online ----

        /// <summary>The server as the picker sees it: a release list, and per folder a download (or a failure).</summary>
        internal sealed class FakeOnlineSource : XLEdge.Online.IOnlineReleaseSource
        {
            public List<VersionInfo> Releases = new List<VersionInfo>();
            public readonly Dictionary<string, (string Manifest, byte[] Zip)> Downloads = new Dictionary<string, (string, byte[])>();
            public Exception ListFailure;
            public readonly List<string> Downloaded = new List<string>();
            public Task<IReadOnlyList<VersionInfo>> GetReleasesAsync(CancellationToken ct) =>
                ListFailure != null ? Task.FromException<IReadOnlyList<VersionInfo>>(ListFailure) : Task.FromResult<IReadOnlyList<VersionInfo>>(Releases);
            public Task<(string ManifestJson, byte[] Zip)> DownloadAsync(string folderName, CancellationToken ct)
            {
                Downloaded.Add(folderName);
                return Task.FromResult(Downloads[folderName]);
            }
            public void Dispose() { }

            /// <summary>A release the server publishes: a core zip and its manifest (checksum right, or <paramref name="badChecksum"/>).</summary>
            public void Publish(string version, string date, string folder, bool badChecksum = false)
            {
                byte[] zip = ReleaseFolders.FakeZip(("XLEdge.Addin.Core.dll", folder), ("XLEdge.Contracts.dll", "contracts"));
                string checksum = badChecksum ? new string('0', 64) : ReleaseFolders.Sha256(zip);
                Releases.Add(new VersionInfo { Version = version, ReleaseDate = date, FolderName = folder, FileName = folder + ".zip", Checksum = checksum, ContractVersion = 1 });
                Downloads[folder] = (ReleaseFolders.Manifest(version, date, folder, checksum), zip);
            }
        }

        private static readonly LoginInfo LoggedIn = new LoginInfo { LoginUrl = "https://h/reporting", LoginToken = "t", LoginUserName = "u", IsLoggedIn = true };

        /// <summary>Runs <paramref name="work"/> on the picker's thread and pumps until it ends (the picker awaits).</summary>
        private static void Await(Func<Task> work)
        {
            Task task = work();
            Assert.IsTrue(Win32Test.PumpUntil(() => task.IsCompleted, TimeSpan.FromSeconds(30)), "the picker finished");
            task.GetAwaiter().GetResult();
        }

        private static void WithOnlinePicker(Action<XLEdgeReloadSourcePicker, FakeOnlineSource, FakeContext> use)
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                var context = new FakeContext(paths) { Logger = new RecordingLogger(), Version = "11.1.3.0", ReleaseDate = "2026-11-03T10:00:00", ShellVersion = "11.1.3.0" };
                var source = new FakeOnlineSource();
                OnSta(() =>
                {
                    var picker = new XLEdgeReloadSourcePicker(context, LoggedIn, _ => source);
                    try { use(picker, source, context); }
                    finally { picker.Close(); }
                });
            }
        }

        [TestMethod]
        public void SourcePicker_LoggedIn_OnlineEnabled()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                Assert.IsTrue(picker.IsOnlineEnabled);
                picker.SelectOnline();
                Assert.IsFalse(picker.IsOfflineSelected);
                Assert.IsFalse(picker.IsReloadEnabled, "nothing checked yet");
            });
        }

        [TestMethod]
        public void SourcePicker_Check_ListsClassifiedRows()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.Publish("11.1.3.0", "2026-11-03T10:00:00", "vLoaded");
                source.Publish("11.1.4.0", "2026-11-10T10:00:00", "vNew");
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                CollectionAssert.AreEqual(new[] { "vNew", "vLoaded" }, picker.OnlineRows.Select(r => r.FolderName).ToArray());
                Assert.IsTrue(picker.OnlineRows[0].IsChecked, "the newest new release is ticked");
                Assert.IsFalse(picker.OnlineRows[1].IsSelectable, "the loaded release cannot be chosen");
                Assert.IsTrue(picker.IsReloadEnabled);
            });
        }

        [TestMethod]
        public void SourcePicker_ReloadOnline_DownloadsEveryTickedRelease_LoadsTheNewest()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.Publish("11.1.4.0", "2026-11-10T10:00:00", "v4");
                source.Publish("11.1.5.0", "2026-11-20T10:00:00", "v5");
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                foreach (var row in picker.OnlineRows) row.IsChecked = true;
                Await(() => picker.ReloadOnlineAsync());
                CollectionAssert.AreEqual(new[] { "v4", "v5" }, source.Downloaded, "oldest first");
                Assert.AreEqual("v5", picker.OnlineRelease?.FolderName, "the newest loads");
                Assert.AreEqual("Online", picker.SelectedSource);
                foreach (string folder in new[] { "v4", "v5" })
                    Assert.IsTrue(System.IO.File.Exists(System.IO.Path.Combine(context.Paths.VersionsPath, folder, "XLEdge.Addin.Core.dll")), folder + " catalogued");
            });
        }

        [TestMethod]
        public void SourcePicker_ReloadOnline_OneFails_KeepsTheOthers_AndSaysWhichFailed()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.Publish("11.1.4.0", "2026-11-10T10:00:00", "v4");
                source.Publish("11.1.5.0", "2026-11-20T10:00:00", "v5", badChecksum: true);
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                foreach (var row in picker.OnlineRows) row.IsChecked = true;
                Await(() => picker.ReloadOnlineAsync());
                Assert.AreEqual("v4", picker.OnlineRelease?.FolderName, "the newest release that was catalogued");
                StringAssert.Contains(picker.StatusText, "11.1.5.0", "the failed release is named");
                Assert.IsFalse(System.IO.Directory.Exists(System.IO.Path.Combine(context.Paths.VersionsPath, "v5")));
            });
        }

        [TestMethod]
        public void SourcePicker_ReloadOnline_AllFail_StaysOpenWithTheErrors()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.Publish("11.1.5.0", "2026-11-20T10:00:00", "v5", badChecksum: true);
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                Await(() => picker.ReloadOnlineAsync());
                Assert.IsNull(picker.OnlineRelease);
                Assert.IsNull(picker.SelectedSource);
                Assert.IsTrue(picker.IsVisibleOrOpen, "not closed");
                StringAssert.Contains(picker.StatusText, "11.1.5.0");
            });
        }

        [TestMethod]
        public void SourcePicker_ReloadOnline_NothingTicked_ReloadDisabled()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.Publish("11.1.4.0", "2026-11-10T10:00:00", "v4");
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                picker.OnlineRows[0].IsChecked = false;
                Assert.IsFalse(picker.IsReloadEnabled);
            });
        }

        [TestMethod]
        public void SourcePicker_Check_Unauthorized_ShowsTheSessionMessage()
        {
            WithOnlinePicker((picker, source, context) =>
            {
                source.ListFailure = new XLEdge.Online.OnlineReleaseException(XLEdge.Online.OnlineReleaseClient.SessionExpired, 401);
                picker.SelectOnline();
                Await(() => picker.CheckForUpdatesAsync());
                StringAssert.Contains(picker.StatusText, "Your XLEdge session has expired. Log in again, then check for updates again.");
                Assert.AreEqual(0, picker.OnlineRows.Count);
            });
        }

        [TestMethod]
        public void ShellModule_ReloadsIntoTheOnlineRelease()
        {
            // ShowReloadWindow is shell ADX code (checked in Excel): the wiring is checked in the source.
            string module = System.IO.File.ReadAllText(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(typeof(HostWindowsTests).Assembly.Location), @"..\..\..\XLEdge\AddinModule.cs"));
            string show = CorePortParityTests.MethodText(module.Replace("\r\n", "\n"), "void ShowReloadWindow()");
            Assert.IsNotNull(show);
            StringAssert.Contains(show, "GetLoginInfo()");
            StringAssert.Contains(show, "new XLEdgeReloadSourcePicker(_context, login)");
            StringAssert.Contains(show, "Host.Reload(() => online)");
            int ask = show.IndexOf("GetLoginInfo()", StringComparison.Ordinal);
            Assert.IsTrue(show.LastIndexOf("try", ask, StringComparison.Ordinal) >= 0, "an older core may not answer");
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
