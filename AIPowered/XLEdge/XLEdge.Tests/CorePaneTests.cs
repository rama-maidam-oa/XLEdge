using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Hosting;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// The Stage 1 spike, automated: the built core, in its own AppDomain, hosts a WebView2 per Excel window in a window the
    /// shell re-parents into its pane (a WinForms form stands in for the ADX pane). Needs the WebView2 Runtime (Edge).
    /// </summary>
    [TestClass]
    public class CorePaneTests
    {
        private static readonly TimeSpan PageTimeout = TimeSpan.FromSeconds(30);

        /// <summary>The built package, installed into a temp profile and loaded; disposed = core shut down and unloaded.</summary>
        internal sealed class LoadedCore : IDisposable
        {
            private readonly AddinDomainLoader _loader = new AddinDomainLoader();
            private bool _unloaded;

            public LoadedCore(TempFolder temp, string profile = "data")
            {
                var paths = new PathProvider(temp.Combine("inst"), temp.Combine(profile));
                paths.Ensure();
                Log = new RecordingLogger();
                Context = new FakeContext(paths) { Logger = Log };
                string folder = CorePackageTests.PackageFolder;
                Release = new UpdateBootstrapper().ExtractAndCatalog(Context,
                    File.ReadAllText(Path.Combine(folder, "manifest.json")), File.ReadAllBytes(Directory.GetFiles(folder, "*.zip").Single()), "Install");
                Assert.IsNotNull(Release, Log.Dump());
                Context.Version = Release.Version;
                Addin = _loader.Load(Context, Release);
            }

            public RecordingLogger Log { get; }
            public FakeContext Context { get; }
            public ResolvedRelease Release { get; }
            public IXLEdgeAddin Addin { get; }

            public bool WaitForLog(string text, TimeSpan timeout) => Win32Test.PumpUntil(() => Log.Contains(text), timeout);

            public void Unload()
            {
                if (_unloaded) return;
                _unloaded = true;
                try { Addin.Shutdown(); }
                finally { _loader.Unload(Context); }
            }

            public void Dispose() => Unload();
        }

        private static Form NewPane()
        {
            var pane = new Form { Width = 420, Height = 640, Text = "pane" };
            pane.Show();
            return pane;
        }

        [TestMethod]
        public void CreateTaskPaneContent_ReturnsAHiddenToolWindow()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var excel = NewPane())
                {
                    excel.Activate();
                    IntPtr active = Win32Test.GetActiveWindow();
                    IntPtr content = core.Addin.CreateTaskPaneContent(excel.Handle);
                    // The thread's active window shows an activation even when this test process is not in the foreground.
                    Assert.AreEqual(active, Win32Test.GetActiveWindow(), "showing the content must not activate it");
                    Assert.IsTrue(Win32Test.IsWindow(content), core.Log.Dump());
                    long ex = Win32Test.ExStyle(content);
                    Assert.AreNotEqual(0L, ex & Win32Test.WS_EX_TOOLWINDOW, "tool window: never in the taskbar or Alt-Tab");
                    Assert.AreEqual(0L, ex & Win32Test.WS_EX_APPWINDOW);
                    var (left, top) = Win32Test.ScreenPosition(content);
                    Assert.IsTrue(left <= -30000 && top <= -30000, $"off-screen until embedded, was {left},{top}");
                }
            });
        }

        [TestMethod]
        public void CreateTaskPaneContent_IsPerExcelWindowAndIdempotent()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var excel1 = NewPane())
                using (var excel2 = NewPane())
                {
                    IntPtr first = core.Addin.CreateTaskPaneContent(excel1.Handle);
                    Assert.AreEqual(first, core.Addin.CreateTaskPaneContent(excel1.Handle), "same Excel window, same content");
                    IntPtr second = core.Addin.CreateTaskPaneContent(excel2.Handle);
                    Assert.AreNotEqual(first, second);
                    Assert.IsTrue(Win32Test.IsWindow(second));
                }
            });
        }

        [TestMethod]
        public void EmbeddedContent_LoadsTheSpikePage_AndTheTitleBusWorks()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var pane = NewPane())
                {
                    IntPtr content = core.Addin.CreateTaskPaneContent(pane.Handle);
                    Assert.IsTrue(PaneEmbedding.Embed(content, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height));
                    core.Addin.OnTaskPaneShown(pane.Handle);
                    // The page sends "ready" while it loads, so the command can arrive before NavigationCompleted: wait for both.
                    Assert.IsTrue(core.WaitForLog("PaneContent: command 'ready'", PageTimeout), core.Log.Dump());
                    Assert.IsTrue(core.WaitForLog("PaneContent: navigation completed (success=True", PageTimeout), core.Log.Dump());
                    Assert.IsFalse(core.Log.Lines.Any(l => l.StartsWith("ERROR") || l.StartsWith("EXCEPTION")), core.Log.Dump());
                }
            });
        }

        /// <summary>What the shell does before an unload: take the content out of the pane, close it, shut the core down, unload.</summary>
        private static void DetachAndUnload(LoadedCore core, Form pane, IntPtr content)
        {
            PaneEmbedding.Release(content);
            core.Addin.CloseTaskPaneContent(pane.Handle);
            core.Unload();
        }

        [TestMethod]
        public void Unload_WithEmbeddedContent_LeavesNoCoreWindowInTheProcess()
        {
            // A window whose window procedure lives in the unloaded AppDomain (a form, the WinForms parking window or
            // marshaling control, a SystemEvents window) crashes Excel on the next message it receives.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var pane = NewPane())
                {
                    var before = Win32Test.ProcessWindows();
                    var core = new LoadedCore(t);
                    IntPtr content = core.Addin.CreateTaskPaneContent(pane.Handle);
                    PaneEmbedding.Embed(content, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height);
                    Assert.IsTrue(core.WaitForLog("PaneContent: command 'ready'", PageTimeout), core.Log.Dump());
                    DetachAndUnload(core, pane, content);
                    AssertNothingOfTheCoreLeft(before, core);
                }
            });
        }

        [TestMethod]
        public void Unload_WhileTheWebViewIsStarting_LeavesNothing()
        {
            // Login then Reload (or Excel closing) at once: the environment / controller creation is still pending, so its
            // continuations are queued for the domain being unloaded.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var pane = NewPane())
                {
                    var before = Win32Test.ProcessWindows();
                    var core = new LoadedCore(t);
                    IntPtr content = core.Addin.CreateTaskPaneContent(pane.Handle);
                    PaneEmbedding.Embed(content, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height);
                    DetachAndUnload(core, pane, content);
                    AssertNothingOfTheCoreLeft(before, core);
                    Assert.IsFalse(core.Log.Contains("the WebView2 could not be started"), "closing during start-up is not a failure: " + core.Log.Dump());
                }
            });
        }

        /// <summary>
        /// After an unload: no window of a managed class (WinForms, WPF, SystemEvents) and no browser process of the core is
        /// left, and pumping messages is safe. Native helper windows may stay (COM's OleMainThreadWndClass, GDI+'s hook window,
        /// IME): their window procedures are in ole32/gdiplus/imm32, not in the unloaded domain.
        /// </summary>
        private static void AssertNothingOfTheCoreLeft(System.Collections.Generic.Dictionary<IntPtr, string> before, LoadedCore core)
        {
            bool Managed(string cls) => cls.StartsWith("WindowsForms10.", StringComparison.Ordinal)
                || cls.StartsWith("HwndWrapper", StringComparison.Ordinal) || cls.StartsWith(".NET-", StringComparison.Ordinal);
            Win32Test.PumpUntil(() => !Win32Test.ProcessWindows().Any(w => !before.ContainsKey(w.Key) && Managed(w.Value)), TimeSpan.FromSeconds(5));
            var added = Win32Test.ProcessWindows().Where(w => !before.ContainsKey(w.Key)).ToList();
            var left = added.Where(w => Managed(w.Value)).Select(w => $"{w.Key}:{w.Value} on {Win32Test.Owner(w.Key)}").ToList();
            Assert.AreEqual(0, left.Count, "managed windows left after the unload: " + string.Join(", ", left));
            Console.WriteLine("native windows added by the session: " + string.Join(", ", added.Select(w => w.Value)));
            Assert.IsTrue(Win32Test.PumpUntil(() => BrowserProcesses(core.Context.Paths.BrowserDataPath) == 0, TimeSpan.FromSeconds(15)),
                "a WebView2 browser process of the unloaded core is still running");
            Win32Test.Pump(TimeSpan.FromSeconds(3));
        }

        /// <summary>WebView2 renderer processes on this user data folder (one per live page).</summary>
        private static int Renderers(string userDataFolder)
        {
            using (var search = new System.Management.ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE Name = 'msedgewebview2.exe'"))
            using (var results = search.Get())
            {
                return results.Cast<System.Management.ManagementObject>().Select(p => (p["CommandLine"] as string) ?? "")
                    .Count(c => c.IndexOf(userDataFolder, StringComparison.OrdinalIgnoreCase) >= 0 && c.Contains("--type=renderer"));
            }
        }

        private static System.Collections.Generic.List<System.Diagnostics.Process> BrowserProcessList(string userDataFolder)
        {
            using (var search = new System.Management.ManagementObjectSearcher("SELECT ProcessId, CommandLine FROM Win32_Process WHERE Name = 'msedgewebview2.exe'"))
            using (var results = search.Get())
            {
                var list = new System.Collections.Generic.List<System.Diagnostics.Process>();
                foreach (var p in results.Cast<System.Management.ManagementObject>())
                {
                    if (((p["CommandLine"] as string) ?? "").IndexOf(userDataFolder, StringComparison.OrdinalIgnoreCase) < 0) continue;
                    try { list.Add(System.Diagnostics.Process.GetProcessById(Convert.ToInt32(p["ProcessId"]))); }
                    catch (ArgumentException) { /* already gone */ }
                }
                return list;
            }
        }

        [TestMethod]
        public void BrowserCrash_IsLogged_AndTheCoreStaysUsable()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var pane = NewPane())
                {
                    IntPtr content = core.Addin.CreateTaskPaneContent(pane.Handle);
                    PaneEmbedding.Embed(content, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height);
                    Assert.IsTrue(core.WaitForLog("PaneContent: command 'ready'", PageTimeout), core.Log.Dump());
                    foreach (var process in BrowserProcessList(core.Context.Paths.BrowserDataPath))
                    {
                        try { process.Kill(); } catch (Exception) { /* exited meanwhile */ }
                    }
                    Assert.IsTrue(core.WaitForLog("PaneContent: WebView2 process failed", PageTimeout), core.Log.Dump());
                    core.Addin.OnTaskPaneShown(pane.Handle);
                    Assert.AreEqual(content, core.Addin.CreateTaskPaneContent(pane.Handle), "the pane keeps its (empty) content; Excel is unaffected");
                }
            });
        }

        [TestMethod]
        public void CreateTaskPaneContent_AfterThePaneWasDestroyed_DisposesTheStaleContent()
        {
            // An Excel window closing destroys its pane, and with it the child content window; the next content for that key
            // must not leave the old form (and its WebView2) alive in the core.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                {
                    IntPtr key, stale;
                    string udf = core.Context.Paths.BrowserDataPath;
                    using (var pane = NewPane())
                    {
                        key = pane.Handle;
                        stale = core.Addin.CreateTaskPaneContent(key);
                        PaneEmbedding.Embed(stale, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height);
                        Assert.IsTrue(core.WaitForLog("PaneContent: command 'ready'", PageTimeout), core.Log.Dump());
                    }
                    Assert.IsFalse(Win32Test.IsWindow(stale), "destroying the pane destroys its child content");
                    int renderersBefore = Renderers(udf);
                    Assert.AreEqual(1, renderersBefore, "precondition: the stale content still has its renderer");

                    using (var pane = NewPane())
                    {
                        IntPtr fresh = core.Addin.CreateTaskPaneContent(key);
                        Assert.AreNotEqual(stale, fresh);
                        PaneEmbedding.Embed(fresh, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height);
                        Assert.IsTrue(Win32Test.PumpUntil(() => core.Log.Lines.Count(l => l.Contains("PaneContent: command 'ready'")) == 2, PageTimeout), core.Log.Dump());
                        // The stale form was disposed with its WebView2, so only the fresh page's renderer is left.
                        Assert.IsTrue(Win32Test.PumpUntil(() => Renderers(udf) == 1, TimeSpan.FromSeconds(15)), $"renderers: {Renderers(udf)}");
                    }
                }
            });
        }

        [TestMethod]
        public void CreateTaskPaneContent_ZeroExcelWindow_CreatesNothing()
        {
            // The shell could not read the Excel window: a shared Zero key would hand one pane's content to another.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                {
                    Assert.AreEqual(IntPtr.Zero, core.Addin.CreateTaskPaneContent(IntPtr.Zero));
                    core.Addin.OnTaskPaneShown(IntPtr.Zero);
                    core.Addin.CloseTaskPaneContent(IntPtr.Zero);
                    Assert.IsFalse(core.Log.Lines.Any(l => l.StartsWith("ERROR") || l.StartsWith("EXCEPTION")), core.Log.Dump());
                }
            });
        }

        /// <summary>msedgewebview2.exe processes started on this user data folder.</summary>
        private static int BrowserProcesses(string userDataFolder)
        {
            using (var search = new System.Management.ManagementObjectSearcher("SELECT CommandLine FROM Win32_Process WHERE Name = 'msedgewebview2.exe'"))
            using (var results = search.Get())
            {
                return results.Cast<System.Management.ManagementObject>()
                    .Count(p => ((p["CommandLine"] as string) ?? "").IndexOf(userDataFolder, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        [TestMethod]
        public void Shutdown_ClosesEveryContentWindow()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var excel1 = NewPane())
                using (var excel2 = NewPane())
                {
                    IntPtr first = core.Addin.CreateTaskPaneContent(excel1.Handle);
                    IntPtr second = core.Addin.CreateTaskPaneContent(excel2.Handle);
                    core.Addin.Shutdown();
                    Assert.IsFalse(Win32Test.IsWindow(first));
                    Assert.IsFalse(Win32Test.IsWindow(second));
                    Assert.AreEqual(IntPtr.Zero, core.Addin.CreateTaskPaneContent(excel1.Handle), "a shut-down core creates no more windows");
                }
            });
        }

        [TestMethod]
        public void Reload_SameUserDataFolder_SecondCoreShowsThePageAgain()
        {
            // Hot reload: the next core creates its own WebView2 environment on the same user data folder.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var pane = NewPane())
                {
                    foreach (string round in new[] { "first", "second" })
                    {
                        var core = new LoadedCore(t);
                        IntPtr content = core.Addin.CreateTaskPaneContent(pane.Handle);
                        Assert.IsTrue(PaneEmbedding.Embed(content, pane.Handle, pane.ClientSize.Width, pane.ClientSize.Height), round);
                        Assert.IsTrue(core.WaitForLog("PaneContent: command 'ready'", PageTimeout), round + ": " + core.Log.Dump());
                        Assert.IsFalse(core.Log.Lines.Any(l => l.StartsWith("ERROR") || l.StartsWith("EXCEPTION")), round + ": " + core.Log.Dump());
                        DetachAndUnload(core, pane, content);
                    }
                }
            });
        }

        [TestMethod]
        public void CloseTaskPaneContent_DestroysOnlyThatWindow()
        {
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                using (var core = new LoadedCore(t))
                using (var excel1 = NewPane())
                using (var excel2 = NewPane())
                {
                    IntPtr first = core.Addin.CreateTaskPaneContent(excel1.Handle);
                    IntPtr second = core.Addin.CreateTaskPaneContent(excel2.Handle);
                    core.Addin.CloseTaskPaneContent(excel1.Handle);
                    Assert.IsFalse(Win32Test.IsWindow(first));
                    Assert.IsTrue(Win32Test.IsWindow(second));
                    Assert.AreNotEqual(first, core.Addin.CreateTaskPaneContent(excel1.Handle), "a closed pane gets new content");
                }
            });
        }
    }
}
