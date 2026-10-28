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
                    IntPtr foreground = Win32Test.GetForegroundWindow();
                    excel.Activate();
                    IntPtr active = Win32Test.GetActiveWindow();
                    IntPtr content = core.Addin.CreateTaskPaneContent(excel.Handle);
                    // The thread's active window shows an activation even when this test process is not in the foreground.
                    Assert.AreEqual(active, Win32Test.GetActiveWindow(), "showing the content must not activate it");
                    Assert.IsTrue(Win32Test.IsWindow(content), core.Log.Dump());
                    long ex = Win32Test.ExStyle(content);
                    Assert.AreNotEqual(0L, ex & Win32Test.WS_EX_TOOLWINDOW, "tool window: never in the taskbar or Alt-Tab");
                    Assert.AreEqual(0L, ex & Win32Test.WS_EX_APPWINDOW);
                    Assert.AreEqual(foreground, Win32Test.GetForegroundWindow(), "creating the content must not take the focus from Excel");
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
