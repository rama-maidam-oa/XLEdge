using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Pane
{
    /// <summary>
    /// The content of one Excel window's task pane: a borderless top-level window holding a WebView2. Only its handle crosses
    /// to the shell, which re-parents it into the ADX pane. Until then it sits off-screen, out of the taskbar and Alt-Tab, and
    /// is never activated (showing it must not take the focus from Excel).
    /// </summary>
    internal sealed class PaneContentWindow : Form
    {
        private const int WS_EX_TOOLWINDOW = 0x00000080;
        private const string CommandPrefix = "xledge:";
        private const string LogPrefix = "PaneContent: ";

        private readonly IXLEdgeContext _context;
        private readonly Func<Task<CoreWebView2Environment>> _environment;
        private readonly WebView2 _web;
        private bool _started;

        public PaneContentWindow(IXLEdgeContext context, Func<Task<CoreWebView2Environment>> environment)
        {
            _context = context;
            _environment = environment;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(-32000, -32000);
            Size = new Size(400, 600);
            BackColor = Color.White;
            _web = new WebView2 { Dock = DockStyle.Fill, DefaultBackgroundColor = Color.White };
            Controls.Add(_web);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        /// <summary>Starts the WebView2 and the spike page once (idempotent; runs on, and returns to, the calling UI thread).</summary>
        public void Start()
        {
            if (_started) return;
            _started = true;
            StartAsync();
        }

        public void OnShown() => Log(l => l.LogDebug(LogPrefix + $"pane shown ({ClientSize.Width}x{ClientSize.Height})."));

        // async void: an event-style entry point; every exception is caught here.
        private async void StartAsync()
        {
            try
            {
                CoreWebView2Environment environment = await _environment();
                if (IsDisposed) return;
                await _web.EnsureCoreWebView2Async(environment);
                if (IsDisposed) return;
                CoreWebView2 core = _web.CoreWebView2;
                core.Settings.AreDevToolsEnabled = true;
                core.NavigationCompleted += (s, e) => Log(l => l.LogDebug(LogPrefix + $"navigation completed (success={e.IsSuccess}, status={e.WebErrorStatus})."));
                core.DocumentTitleChanged += (s, e) => OnTitle(core.DocumentTitle);
                core.ProcessFailed += (s, e) => Log(l => l.LogError(LogPrefix + $"WebView2 process failed: {e.ProcessFailedKind} ({e.Reason}).", null));
                Log(l => l.LogDebug(LogPrefix + $"WebView2 {environment.BrowserVersionString} ready."));
                core.NavigateToString(SpikePage());
            }
            catch (Exception ex)
            {
                // Typically the WebView2 Runtime is missing: the pane stays empty, Excel is unaffected.
                Log(l => l.LogError(LogPrefix + "the WebView2 could not be started.", ex.ToString()));
            }
        }

        private void OnTitle(string title)
        {
            if (title == null || !title.StartsWith(CommandPrefix, StringComparison.Ordinal)) return;
            string command = title.Substring(CommandPrefix.Length);
            Log(l => l.LogDebug(LogPrefix + $"command '{command}'."));
        }

        private string SpikePage()
        {
            using (Stream stream = typeof(PaneContentWindow).Assembly.GetManifestResourceStream("XLEdge.Addin.Core.Pane.SpikePage.html"))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd().Replace("{{VERSION}}", WebUtility.HtmlEncode(_context?.Version ?? "unknown"));
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                try { _web.Dispose(); }
                catch (Exception ex) { Log(l => l.LogError(LogPrefix + "disposing the WebView2 failed.", ex.ToString())); }
            }
            base.Dispose(disposing);
        }

        private void Log(Action<ILogger> write)
        {
            try
            {
                ILogger logger = _context?.Logger;
                if (logger != null) write(logger);
            }
            catch (Exception)
            {
                // Logging must never break the pane.
            }
        }
    }
}
