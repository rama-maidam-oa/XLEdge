using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Web.WebView2.Core;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Pane
{
    /// <summary>
    /// One <see cref="PaneContentWindow"/> per Excel window, created on the calling thread (Excel's main thread: a call from the
    /// shell runs on the shell's thread). All of them share one WebView2 environment per core AppDomain, with today's XLEdge
    /// options and user data folder.
    /// </summary>
    internal sealed class TaskPaneContentHost
    {
        private readonly IXLEdgeContext _context;
        private readonly Dictionary<IntPtr, PaneContentWindow> _windows = new Dictionary<IntPtr, PaneContentWindow>();
        private Task<CoreWebView2Environment> _environment;

        public TaskPaneContentHost(IXLEdgeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public int Count => _windows.Count;

        /// <summary>
        /// The content window for <paramref name="excelWindow"/> (created and started on first use). Zero (the shell could not
        /// read the Excel window) gets nothing: a shared Zero key would hand one pane's content to another.
        /// </summary>
        public IntPtr Create(IntPtr excelWindow)
        {
            if (excelWindow == IntPtr.Zero) return IntPtr.Zero;
            if (_windows.TryGetValue(excelWindow, out PaneContentWindow existing))
            {
                if (!existing.IsDisposed && existing.IsHandleCreated) return existing.Handle;
                // Its pane was destroyed (Excel window closed) and took the child window with it: the form and its WebView2
                // are still alive here, so dispose them before replacing them.
                _windows.Remove(excelWindow);
                existing.Dispose();
                Log(l => l.LogDebug("PaneContent: discarded content whose window was destroyed."));
            }

            var window = new PaneContentWindow(_context, Environment);
            try { window.Show(); }
            catch
            {
                window.Dispose();
                throw;
            }
            _windows[excelWindow] = window;
            window.Start();
            return window.Handle;
        }

        public void Shown(IntPtr excelWindow)
        {
            if (_windows.TryGetValue(excelWindow, out PaneContentWindow window) && !window.IsDisposed) window.OnShown();
        }

        public void Close(IntPtr excelWindow)
        {
            if (!_windows.TryGetValue(excelWindow, out PaneContentWindow window)) return;
            _windows.Remove(excelWindow);
            window.Dispose();
        }

        public void CloseAll()
        {
            foreach (IntPtr excelWindow in _windows.Keys.ToList()) Close(excelWindow);
        }

        private Task<CoreWebView2Environment> Environment()
        {
            if (_environment == null || _environment.IsFaulted || _environment.IsCanceled)
                _environment = CoreWebView2Environment.CreateAsync(null, _context.Paths.BrowserDataPath, EnvironmentOptions());
            return _environment;
        }

        /// <summary>
        /// Today's XLEdge options, and they must stay exactly these in every core version: a reloaded or rolled-back core, or a
        /// second Excel running another core, shares the user data folder and its browser process, and a different set of
        /// options fails with 0x8007139F.
        /// </summary>
        internal static CoreWebView2EnvironmentOptions EnvironmentOptions() =>
            new CoreWebView2EnvironmentOptions { AllowSingleSignOnUsingOSPrimaryAccount = true };

        private void Log(Action<ILogger> write)
        {
            try
            {
                ILogger logger = _context.Logger;
                if (logger != null) write(logger);
            }
            catch (Exception)
            {
                // Logging must never break the pane.
            }
        }
    }
}
