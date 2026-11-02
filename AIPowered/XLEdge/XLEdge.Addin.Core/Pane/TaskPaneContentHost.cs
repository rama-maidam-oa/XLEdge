using System;
using System.Collections.Generic;
using System.Linq;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Pane
{
    /// <summary>
    /// One <see cref="ADXExcelTaskPane1"/> (11.1.2's pane, as the content of the shell's pane) per Excel window, created on the
    /// calling thread (Excel's main thread: a call from the shell runs on the shell's thread). Their WebView2s share one
    /// environment per core AppDomain (11.1.2's, with today's options and user data folder).
    /// </summary>
    internal sealed class TaskPaneContentHost
    {
        private readonly IXLEdgeContext _context;
        private readonly Dictionary<IntPtr, ADXExcelTaskPane1> _panes = new Dictionary<IntPtr, ADXExcelTaskPane1>();

        public TaskPaneContentHost(IXLEdgeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
        }

        public int Count => _panes.Count;

        /// <summary>The pane of an Excel window, or null.</summary>
        public ADXExcelTaskPane1 Pane(IntPtr excelWindow) =>
            excelWindow != IntPtr.Zero && _panes.TryGetValue(excelWindow, out ADXExcelTaskPane1 pane) && !pane.IsDisposed ? pane : null;

        /// <summary>Every pane, in creation order (11.1.2's TaskPaneInstances).</summary>
        public IList<ADXExcelTaskPane1> Panes => _panes.Values.Where(p => !p.IsDisposed).ToList();

        /// <summary>
        /// The content window for <paramref name="excelWindow"/> (created on first use; creating its handle starts the WebView2 and
        /// the login navigation, as 11.1.2's pane did). Zero (the shell could not read the Excel window) gets nothing: a shared
        /// Zero key would hand one pane's content to another.
        /// </summary>
        public IntPtr Create(IntPtr excelWindow)
        {
            if (excelWindow == IntPtr.Zero) return IntPtr.Zero;
            if (_panes.TryGetValue(excelWindow, out ADXExcelTaskPane1 existing))
            {
                if (!existing.IsDisposed && existing.IsHandleCreated) return existing.Handle;
                // Its pane was destroyed (Excel window closed) and took the child window with it: the form and its WebView2
                // are still alive here, so dispose them before replacing them.
                _panes.Remove(excelWindow);
                existing.Dispose();
                Log(l => l.LogDebug("PaneContent: discarded content whose window was destroyed."));
            }
            var pane = new ADXExcelTaskPane1(excelWindow);
            // Registered first: creating the handle runs 11.1.2's HandleCreated, which may already ask for this pane.
            _panes[excelWindow] = pane;
            try
            {
                // The form's own Show (ADXExcelTaskPane1.Show is the ADX pane's): creates the handle, off-screen, not activated.
                ((System.Windows.Forms.Form)pane).Show();
            }
            catch
            {
                _panes.Remove(excelWindow);
                pane.Dispose();
                throw;
            }
            return pane.Handle;
        }

        public void Shown(IntPtr excelWindow) => Pane(excelWindow)?.OnShownByShell();

        public void Close(IntPtr excelWindow)
        {
            if (!_panes.TryGetValue(excelWindow, out ADXExcelTaskPane1 pane)) return;
            _panes.Remove(excelWindow);
            pane.Dispose();
        }

        public void CloseAll()
        {
            // Each pane in its own try: one WebView2 failing to dispose must not keep the other panes alive into the unload.
            foreach (IntPtr excelWindow in _panes.Keys.ToList())
            {
                try { Close(excelWindow); }
                catch (Exception ex) { Log(l => l.LogError($"TaskPaneContentHost.CloseAll: closing the pane of window {excelWindow} failed.", ex.ToString())); }
            }
        }

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
