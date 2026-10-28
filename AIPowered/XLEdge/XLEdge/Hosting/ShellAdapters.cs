using System;
using System.Windows.Threading;
using AddinExpress.MSO;
using XLEdge.Contracts;

namespace XLEdge.Hosting
{
    /// <summary>The add-in's <see cref="IShellCommands"/>: the ribbon, the task panes and Excel's thread, as AddinModule has them.</summary>
    internal sealed class ShellCommands : IShellCommands
    {
        private readonly AddinModule _module;
        private readonly IRibbonUI _ribbon;
        private readonly Dispatcher _excelThread;

        /// <param name="excelThread">The dispatcher of Excel's main thread (captured in OnRibbonLoaded).</param>
        public ShellCommands(AddinModule module, IRibbonUI ribbon, Dispatcher excelThread)
        {
            _module = module ?? throw new ArgumentNullException(nameof(module));
            _ribbon = ribbon;
            _excelThread = excelThread ?? throw new ArgumentNullException(nameof(excelThread));
        }

        /// <summary>Queued on Excel's message loop, so it runs after the current call (possibly from the core) has returned.</summary>
        public void Post(Action action)
        {
            _excelThread.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                try { action(); }
                catch (Exception ex) { _module.LogError("ShellCommands: a queued action failed.", ex); }
            }));
        }

        public void RunRibbonAction(string controlId) => _module.RunRibbonAction(controlId);
        public void InvalidateRibbon() => _ribbon?.Invalidate();
        public void InvalidateControl(string ribbonXId) => _ribbon?.InvalidateControl(ribbonXId);
        public void DebugModeChanged(bool enabled) => _module.SetDebugMode(enabled);
        public void ShowTaskPane() => _module.ShowTaskPane();

        public void ToggleTaskPane()
        {
            XLEdgeTaskPane pane = _module.ActivePane;
            if (pane?.Visible == true) pane.Visible = false;
            else _module.ShowTaskPane();
        }

        public void HideAllTaskPanes()
        {
            foreach (XLEdgeTaskPane pane in _module.AllPanes()) pane.Visible = false;
        }

        public bool IsTaskPaneVisible() => _module.ActivePane?.Visible == true;

        public void SetTaskPaneCaption(string caption)
        {
            foreach (XLEdgeTaskPane pane in _module.AllPanes()) pane.Text = caption;
        }
    }

    /// <summary>The add-in's <see cref="IPaneHost"/>: every XLEdge task pane (one per Excel window).</summary>
    internal sealed class PaneHost : IPaneHost
    {
        private readonly AddinModule _module;

        public PaneHost(AddinModule module)
        {
            _module = module ?? throw new ArgumentNullException(nameof(module));
        }

        public void DetachAll(IXLEdgeAddin core)
        {
            foreach (XLEdgeTaskPane pane in _module.AllPanes()) pane.ReleaseContent(core);
        }

        /// <summary>Open panes get the new core's content now; hidden ones get it when they are next shown.</summary>
        public void AttachAll()
        {
            foreach (XLEdgeTaskPane pane in _module.AllPanes())
                if (pane.Visible) pane.EmbedContent();
        }
    }
}
