using System;
using System.Reflection;
using AddinExpress.XL;
using XLEdge.Hosting;

namespace XLEdge
{
    /// <summary>
    /// The XLEdge task pane (one per Excel window). It holds no content of its own: the core creates a borderless window for
    /// this Excel window and returns its handle, which <see cref="PaneEmbedding"/> re-parents into the pane and keeps sized to it.
    /// </summary>
    // "AddinModule" alone is ADXExcelTaskPane.AddinModule (an object) here, hence global::XLEdge.AddinModule.
    public partial class XLEdgeTaskPane : ADXExcelTaskPane
    {
        private IntPtr _content;

        /// <summary>Set by the shell right before it shows a pane; Excel's own attempts to restore a pane are ignored.</summary>
        internal static bool ShowRequested { get; set; }

        public XLEdgeTaskPane()
        {
            InitializeComponent();
            HandleCreated += (s, e) => EmbedContent();
            HandleDestroyed += (s, e) => _content = IntPtr.Zero;
            Resize += (s, e) => PaneEmbedding.Resize(_content, ClientSize.Width, ClientSize.Height);
        }

        /// <summary>Handle of the Excel window this pane belongs to.</summary>
        internal IntPtr ExcelWindowHandle
        {
            get
            {
                try
                {
                    object window = WindowObj;
                    return window == null ? IntPtr.Zero : new IntPtr(Convert.ToInt64(window.GetType().InvokeMember("Hwnd", BindingFlags.GetProperty, null, window, null)));
                }
                catch (Exception)
                {
                    return IntPtr.Zero;
                }
            }
        }

        /// <summary>Asks the loaded core for this window's content (again after a reload) and re-parents it into the pane.</summary>
        internal void EmbedContent()
        {
            try
            {
                var addin = global::XLEdge.AddinModule.CurrentInstance?.Host?.Addin;
                if (addin == null || !IsHandleCreated) return;
                IntPtr content = addin.CreateTaskPaneContent(ExcelWindowHandle);
                if (PaneEmbedding.Embed(content, Handle, ClientSize.Width, ClientSize.Height)) _content = content;
            }
            catch (Exception ex)
            {
                global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane.EmbedContent failed.", ex);
            }
        }

        /// <summary>
        /// Before the core is unloaded: takes its content out of the pane and asks it to destroy it, so nothing is left
        /// pointing at a window of an unloaded AppDomain. The next core's content is embedded by <see cref="EmbedContent"/>.
        /// </summary>
        internal void ReleaseContent(global::XLEdge.Contracts.IXLEdgeAddin core)
        {
            IntPtr content = _content;
            _content = IntPtr.Zero;
            try
            {
                PaneEmbedding.Release(content);
                core?.CloseTaskPaneContent(ExcelWindowHandle);
            }
            catch (Exception ex)
            {
                global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane.ReleaseContent failed.", ex);
            }
        }

        private void XLEdgeTaskPane_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            if (!ShowRequested) Visible = false;
        }

        private void XLEdgeTaskPane_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            ShowRequested = false;
            if (!PaneEmbedding.IsAlive(_content)) _content = IntPtr.Zero;
            if (_content == IntPtr.Zero) EmbedContent();
            try { global::XLEdge.AddinModule.CurrentInstance?.Host?.Addin?.OnTaskPaneShown(ExcelWindowHandle); }
            catch (Exception ex) { global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane: OnTaskPaneShown failed.", ex); }
        }

        private void XLEdgeTaskPane_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm = false;
            Visible = false;
        }
    }
}
