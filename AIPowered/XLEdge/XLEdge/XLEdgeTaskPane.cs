using System;
using System.Drawing;
using System.Reflection;
using System.Windows;
using System.Windows.Resources;
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
        // The Excel window handle the content was asked for: the core's key for it. Kept, not re-read through WindowObj at
        // teardown (WindowObj may already be gone, and a different value would close another pane's content).
        private IntPtr _key;

        /// <summary>Set by the shell right before it shows a pane; Excel's own attempts to restore a pane are ignored.</summary>
        internal static bool ShowRequested { get; set; }

        public XLEdgeTaskPane()
        {
            InitializeComponent();
            Icon = LoadTaskPaneIcon();
            HandleCreated += (s, e) => EmbedContent();
            HandleDestroyed += (s, e) => OnPaneDestroyed();
            Resize += (s, e) => PaneEmbedding.Resize(_content, ClientSize.Width, ClientSize.Height);
        }

        private static Icon LoadTaskPaneIcon()
        {
            // Touching PackUriHelper registers the pack: scheme; without it (no WPF loaded yet in Excel) the Uri below
            // parses "application:,,," as a host:port and throws UriFormatException "Invalid port specified".
            _ = System.IO.Packaging.PackUriHelper.UriSchemePack;
            StreamResourceInfo resource = Application.GetResourceStream(
                new Uri("pack://application:,,,/XLEdge;component/Resources/Orbit.ico", UriKind.Absolute));
            if (resource == null)
                throw new InvalidOperationException("The XLEdge task pane icon resource could not be loaded.");

            using (resource.Stream)
            using (var icon = new Icon(resource.Stream))
                return (Icon)icon.Clone();
        }

        /// <summary>
        /// The pane's window is going (its Excel window closed, or ADX recreates it): take the content out first - Windows
        /// would destroy it as a child - and let the core dispose it with its WebView2; otherwise it stays alive in the core.
        /// </summary>
        private void OnPaneDestroyed()
        {
            IntPtr content = _content, key = _key;
            _content = IntPtr.Zero;
            _key = IntPtr.Zero;
            try
            {
                PaneEmbedding.Release(content);
                if (key != IntPtr.Zero) global::XLEdge.AddinModule.CurrentInstance?.Host?.Addin?.CloseTaskPaneContent(key);
            }
            catch (Exception ex)
            {
                global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane: closing the content of a destroyed pane failed.", ex);
            }
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
                // Not readable yet (WindowObj unset): AfterTaskPaneShow tries again. A Zero key would be shared by every pane.
                IntPtr key = ExcelWindowHandle;
                if (key == IntPtr.Zero) return;
                // Created under this pane's DPI awareness, or Windows refuses to parent it here.
                IntPtr content = PaneEmbedding.CreateUnderDpiOf(Handle, () => addin.CreateTaskPaneContent(key));
                if (PaneEmbedding.Embed(content, Handle, ClientSize.Width, ClientSize.Height, out int error))
                {
                    _content = content;
                    _key = key;
                }
                else if (content != IntPtr.Zero)
                {
                    global::XLEdge.AddinModule.CurrentInstance?.LogError($"XLEdgeTaskPane: the content could not be placed in the pane (Win32 error {error}).", null);
                    addin.CloseTaskPaneContent(key);
                }
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
            IntPtr content = _content, key = _key;
            _content = IntPtr.Zero;
            _key = IntPtr.Zero;
            try
            {
                PaneEmbedding.Release(content);
                if (key != IntPtr.Zero) core?.CloseTaskPaneContent(key);
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
            NotifyShown();
        }

        /// <summary>Tells the core this pane is shown (its AfterTaskPaneShow): after ADX shows it, and after a reload re-embeds it.</summary>
        internal void NotifyShown()
        {
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
