using System;
using System.Reflection;
using System.Runtime.InteropServices;
using AddinExpress.XL;

namespace XLEdge
{
    /// <summary>
    /// The XLEdge task pane (one per Excel window). It holds no content of its own: the core creates a borderless window for
    /// this Excel window and returns its handle, which is re-parented into the pane and kept sized to it.
    /// </summary>
    // "AddinModule" alone is ADXExcelTaskPane.AddinModule (an object) here, hence global::XLEdge.AddinModule.
    public partial class XLEdgeTaskPane : ADXExcelTaskPane
    {
        private const int GWL_STYLE = -16;
        private const long WS_CHILD = 0x40000000L;
        private const long WS_POPUP = 0x80000000L;
        private const long WS_CAPTION = 0x00C00000L;
        private const long WS_THICKFRAME = 0x00040000L;
        private const long WS_SYSMENU = 0x00080000L;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;

        private IntPtr _content;

        /// <summary>Set by the shell right before it shows a pane; Excel's own attempts to restore a pane are ignored.</summary>
        internal static bool ShowRequested { get; set; }

        public XLEdgeTaskPane()
        {
            InitializeComponent();
            HandleCreated += (s, e) => EmbedContent();
            HandleDestroyed += (s, e) => _content = IntPtr.Zero;
            Resize += (s, e) => ResizeContent();
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
                if (content == IntPtr.Zero) return;

                long style = GetWindowLongPtr(content, GWL_STYLE).ToInt64();
                style &= ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX);
                style |= WS_CHILD;
                SetWindowLongPtr(content, GWL_STYLE, new IntPtr(style));
                SetParent(content, Handle);
                _content = content;
                ResizeContent();
            }
            catch (Exception ex)
            {
                global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane.EmbedContent failed.", ex);
            }
        }

        private void ResizeContent()
        {
            if (_content != IntPtr.Zero) MoveWindow(_content, 0, 0, ClientSize.Width, ClientSize.Height, true);
        }

        private void XLEdgeTaskPane_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            if (!ShowRequested) Visible = false;
        }

        private void XLEdgeTaskPane_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            ShowRequested = false;
            if (_content == IntPtr.Zero) EmbedContent();
            try { global::XLEdge.AddinModule.CurrentInstance?.Host?.Addin?.OnTaskPaneShown(ExcelWindowHandle); }
            catch (Exception ex) { global::XLEdge.AddinModule.CurrentInstance?.LogError("XLEdgeTaskPane: OnTaskPaneShown failed.", ex); }
        }

        private void XLEdgeTaskPane_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm = false;
            Visible = false;
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hWnd, int x, int y, int width, int height, bool repaint);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int index) =>
            IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, index) : new IntPtr(GetWindowLong32(hWnd, index));

        private static void SetWindowLongPtr(IntPtr hWnd, int index, IntPtr value)
        {
            if (IntPtr.Size == 8) SetWindowLongPtr64(hWnd, index, value);
            else SetWindowLong32(hWnd, index, value.ToInt32());
        }
    }
}
