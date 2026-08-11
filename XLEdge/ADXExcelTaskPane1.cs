using AddinExpress.XL;
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Integration;
using System.Windows.Threading;
using XLEdge.Helpers;
using XLEdge.Utilities;
using XLEdge.Views;

namespace XLEdge
{
#nullable enable
    public partial class ADXExcelTaskPane1 : AddinExpress.XL.ADXExcelTaskPane
    {
        private XLEdgeCTP _wpfControl;
        private ElementHost _host;
        private readonly int _minWidthDip = 600;
        private readonly int _minHeightDip = 400;
        private const int DefaultDpi = 96;
        private const int WM_WINDOWPOSCHANGING = 0x0046;
        private const int WM_SIZING = 0x0214;
        private const int WMSZ_LEFT = 1;
        private const int WMSZ_TOPLEFT = 4;
        private const int WMSZ_BOTTOMLEFT = 7;

        [StructLayout(LayoutKind.Sequential)]
        private struct Windowspos
        {
            public IntPtr hwnd;
            public IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public uint flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hWnd);

        private int GetEffectiveDpi()
        {
            try
            {
                if (this.IsHandleCreated)
                {
                    return (int)GetDpiForWindow(this.Handle);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(GetEffectiveDpi)}: GetDpiForWindow failed, falling back to DeviceDpi/default - {ex.Message}");
            }

            return this.DeviceDpi > 0 ? this.DeviceDpi : DefaultDpi;
        }

        public Dispatcher? GetWpfDispatcher()
        {
            try
            {
                return _wpfControl?.Dispatcher;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(GetWpfDispatcher)}: failed to get WPF control dispatcher - {ex.Message}");
                return null;
            }
        }

        public ADXExcelTaskPane1()
        {
            InitializeComponent();

            this.Text = string.Empty;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            ApplyDpiAwareSizing(GetEffectiveDpi());
            this.DpiChanged += XLEdgeReportsPane_DpiChanged;

            // Same root cause and fix as GLSense's GLConfiguratorPane.cs ("Balance Configurator
            // appears zoomed in for some users"): the WPF content's real native window
            // (ElementHost's HwndSource) is NOT created when _wpfControl/_host are constructed -
            // building a WPF object creates no HWND at all. WinForms creates it lazily, only once
            // the handle is actually needed - typically when this task pane's own handle is
            // realized by Excel/ADX and the control tree cascades handle creation down to its
            // children. A "using" block that only covers the WPF object's managed construction
            // (as this used to) reverts the thread's DPI context back to whatever it was before
            // long before that real handle gets created, so the WPF/WebView2 content ends up
            // rendering under whatever DPI awareness happened to be ambient at that later, untimed
            // moment instead of Per-Monitor-V2. That race is what produces content rendering at the
            // wrong scale/position relative to the pane - most visible at >100% display scaling,
            // or when Excel isn't on the primary monitor at load time - because Windows falls back
            // to bitmap-stretching the content to the monitor's actual DPI instead of it rendering
            // natively. Widened the "using" scope to cover ElementHost creation and Controls.Add,
            // and added a HandleCreated hook below to reapply the same context for the normal
            // ADX-driven case where this task pane's own handle - and so the ElementHost's
            // cascade-created handle - isn't realized until after this constructor has returned.
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _wpfControl = new XLEdgeCTP(this);

                _host = new ElementHost
                {
                    Dock = DockStyle.Fill,
                    MinimumSize = this.MinimumSize,
                    Child = _wpfControl
                };

                this.Controls.Add(_host);
            }

            // Covers the case where this task pane's own native handle - and therefore the
            // ElementHost's cascade-created handle - is realized after this constructor returns
            // (the normal case for an ADX-hosted task pane), so the WPF/WebView2 content's
            // HwndSource still ends up created under Per-Monitor-V2.
            this.HandleCreated += ADXExcelTaskPane1_HandleCreated;

            this.Resize += XLEdgeReportsPane_Resize;
            this.ResizeBegin += XLEdgeReportsPane_ResizeBegin;
            this.ResizeEnd += XLEdgeReportsPane_ResizeEnd;
        }

        private void ADXExcelTaskPane1_HandleCreated(object sender, EventArgs e)
        {
            LogUtility.LogDebug($"ADXExcelTaskPane1.HandleCreated: dpi={GetEffectiveDpi()}, host handle already created={_host?.IsHandleCreated}");

            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                // Touching Handle forces WinForms to realize the ElementHost's native window now,
                // while the per-monitor context is active, if it has not already been created by
                // this point.
                if (_host != null)
                {
                    _ = _host.Handle;
                }
            }
        }

        private void ADXExcelTaskPane1_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm = false;
            this.Visible = false;
        }

        private void XLEdgeReportsPane_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            using (new LogUtility.LogScope("ADXExcelTaskPane1.DpiChanged"))
            {
                LogUtility.LogDebug($"DpiChanged: {e.DeviceDpiOld} -> {e.DeviceDpiNew}, pane size before: {this.Width}x{this.Height}");
                ApplyDpiAwareSizing(e.DeviceDpiNew);
                _wpfControl?.RefreshWebViewHeight();
                LogUtility.LogDebug($"DpiChanged: pane size after: {this.Width}x{this.Height}");
            }
        }

        private void ApplyDpiAwareSizing(float dpiX)
        {
            var scale = dpiX / 96f;
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);

            LogUtility.LogDebug($"ApplyDpiAwareSizing: dpi={dpiX}, scale={scale:F2}, minWidthPx={minWidthPx}, minHeightPx={minHeightPx}, pane size before: {this.Width}x{this.Height}");

            this.MinimumSize = new Size(minWidthPx, minHeightPx);
            if (_host != null)
            {
                _host.MinimumSize = this.MinimumSize;
            }

            if (this.Width < minWidthPx)
                this.Width = minWidthPx;
            if (this.Height < minHeightPx)
                this.Height = minHeightPx;

            LogUtility.LogDebug($"ApplyDpiAwareSizing: pane size after: {this.Width}x{this.Height}, host size: {_host?.Width}x{_host?.Height}");
        }

        private void XLEdgeReportsPane_ResizeBegin(object sender, EventArgs e)
        {
            LogUtility.LogDebug($"ResizeBegin: pane size={this.Width}x{this.Height}");
        }

        private void XLEdgeReportsPane_ResizeEnd(object sender, EventArgs e)
        {
            LogUtility.LogDebug($"ResizeEnd: pane size={this.Width}x{this.Height}, host size={_host?.Width}x{_host?.Height}");
            _wpfControl?.RefreshWebViewHeight();
        }

        // Guards against infinite Resize reentrancy: setting this.Width/_host.Width below
        // synchronously re-raises this same Resize event before the setter returns. Confirmed via
        // a real test log at >100% DPI - without this guard, the reentrant call re-ran the exact
        // same clamp logic, which (for reasons not fully pinned down at the Win32-message level,
        // but consistently reproduced) kept flipping the pane back to a too-narrow width right
        // after each clamp, producing dozens of recursive Resize calls per second - visually
        // indistinguishable from "distorted" while it churned, and it only ever visually settled
        // once a later, non-reentrant interaction (e.g. clicking the pane's border) issued one
        // more resize that didn't re-trigger the loop. With the guard, a reentrant call just
        // returns immediately instead of re-clamping, so each real resize event settles in at
        // most two steps (the user's/OS's resize, then this handler's one corrective clamp).
        private bool _isEnforcingMinWidth;

        private void XLEdgeReportsPane_Resize(object sender, EventArgs e)
        {
            if (_isEnforcingMinWidth)
            {
                return;
            }

            using (new LogUtility.LogScope("ADXExcelTaskPane1.Resize"))
            {
                try
                {
                    var dpi = GetEffectiveDpi();
                    var dipWidth = this.Width * DefaultDpi / (float)dpi;

                    LogUtility.LogDebug($"Resize: dpi={dpi}, pane size={this.Width}x{this.Height}, dipWidth={dipWidth:F1} (min required={_minWidthDip}), host size={_host?.Width}x{_host?.Height}, WebCtrl actual size={_wpfControl?.GetWebCtrlActualSize()}");

                    if (dipWidth < _minWidthDip)
                    {
                        int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                        LogUtility.LogDebug($"Resize: dipWidth below minimum - clamping pane width to {minWidthPx}px");

                        _isEnforcingMinWidth = true;
                        try
                        {
                            this.Width = minWidthPx;
                            if (_host != null)
                            {
                                _host.Width = minWidthPx;
                            }
                        }
                        finally
                        {
                            _isEnforcingMinWidth = false;
                        }
                    }

                    _wpfControl?.RefreshWebViewHeight();
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"XLEdgeReportsPane_Resize error: {ex.Message}");
                }
            }
        }

        private void ADXExcelTaskPane1_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                if (!XLEdgeAppState.Instance.EdgePaneShown)
                {
                    this.Visible = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in BeforeTaskPaneShow");
            }
        }

        private void ADXExcelTaskPane1_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            using (new LogUtility.LogScope("ADXExcelTaskPane1.AfterTaskPaneShow"))
            {
                try
                {
                    var dpi = GetEffectiveDpi();
                    int targetWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                    LogUtility.LogDebug($"AfterTaskPaneShow: dpi={dpi}, targetWidthPx={targetWidthPx}, pane size before: {this.Width}x{this.Height}, WebCtrl actual size={_wpfControl?.GetWebCtrlActualSize()}");

                    if (this.Width < targetWidthPx)
                    {
                        this.Width = targetWidthPx;
                    }

                    XLEdgeAppState.Instance.EdgePaneShown = false;
                    _wpfControl?.RefreshWebViewHeight();
                    LogUtility.LogDebug($"AfterTaskPaneShow: pane size after: {this.Width}x{this.Height}");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Error in AfterTaskPaneShow");
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_SIZING && m.LParam != IntPtr.Zero)
            {
                var rc = (Rect)Marshal.PtrToStructure(m.LParam, typeof(Rect));
                int width = rc.Right - rc.Left;
                if (width < this.MinimumSize.Width)
                {
                    int minWidth = this.MinimumSize.Width;
                    LogUtility.LogDebug($"WndProc WM_SIZING: requested width={width} below MinimumSize.Width={minWidth} - clamping (wParam={(int)m.WParam})");
                    switch ((int)m.WParam)
                    {
                        case WMSZ_LEFT:
                        case WMSZ_TOPLEFT:
                        case WMSZ_BOTTOMLEFT:
                            rc.Left = rc.Right - minWidth;
                            break;
                        default:
                            rc.Right = rc.Left + minWidth;
                            break;
                    }

                    Marshal.StructureToPtr(rc, m.LParam, true);
                }
            }
            else if (m.Msg == WM_WINDOWPOSCHANGING && m.LParam != IntPtr.Zero)
            {
                var pos = (Windowspos)Marshal.PtrToStructure(m.LParam, typeof(Windowspos));
                if (pos.cx < this.MinimumSize.Width)
                {
                    LogUtility.LogDebug($"WndProc WM_WINDOWPOSCHANGING: requested cx={pos.cx} below MinimumSize.Width={this.MinimumSize.Width} - clamping");
                    pos.cx = this.MinimumSize.Width;
                    Marshal.StructureToPtr(pos, m.LParam, true);
                }
            }

            base.WndProc(ref m);
        }

        protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
        {
            var dpi = GetEffectiveDpi();
            int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
            int minHeightPx = (int)Math.Round(_minHeightDip * dpi / (float)DefaultDpi);

            int requestedWidth = width;
            int requestedHeight = height;

            if ((specified & BoundsSpecified.Width) != 0 && width < minWidthPx)
            {
                width = minWidthPx;
            }

            if ((specified & BoundsSpecified.Height) != 0 && height < minHeightPx)
            {
                height = minHeightPx;
            }

            LogUtility.LogDebug($"SetBoundsCore: requested=({x},{y},{requestedWidth}x{requestedHeight}), specified={specified}, minWidthPx={minWidthPx}, applied=({x},{y},{width}x{height})");

            base.SetBoundsCore(x, y, width, height, specified);
        }

        public async Task<bool> LogoutAsync(string loginUrl, CancellationToken token)
        {
            try
            {
                if (_wpfControl == null)
                {
                    LogUtility.LogWarn("WPF control is null in LogoutAsync.");
                    return false;
                }

                return await _wpfControl.LogoutSessionAsync(loginUrl, token);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("LogoutAsync cancelled.");
                throw;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to logout in ADXExcelTaskPane1.LogoutAsync");
                return false;
            }
        }

        public async Task ExecuteScriptAsync(string script)
        {
            try
            {
                if (_wpfControl == null)
                {
                    LogUtility.LogWarn("WPF control is null in ExecuteScriptAsync.");
                    return;
                }

                await _wpfControl.ExecuteScriptAsync(script);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to execute script in ADXExcelTaskPane1.ExecuteScriptAsync");
            }
        }

        public async Task RefreshLoginNavigationAsync()
        {
            try
            {
                if (_wpfControl == null)
                {
                    LogUtility.LogWarn("WPF control is null in RefreshLoginNavigationAsync.");
                    return;
                }

                await _wpfControl.RefreshLoginNavigationAsync();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to refresh login navigation in ADXExcelTaskPane1.RefreshLoginNavigationAsync");
            }
        }

        public async Task NavigateBlankAsync()
        {
            try
            {
                if (_wpfControl == null)
                {
                    LogUtility.LogWarn("WPF control is null in NavigateBlankAsync.");
                    return;
                }

                await _wpfControl.NavigateBlankAsync();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to navigate WebView blank in ADXExcelTaskPane1.NavigateBlankAsync");
            }
        }

        public void HidePaneSafe()
        {
            try
            {
                if (this.IsDisposed)
                {
                    return;
                }

                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(HidePaneSafe));
                    return;
                }

                this.Visible = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to hide ADXExcelTaskPane1 safely.");
            }
        }

        public void ReleaseFocusToExcel()
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(ReleaseFocusToExcel));
                    return;
                }

                if (_wpfControl != null)
                {
                    _wpfControl.ReleaseFocusToExcel();
                }

                try
                {
                    this.Focus();

                    var excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                    if (excelApp != null)
                    {
                        ExcelWindowHelper.ActivateExcelMainWindow(excelApp);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ADXExcelTaskPane1.ReleaseFocusToExcel: focus failed - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ADXExcelTaskPane1.ReleaseFocusToExcel failed");
            }
        }

        public void RefreshWebViewHeight()
        {
            try
            {
                if (_wpfControl != null)
                {
                    _wpfControl.RefreshWebViewHeight();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"RefreshWebViewHeight error: {ex.Message}");
            }
        }
    }
#nullable restore
}