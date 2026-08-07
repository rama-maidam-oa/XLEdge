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

            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _wpfControl = new XLEdgeCTP(this);
            }

            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                MinimumSize = this.MinimumSize,
                Child = _wpfControl
            };

            this.Controls.Add(_host);
            this.Resize += XLEdgeReportsPane_Resize;
            this.ResizeBegin += XLEdgeReportsPane_ResizeBegin;
            this.ResizeEnd += XLEdgeReportsPane_ResizeEnd;
        }

        private void ADXExcelTaskPane1_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm = false;
            this.Visible = false;
        }

        private void XLEdgeReportsPane_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            ApplyDpiAwareSizing(e.DeviceDpiNew);
            _wpfControl?.RefreshWebViewHeight();
        }

        private void ApplyDpiAwareSizing(float dpiX)
        {
            var scale = dpiX / 96f;
            int minWidthPx = (int)Math.Round(_minWidthDip * scale);
            int minHeightPx = (int)Math.Round(_minHeightDip * scale);

            this.MinimumSize = new Size(minWidthPx, minHeightPx);
            if (_host != null)
            {
                _host.MinimumSize = this.MinimumSize;
            }

            if (this.Width < minWidthPx)
                this.Width = minWidthPx;
            if (this.Height < minHeightPx)
                this.Height = minHeightPx;
        }

        private void XLEdgeReportsPane_ResizeBegin(object sender, EventArgs e)
        {
            // Handle resize begin if needed
        }

        private void XLEdgeReportsPane_ResizeEnd(object sender, EventArgs e)
        {
            _wpfControl?.RefreshWebViewHeight();
        }

        private void XLEdgeReportsPane_Resize(object sender, EventArgs e)
        {
            try
            {
                var dpi = GetEffectiveDpi();
                var dipWidth = this.Width * DefaultDpi / (float)dpi;

                if (dipWidth < _minWidthDip)
                {
                    int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                    this.Width = minWidthPx;
                    if (_host != null)
                    {
                        _host.Width = minWidthPx;
                    }
                }

                _wpfControl?.RefreshWebViewHeight();
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"XLEdgeReportsPane_Resize error: {ex.Message}");
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
            try
            {
                var dpi = GetEffectiveDpi();
                int targetWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                if (this.Width < targetWidthPx)
                {
                    this.Width = targetWidthPx;
                }

                XLEdgeAppState.Instance.EdgePaneShown = false;
                _wpfControl?.RefreshWebViewHeight();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in AfterTaskPaneShow");
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

            if ((specified & BoundsSpecified.Width) != 0 && width < minWidthPx)
            {
                width = minWidthPx;
            }

            if ((specified & BoundsSpecified.Height) != 0 && height < minHeightPx)
            {
                height = minHeightPx;
            }

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