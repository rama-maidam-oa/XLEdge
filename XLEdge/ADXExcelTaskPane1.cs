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
        private readonly int _minHeightDip = 800;
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
                // Safe to ignore: best-effort DPI lookup via GetDpiForWindow; falls back to
                // this.DeviceDpi/DefaultDpi below.
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
                // Safe to ignore: best-effort dispatcher lookup; caller treats a null return as
                // "WPF control/dispatcher not available yet".
                LogUtility.LogDebug($"{nameof(GetWpfDispatcher)}: failed to get WPF control dispatcher - {ex.Message}");
                return null;
            }
        }
        public ADXExcelTaskPane1()
        {
            InitializeComponent();

            // Enable DPI-aware sizing for the WinForms host
            this.AutoScaleMode = AutoScaleMode.Dpi;

            ApplyDpiAwareSizing(GetEffectiveDpi());
            this.DpiChanged += XLEdgeReportsPane_DpiChanged;

            // Create WPF control with task pane reference under per-monitor DPI context
            using (DpiAwarenessHelper.SetPerMonitorAware())
            {
                _wpfControl = new XLEdgeCTP(this);
            }

            _wpfControl.OnCloseRequested += () => this.Visible = false;

            // Host WPF control inside WinForms
            _host = new ElementHost
            {
                Dock = DockStyle.Fill,
                MinimumSize = this.MinimumSize,
                Child = _wpfControl
            };

            this.Controls.Add(_host);

            // Handle resize events
            this.Resize += XLEdgeReportsPane_Resize;
        }
        private void ADXExcelTaskPane1_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm=false; // Prevent the default close behavior
            this.Visible = false; // Just hide the pane instead of closing
        }
        private void XLEdgeReportsPane_DpiChanged(object sender, DpiChangedEventArgs e)
        {
            ApplyDpiAwareSizing(e.DeviceDpiNew);
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

            this.Size = new Size(Math.Max(this.Width, minWidthPx), Math.Max(this.Height, minHeightPx));
        }
        private void XLEdgeReportsPane_Resize(object sender, EventArgs e)
        {
            var dpi = GetEffectiveDpi();
            var dipWidth = this.Width * DefaultDpi / (float)dpi;

            // Ensure minimum size in DIPs
            if (dipWidth < _minWidthDip)
            {
                int minWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                this.Width = minWidthPx;
                if (_host != null)
                {
                    _host.Width = minWidthPx;
                }
            }

            // Update WPF control if needed
            _wpfControl?.UpdateLayout();
        }
        private void ADXExcelTaskPane1_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                if (!XLEdgeAppState.Instance.EdgePaneShown)
                {
                    this.Visible = false; // Hide the pane until WebView2 is ready
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error enabling DevTools in AfterTaskPaneShow");
            }
        }

        private void ADXExcelTaskPane1_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            try
            {
                // Scale the target pane width by the current DPI so it matches XLEdgeCTP's
                // DIP-based MinWidth, keeping the header's Close button on-screen.
                var dpi = GetEffectiveDpi();
                int targetWidthPx = (int)Math.Round(_minWidthDip * dpi / (float)DefaultDpi);
                if (this.Width != targetWidthPx)
                {
                    this.Width = targetWidthPx;
                }

                XLEdgeAppState.Instance.EdgePaneShown = false;

                //TODO: Revisit this logic - we should not need to get the addin instance here, and if we do it should be via a well-known static property, not a call to CurrentInstance which can have lifecycle issues. For now, we will rely on the WebView2 initialization check before navigating, and remove this code to avoid potential issues.
                //try
                //{
                //    if (addinInstance == null)
                //    {
                //        try
                //        {
                //            addinInstance = AddinModule.CurrentInstance.GetGLSenseAddinObject();
                //        }
                //        catch (Exception ex)
                //        {
                //            addinInstance = null;
                //        }
                //    }
                //}
                //catch (Exception ex)
                //{
                //    XLEdgeLogger.Error(ex);
                //}

               
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
        // Ported from VB's InvokedFromGLSense: triggers a login-URL navigation refresh on the
        // already-hosted WebView2 (used when a GLSense session update arrives while the pane is
        // already visible, so ADXAfterTaskPaneShow won't re-fire to do this automatically).
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

        // Ported from VB's InvokedFromGLSense no-permission branch: blanks the WebView2 before hiding the pane.
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
        /// <summary>
        /// Forces focus away from the WebView2/task pane back to Excel
        /// </summary>
        public void ReleaseFocusToExcel()
        {
            try
            {
                // Ensure we're on the UI thread
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(ReleaseFocusToExcel));
                    return;
                }

                if (_wpfControl != null)
                {
                    _wpfControl.ReleaseFocusToExcel();
                }

                // Additional: force focus away from the task pane itself
                try
                {
                    // Move focus to a dummy control or the task pane itself
                    this.Focus();

                    // Then immediately release focus back to Excel using Windows API
                    var excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                    if (excelApp != null)
                    {
                        // Use Windows API to bring Excel to foreground
                        IntPtr excelHandle = new IntPtr(excelApp.Hwnd);
                        ExcelWindowHelper.ActivateExcelMainWindow(excelApp);

                        // Send a dummy key to force Excel to recognize focus
                        excelApp.SendKeys("{F2}");
                        System.Threading.Thread.Sleep(10);
                        excelApp.SendKeys("{ESC}");
                        System.Threading.Thread.Sleep(10);
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
    }
#nullable restore
}
