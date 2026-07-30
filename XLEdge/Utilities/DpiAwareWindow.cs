using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using Wpf.Ui.Controls;
using XLEdge.Helpers;

namespace XLEdge.Utilities
{
    // Base window class for custom-chrome dialogs: handles per-monitor DPI awareness, scales window
    // content to fit the available screen work area, and provides Escape-to-close plus dialog/owner
    // positioning helpers for derived windows.
    public class DpiAwareWindow : FluentWindow
    {
        private HwndSource _hwndSource;
        private double _currentScaleFactor = 1.0;
        private readonly ScaleTransform _dpiScaleTransform = new ScaleTransform(1.0, 1.0);
        private readonly string _windowName;
        private bool _layoutRefreshPending;
        private bool _initialLayoutApplied;
        private double _initialMaxWidth = double.NaN;
        private double _initialMaxHeight = double.NaN;
        private double _initialMinHeight = double.NaN;

        public bool EnableAutoLayoutRefresh { get; set; } = true;
        public bool EnableExcelCentering { get; set; } = true;
        public bool EnableEscapeToClose { get; set; } = true;

        public double CurrentScaleFactor => _currentScaleFactor;
        public bool AutoClampToWorkArea { get; set; } = true;
        public double WorkAreaMargin { get; set; } = 24d;
        public double? MaxWidthCap { get; set; } = 1400d;
        public double? MaxHeightCap { get; set; } = null;

        /// <summary>
        /// When set to true, completely disables all auto-sizing and clamping logic.
        /// Use this for message boxes and dialogs that should respect user resizing.
        /// </summary>
        public bool DisableAutoSizing { get; set; } = false;

        public double MinContentScale { get; set; } = 0.85;

        public DpiAwareWindow()
        {
            try
            {
                _windowName = GetType().Name;

                // Defensive fallback: ensures Wpf.Ui's theme/resources are initialized before this
                // window's own resources are resolved, in case bootstrap hasn't run yet.
                if (!WpfUiBootstrapper.IsInitialized)
                {
                    WpfUiBootstrapper.Init(XLEdgeAppConstants.GLAccentHex, XLEdgeAppConstants.GLTheme);
                }

                // FluentWindow's ExtendsContentIntoTitleBar defaults to true, which is incompatible
                // with this codebase's fully custom-chrome windows (WindowStyle="None" +
                // AllowsTransparency="True"). Disabling it here centrally avoids that conflict for
                // every derived window.
                this.ExtendsContentIntoTitleBar = false;

                AddHandler(UIElement.PreviewMouseDownEvent, new MouseButtonEventHandler(OnWindowPreviewMouseDown), true);
                AddHandler(UIElement.PreviewKeyDownEvent, new KeyEventHandler(OnWindowPreviewKeyDown), true);
                AddHandler(UIElement.PreviewTextInputEvent, new TextCompositionEventHandler(OnWindowPreviewTextInput), true);

                WpfAppManager.EnsureApplication();

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        using (DpiAwarenessHelper.SetPerMonitorAware())
                        {
                            this.UseLayoutRounding = true;
                            this.SnapsToDevicePixels = true;
                            TextOptions.SetTextFormattingMode(this, TextFormattingMode.Display);
                            RenderOptions.SetBitmapScalingMode(this, BitmapScalingMode.HighQuality);

                            this.SourceInitialized += OnSourceInitialized;
                            this.Loaded += OnLoaded;
                            this.ContentRendered += OnContentRenderedDebug;
                            this.Closed += OnClosedDebug;
                            this.Unloaded += OnUnloadedDebug;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogError($"Error in DpiAwareWindow constructor: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Fatal error in DpiAwareWindow: {ex.Message}");
            }
        }

        /// <summary>
        /// Prevents FluentWindow's base implementation from coercing WindowStyle, which would
        /// conflict with this codebase's custom-drawn window chrome.
        /// </summary>
        protected override void OnExtendsContentIntoTitleBarChanged(bool oldValue, bool newValue)
        {
            // Intentionally does not call the base implementation - see summary above.
        }

        protected override void OnInitialized(EventArgs e)
        {
            try
            {
                base.OnInitialized(e);
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception during window initialization: {ex.Message}");
            }
        }

        public void SetExcelOwner(IntPtr excelHwnd)
        {
            try
            {
                var helper = new WindowInteropHelper(this);
                helper.Owner = excelHwnd;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error setting Excel owner: {ex.Message}");
            }
        }

        public bool? ShowDialogWithOwner(IntPtr excelHwnd)
        {
            try
            {
                SetExcelOwner(excelHwnd);
                return this.ShowDialog();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception showing dialog: {ex.Message}");
                System.Threading.Thread.Sleep(100);
                SetExcelOwner(excelHwnd);
                return this.ShowDialog();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error showing dialog: {ex.Message}");
                try
                {
                    return this.ShowDialog();
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogError($"Critical error showing dialog: {innerEx.Message}");
                    return null;
                }
            }
        }

        /// <summary>
        /// Same as ShowDialogWithOwner(IntPtr), but also positions the window at an explicit
        /// SCREEN-PIXEL location (e.g. the calendar popup that anchors itself under the Excel cell the
        /// user clicked, computed via Application.ActiveWindow.PointsToScreenPixelsX/Y) instead of
        /// letting WPF center it on its owner.
        /// </summary>
        public bool? ShowDialogWithOwner(IntPtr excelHwnd, double explicitLeft, double explicitTop)
        {
            try
            {
                SetExcelOwner(excelHwnd);
                PositionAtScreenPixels(explicitLeft, explicitTop);
                return this.ShowDialog();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception showing dialog: {ex.Message}");
                System.Threading.Thread.Sleep(100);
                SetExcelOwner(excelHwnd);
                PositionAtScreenPixels(explicitLeft, explicitTop);
                return this.ShowDialog();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error showing dialog: {ex.Message}");
                try
                {
                    return this.ShowDialog();
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogError($"Critical error showing dialog: {innerEx.Message}");
                    return null;
                }
            }
        }

        /// <summary>
        /// Converts an explicit SCREEN-PIXEL location to WPF DIPs and applies it as this window's
        /// startup position. WPF's Left/Top are DIPs (96 = 1:1), not raw screen pixels, so the pixel
        /// coordinates are converted using the DPI of the monitor actually under that point - not this
        /// window's own scale factor, which isn't known yet since the window hasn't been shown/sourced -
        /// so this positions correctly even on a mixed-DPI multi-monitor setup where the target monitor
        /// differs from the primary one.
        /// </summary>
        private void PositionAtScreenPixels(double explicitLeft, double explicitTop)
        {
            try
            {
                double scale = DpiAwarenessHelper.GetDpiForScreenPoint(explicitLeft, explicitTop) / 96.0;
                if (scale <= 0)
                {
                    scale = 1.0;
                }

                WindowStartupLocation = WindowStartupLocation.Manual;
                Left = explicitLeft / scale;
                Top = explicitTop / scale;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error positioning window at explicit screen location: {ex.Message}");
            }
        }

        public void ShowWithOwner(IntPtr excelHwnd)
        {
            try
            {
                SetExcelOwner(excelHwnd);
                this.Show();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception showing window: {ex.Message}");
                System.Threading.Thread.Sleep(100);
                SetExcelOwner(excelHwnd);
                this.Show();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error showing window: {ex.Message}");
                try
                {
                    this.Show();
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogError($"Critical error showing window: {innerEx.Message}");
                }
            }
        }

        private void OnWindowPreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            DismissActiveToast();
        }

        private void OnWindowPreviewKeyDown(object sender, KeyEventArgs e)
        {
            DismissActiveToast();

            if (e.Key == Key.Escape && IsInteractionOverlayVisible())
            {
                e.Handled = true;
                return;
            }

            if (EnableEscapeToClose && e.Key == Key.Escape)
            {
                e.Handled = true;
                Close();
            }
        }

        private void OnWindowPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            DismissActiveToast();
        }

        private void DismissActiveToast()
        {
            try
            {
                if (FindName("AppOverlayControl") is FrameworkElement overlay)
                {
                    var dismissMethod = overlay.GetType().GetMethod("DismissToast", Type.EmptyTypes);
                    dismissMethod?.Invoke(overlay, null);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[{_windowName}] toast dismiss ignored: {ex.Message}");
            }
        }

        private bool IsInteractionOverlayVisible()
        {
            try
            {
                if (FindName("AppOverlayControl") is XLEdge.Views.AppOverlay overlay)
                    return overlay.IsBusyVisible || overlay.IsConfirmVisible;
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"[{_windowName}] overlay interaction check ignored: {ex.Message}");
            }

            return false;
        }

        private void OnSourceInitialized(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] source initialized");
                _hwndSource = PresentationSource.FromVisual(this) as HwndSource;
                _hwndSource?.AddHook(WndProc);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in OnSourceInitialized: {ex.Message}");
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] loaded - applying DPI adjustments");


                if (!DisableAutoSizing)
                {
                    CaptureInitialWindowConstraints();
                    QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Loaded);
                }

                // For non-Manual SizeToContent windows, force a resettle so the window doesn't get
                // stuck at a stale/undersized initial measurement. No-op for Manual-sized windows.
                if (!DisableAutoSizing && this.SizeToContent != SizeToContent.Manual)
                {
                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in OnLoaded: {ex.Message}");
            }
        }

        private void QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority priority)
        {
            if (_layoutRefreshPending)
                return;

            _layoutRefreshPending = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                _layoutRefreshPending = false;
                ApplyLayoutRefresh();
            }), priority);
        }

        private void ApplyLayoutRefresh()
        {
            try
            {
                if (!EnableAutoLayoutRefresh || DisableAutoSizing)
                    return;

                AdjustForCurrentDpi();
                FitToAvailableWorkArea();

                if (EnableExcelCentering && !_initialLayoutApplied)
                {
                    _initialLayoutApplied = true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error applying layout refresh: {ex.Message}");
            }
        }

        public void RefreshWindowLayout()
        {
            try
            {
                if (DisableAutoSizing)
                    return;

                if (!Dispatcher.CheckAccess())
                {
                    QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Background);
                    return;
                }

                if (Content is FrameworkElement root)
                {
                    root.InvalidateMeasure();
                    root.InvalidateArrange();
                }

                QueueLayoutRefresh(System.Windows.Threading.DispatcherPriority.Render);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error refreshing window layout: {ex.Message}");
            }
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            try
            {
                const int WM_DPICHANGED = 0x02E0;

                if (msg == WM_DPICHANGED && !DisableAutoSizing)
                {
                    uint newDpi = (uint)wParam;
                    AdjustForDpiChange(newDpi, lParam);
                    handled = true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in WndProc: {ex.Message}");
            }

            return IntPtr.Zero;
        }

        private void AdjustForCurrentDpi()
        {
            try
            {
                _currentScaleFactor = GetCurrentScaleFactor();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error adjusting DPI: {ex.Message}");
            }
        }

        private void AdjustForDpiChange(uint newDpi, IntPtr lParam)
        {
            try
            {
                var scaleFactor = GetScaleFactorFromDpi(newDpi);

                ApplyScaleTransform(scaleFactor);

                if (lParam != IntPtr.Zero)
                {
                    var rect = Marshal.PtrToStructure<Rect>(lParam);

                    this.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        try
                        {
                            this.Left = rect.Left / scaleFactor;
                            this.Top = rect.Top / scaleFactor;
                            this.Width = rect.Width / scaleFactor;
                            this.Height = rect.Height / scaleFactor;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogError($"Error resizing window: {ex.Message}");
                        }
                    }));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in AdjustForDpiChange: {ex.Message}");
            }
        }

        private void ApplyScaleTransform(double scaleFactor)
        {
            if (Content is not FrameworkElement element)
                return;

            if (Math.Abs(scaleFactor - 1.0) < 0.001)
            {
                element.LayoutTransform = Transform.Identity;
                return;
            }

            if (Math.Abs(scaleFactor - _currentScaleFactor) < 0.001)
                return;

            try
            {
                _dpiScaleTransform.ScaleX = scaleFactor;
                _dpiScaleTransform.ScaleY = scaleFactor;
                element.LayoutTransform = _dpiScaleTransform;
                element.InvalidateMeasure();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error applying scale transform: {ex.Message}");
            }
        }

        private void FitToAvailableWorkArea()
        {
            if (!AutoClampToWorkArea || DisableAutoSizing)
                return;

            try
            {
                if (Content is not FrameworkElement root)
                    return;

                var workArea = SystemParameters.WorkArea;
                var availableWidth = Math.Max(0, workArea.Width - (WorkAreaMargin * 2));
                var availableHeight = Math.Max(0, workArea.Height - (WorkAreaMargin * 2));

                var requestedMaxWidth = GetEffectiveRequestedMaxWidth();
                if (!double.IsPositiveInfinity(requestedMaxWidth))
                    availableWidth = Math.Min(availableWidth, requestedMaxWidth);

                if (MaxWidthCap.HasValue)
                    availableWidth = Math.Min(availableWidth, MaxWidthCap.Value);

                if (MaxHeightCap.HasValue)
                    availableHeight = Math.Min(availableHeight, MaxHeightCap.Value);

                root.Measure(new Size(double.PositiveInfinity, double.PositiveInfinity));
                var desiredWidth = root.DesiredSize.Width;
                var desiredHeight = root.DesiredSize.Height;

                // If an initial MinHeight was specified on the window, treat it as
                // a logical minimum content height when calculating fit and scale.
                // This prevents the window from being rendered too-small when no
                // explicit content requires space, while still allowing the
                // clamping logic below to reduce the window if the work area is
                // smaller (so it will not prevent clamping to the taskbar).
                if (!double.IsNaN(_initialMinHeight) && _initialMinHeight > 0)
                {
                    desiredHeight = Math.Max(desiredHeight, _initialMinHeight);
                }

                if (desiredWidth <= 0 || desiredHeight <= 0)
                    return;

                var rawScale = Math.Min(availableWidth / desiredWidth, availableHeight / desiredHeight);
                var fitScale = Math.Min(1.0, rawScale);

                if (MinContentScale > 0 && fitScale < MinContentScale)
                {
                    fitScale = MinContentScale;
                }

                ApplyScaleTransform(fitScale);

                var targetWidth = Math.Min(desiredWidth * fitScale, availableWidth);
                var targetHeight = Math.Min(desiredHeight * fitScale, availableHeight);

                // Capture the size/position as they stood before this method changes them, so we
                // can recenter around the same center point afterward (see RecenterAfterSizeChange
                // for why this matters - WindowStartupLocation only centers the window once, and any
                // resize after that anchors at the current Left/Top, silently drifting the window
                // off-center as it grows/shrinks).
                double previousLeft = Left;
                double previousTop = Top;
                double previousWidth = Width;
                double previousHeight = Height;
                bool sizeChanged = false;

                if (targetWidth > 0 && Math.Abs(targetWidth - previousWidth) > 0.5)
                {
                    Width = targetWidth;
                    sizeChanged = true;
                }

                if (targetHeight > 0 && Math.Abs(targetHeight - previousHeight) > 0.5)
                {
                    Height = targetHeight;
                    sizeChanged = true;
                }

                MaxWidth = availableWidth;
                MaxHeight = availableHeight;

                if (sizeChanged)
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error fitting window to work area: {ex.Message}");
            }
        }

        /// <summary>
        /// Recenters the window around the same center point it had before Width/Height were just
        /// changed by FitToAvailableWorkArea/EnsureFitsWorkArea, instead of leaving Left/Top
        /// untouched. A resize always grows/shrinks anchored at the window's current top-left
        /// corner, so without this, any post-centering resize (e.g. content growing once async data
        /// finishes loading, or the safety clamp in EnsureFitsWorkArea kicking in) silently drifts
        /// the window's true center away from wherever WindowStartupLocation originally centered it
        /// (typically CenterOwner against the Excel window) - this was the root cause of windows
        /// appearing off-center. Only called when this class itself changed Width/Height; never
        /// touches Left/Top for a plain user-initiated drag-resize (ResizeMode="CanResize"), since
        /// that doesn't go through either of those two methods unless the drag actually violates
        /// Min/MaxWidth/Height, in which case re-centering after the forced clamp is correct anyway.
        /// Ported from GLSense's DpiAwareWindow.cs, which had the identical resize-without-recenter
        /// bug (see D:\SQLLite_Test\GLSense\FinalWorkingCode's CLAUDE.md for that write-up).
        /// </summary>
        private void RecenterAfterSizeChange(double previousLeft, double previousTop, double previousWidth, double previousHeight)
        {
            try
            {
                if (double.IsNaN(previousLeft) || double.IsNaN(previousTop) ||
                    double.IsNaN(previousWidth) || double.IsNaN(previousHeight) ||
                    previousWidth <= 0 || previousHeight <= 0 ||
                    double.IsNaN(Width) || double.IsNaN(Height) ||
                    Width <= 0 || Height <= 0)
                {
                    return;
                }

                double centerX = previousLeft + (previousWidth / 2.0);
                double centerY = previousTop + (previousHeight / 2.0);

                double newLeft = centerX - (Width / 2.0);
                double newTop = centerY - (Height / 2.0);

                // Clamp so recentering never pushes the window off the visible work area (e.g. if
                // the old center point was near a screen edge).
                var workArea = SystemParameters.WorkArea;
                if (Width < workArea.Width)
                    newLeft = Math.Max(workArea.Left, Math.Min(newLeft, workArea.Right - Width));
                if (Height < workArea.Height)
                    newTop = Math.Max(workArea.Top, Math.Min(newTop, workArea.Bottom - Height));

                Left = newLeft;
                Top = newTop;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error recentering window after size change: {ex.Message}");
            }
        }

        private double GetCurrentScaleFactor()
        {
            try
            {
                if (_hwndSource?.CompositionTarget != null)
                {
                    var scale = _hwndSource.CompositionTarget.TransformToDevice.M11;
                    if (scale > 0)
                        return scale;
                }
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort scale lookup via HwndSource CompositionTarget; falls
                // through to the DpiAwarenessHelper fallback below.
                LogUtility.LogDebug($"{nameof(GetCurrentScaleFactor)}: CompositionTarget scale lookup failed, trying DPI fallback - {ex.Message}");
            }

            try
            {
                var dpi = DpiAwarenessHelper.GetWindowDpi(this);
                if (dpi > 0)
                    return dpi / 96.0;
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort DPI lookup; falls back to 1.0 (no scaling).
                LogUtility.LogDebug($"{nameof(GetCurrentScaleFactor)}: DpiAwarenessHelper.GetWindowDpi failed, defaulting to 1.0 - {ex.Message}");
            }

            return 1.0;
        }

        private static double GetScaleFactorFromDpi(uint dpi)
        {
            var scale = dpi / 96.0;
            return scale > 0 ? scale : 1.0;
        }

        protected double DipToPixels(double dip)
        {
            return dip * _currentScaleFactor;
        }

        protected double PixelsToDip(double pixels)
        {
            return pixels / _currentScaleFactor;
        }

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
            public readonly int Width => Right - Left;
            public readonly int Height => Bottom - Top;
        }

        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            // Skip clamping if auto-sizing is disabled
            if (DisableAutoSizing || !AutoClampToWorkArea)
                return;

            try
            {
                EnsureFitsWorkArea();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error clamping window size: {ex.Message}");
            }
        }

        protected void EnsureFitsWorkArea(double? marginOverride = null)
        {
            // Skip if auto-sizing is disabled
            if (DisableAutoSizing)
                return;

            var margin = marginOverride ?? WorkAreaMargin;
            try
            {
                double previousLeft = Left;
                double previousTop = Top;
                double previousWidth = Width;
                double previousHeight = Height;
                bool sizeChanged = false;

                var workArea = SystemParameters.WorkArea;

                var baseMaxWidth = Math.Max(0, workArea.Width - margin);
                var baseMaxHeight = Math.Max(0, workArea.Height - margin);

                var requestedMaxWidth = GetEffectiveRequestedMaxWidth();

                if (!double.IsPositiveInfinity(requestedMaxWidth))
                {
                    baseMaxWidth = Math.Min(baseMaxWidth, requestedMaxWidth);
                }

                if (MaxWidthCap.HasValue)
                {
                    baseMaxWidth = Math.Min(baseMaxWidth, MaxWidthCap.Value);
                }

                if (MaxHeightCap.HasValue)
                {
                    baseMaxHeight = Math.Min(baseMaxHeight, MaxHeightCap.Value);
                }

                var effectiveMaxWidth = double.IsPositiveInfinity(MaxWidth)
                    ? baseMaxWidth
                    : Math.Min(MaxWidth, baseMaxWidth);

                var effectiveMaxHeight = double.IsPositiveInfinity(MaxHeight)
                    ? baseMaxHeight
                    : Math.Min(MaxHeight, baseMaxHeight);

                MaxWidth = effectiveMaxWidth;
                MaxHeight = effectiveMaxHeight;

                if (MinWidth > effectiveMaxWidth)
                {
                    MinWidth = effectiveMaxWidth;
                }

                if (MinHeight > effectiveMaxHeight)
                {
                    MinHeight = effectiveMaxHeight;
                }

                if (Width > effectiveMaxWidth)
                {
                    Width = effectiveMaxWidth;
                    sizeChanged = true;
                }
                else if (Width < MinWidth)
                {
                    Width = MinWidth;
                    sizeChanged = true;
                }

                if (Height > effectiveMaxHeight)
                {
                    Height = effectiveMaxHeight;
                    sizeChanged = true;
                }
                else if (Height < MinHeight)
                {
                    Height = MinHeight;
                    sizeChanged = true;
                }

                // Only recenter when this method itself just forced a clamp - a plain user
                // drag-resize (ResizeMode="CanResize") stays within Min/MaxWidth/Height and never
                // reaches here, so ordinary manual resizing is left untouched.
                if (sizeChanged)
                    RecenterAfterSizeChange(previousLeft, previousTop, previousWidth, previousHeight);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error ensuring window fits work area: {ex.Message}");
            }
        }

        private void OnContentRenderedDebug(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] content rendered");

                // ContentRendered fires only after the window's content has actually painted, so the
                // resettle here (unlike the one in OnLoaded) can reliably force a native frame update.
                if (!DisableAutoSizing && this.SizeToContent != SizeToContent.Manual)
                {
                    ForceSizeToContentResettle();
                    PumpDispatcherFrame();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in OnContentRenderedDebug: {ex.Message}");
            }
        }

        // Forces a genuine native HWND resize by toggling SizeToContent off/on with a full-pixel
        // Width/Height nudge (sub-pixel nudges round to zero physical device pixels at most DPI
        // scales), then asks Windows to recompute the window's non-client frame via
        // SetWindowPos(SWP_FRAMECHANGED).
        protected void ForceSizeToContentResettle()
        {
            try
            {
                var mode = this.SizeToContent;
                this.SizeToContent = SizeToContent.Manual;
                this.UpdateLayout();

                if (this.ActualWidth > 0 && this.ActualHeight > 0)
                {
                    this.Width = this.ActualWidth + 1.0;
                    this.Height = this.ActualHeight + 1.0;
                    this.UpdateLayout();
                }

                this.SizeToContent = mode;
                this.UpdateLayout();

                if (_hwndSource?.Handle != null && _hwndSource.Handle != IntPtr.Zero)
                {
                    SetWindowPos(_hwndSource.Handle, IntPtr.Zero, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
                }

                LogUtility.LogDebug($"[{_windowName}] SizeToContent resettled ({mode})");
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in ForceSizeToContentResettle: {ex.Message}");
            }
        }

        // WPF's "DoEvents" equivalent: pumps a nested dispatcher frame so pending layout/resize work
        // is fully flushed before the caller continues.
        protected void PumpDispatcherFrame()
        {
            try
            {
                var frame = new System.Windows.Threading.DispatcherFrame();
                this.Dispatcher.BeginInvoke(new Action(() => frame.Continue = false),
                    System.Windows.Threading.DispatcherPriority.Background);
                System.Windows.Threading.Dispatcher.PushFrame(frame);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in PumpDispatcherFrame: {ex.Message}");
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_FRAMECHANGED = 0x0020;

        private void OnClosedDebug(object sender, EventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] closed");
            }
            catch (Exception ex)
            {
                // Safe to ignore/expected: deliberately not routed through LogUtility here since the
                // logging call itself is what just failed - avoid a secondary failure loop.
                System.Diagnostics.Debug.WriteLine($"{nameof(OnClosedDebug)}: failed to log closed event for [{_windowName}] - {ex.Message}");
            }
        }

        private void OnUnloadedDebug(object sender, RoutedEventArgs e)
        {
            try
            {
                LogUtility.LogDebug($"[{_windowName}] unloaded");
            }
            catch (Exception ex)
            {
                // Safe to ignore/expected: deliberately not routed through LogUtility here since the
                // logging call itself is what just failed - avoid a secondary failure loop.
                System.Diagnostics.Debug.WriteLine($"{nameof(OnUnloadedDebug)}: failed to log unloaded event for [{_windowName}] - {ex.Message}");
            }
        }

        private void CaptureInitialWindowConstraints()
        {
            if (double.IsNaN(_initialMaxWidth))
            {
                _initialMaxWidth = MaxWidth;
            }

            if (double.IsNaN(_initialMaxHeight))
            {
                _initialMaxHeight = MaxHeight;
            }

            if (double.IsNaN(_initialMinHeight))
            {
                _initialMinHeight = MinHeight;
            }
        }

        private double GetEffectiveRequestedMaxWidth()
        {
            var maxWidth = double.IsNaN(_initialMaxWidth) ? MaxWidth : _initialMaxWidth;

            if (double.IsPositiveInfinity(maxWidth))
                return double.PositiveInfinity;

            return maxWidth + 200;
        }
    }
}
