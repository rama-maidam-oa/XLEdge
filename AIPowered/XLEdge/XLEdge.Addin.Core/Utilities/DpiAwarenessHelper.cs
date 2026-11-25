using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace XLEdge.Utilities
{
    public static class DpiAwarenessHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        // OISR-23017: GetDpiForMonitor lives in shcore.dll - declared against user32.dll it threw
        // EntryPointNotFoundException on every call (caught), so the monitor DPI was never read.
        [DllImport("shcore.dll")]
        private static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);
        private const uint MONITOR_DEFAULTTONEAREST = 0x00000002;
        private const int MDT_EFFECTIVE_DPI = 0;


        [DllImport("user32.dll")]
        private static extern uint GetDpiForSystem();

        private static int _lastLoggedDpi;

        // OISR-23017: WPF only re-renders a Per-Monitor-V2 window at its own monitor's DPI (and on WM_DPICHANGED)
        // when this compatibility switch is off. Hosted in Excel there is no app.config of ours to set it, and its
        // default leaves WPF drawing every window at the DPI Excel started with. AppContext switches are read once
        // and cached, so this must run before the first WPF window/HwndSource is created.
        public static void EnableWpfPerMonitorScaling()
        {
            try
            {
                AppContext.SetSwitch("Switch.System.Windows.DoNotScaleForDpiChanges", false);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(EnableWpfPerMonitorScaling)}: {ex.Message}");
            }
        }

        // OISR-23017: the effective DPI of the monitor a window is on, independent of the calling thread's DPI
        // awareness (GetDpiForMonitor reports the system DPI to a non-per-monitor caller). 0 if it can't be read.
        public static uint GetMonitorDpi(IntPtr hwnd)
        {
            try
            {
                if (hwnd == IntPtr.Zero)
                    return 0;

                using (SetPerMonitorAware())
                {
                    var monitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
                    if (monitor != IntPtr.Zero && GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, out uint dpiX, out _) == 0)
                        return dpiX;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(GetMonitorDpi)}: monitor DPI lookup failed - {ex.Message}");
            }

            return 0;
        }

        // OISR-23017: the log records the scale Excel started at; a scale change after that (docking, Settings >
        // Display) was invisible in the log. Logs one WARN line per change. Baseline is the system DPI.
        public static void LogIfDpiChanged(uint dpi, string source)
        {
            try
            {
                if (dpi == 0)
                    return;

                int previous = System.Threading.Interlocked.Exchange(ref _lastLoggedDpi, (int)dpi);
                if (previous == 0)
                    previous = (int)GetDpiForSystem();

                if (previous != 0 && previous != (int)dpi)
                    LogUtility.LogWarn($"Display scale changed: {previous} DPI ({previous * 100 / 96}%) -> {dpi} DPI ({dpi * 100 / 96}%), detected by {source}");
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(LogIfDpiChanged)}: {ex.Message}");
            }
        }

        public static IDisposable SetPerMonitorAware()
        {
            IntPtr oldContext = IntPtr.Zero;

            try
            {
                oldContext = SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
            }
            catch (Exception ex)
            {
                // Safe to ignore/expected: SetThreadDpiAwarenessContext isn't available on older
                // Windows versions - fall back to whatever DPI awareness is already in effect.
                LogUtility.LogDebug($"{nameof(SetPerMonitorAware)}: SetThreadDpiAwarenessContext failed (older Windows?) - {ex.Message}");
            }

            return new DpiContextDisposer(oldContext);
        }

        public static double GetWindowDpi(Window window)
        {
            try
            {
                var presentationSource = PresentationSource.FromVisual(window);
                if (presentationSource?.CompositionTarget != null)
                {
                    return 96.0 * presentationSource.CompositionTarget.TransformToDevice.M11;
                }
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort DPI lookup via PresentationSource; falls back to the
                // Win32 monitor/window DPI lookup below.
                LogUtility.LogDebug($"{nameof(GetWindowDpi)}: PresentationSource DPI lookup failed, trying Win32 fallback - {ex.Message}");
            }

            try
            {
                var handle = new System.Windows.Interop.WindowInteropHelper(window).Handle;
                if (handle != IntPtr.Zero)
                {
                    // Try to get monitor DPI first for better accuracy
                    var monitor = MonitorFromWindow(handle, MONITOR_DEFAULTTONEAREST);
                    if (monitor != IntPtr.Zero)
                    {
                        if (GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, out uint dpiX, out uint dpiY) == 0)
                        {
                            return dpiX;
                        }
                    }

                    // Fallback to window DPI
                    return GetDpiForWindow(handle);
                }
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort Win32 DPI lookup; falls back to 96 (100% scaling).
                LogUtility.LogDebug($"{nameof(GetWindowDpi)}: Win32 monitor/window DPI lookup failed, defaulting to 96 - {ex.Message}");
            }


            return 96.0;
        }

        /// <summary>
        /// Returns the effective DPI of whichever monitor is under the given SCREEN-PIXEL point (e.g.
        /// coordinates from Excel's Application.ActiveWindow.PointsToScreenPixelsX/Y). Unlike
        /// GetWindowDpi, this doesn't need an existing window/HWND - useful for positioning a window
        /// at an explicit screen location before it's shown (and therefore before it has a monitor of
        /// its own to query), matching the correct DPI even when that monitor differs from the primary
        /// one in a mixed-DPI multi-monitor setup.
        /// </summary>
        public static double GetDpiForScreenPoint(double x, double y)
        {
            try
            {
                var pt = new POINT { X = (int)Math.Round(x), Y = (int)Math.Round(y) };
                IntPtr monitor = MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST);
                if (monitor != IntPtr.Zero && GetDpiForMonitor(monitor, MDT_EFFECTIVE_DPI, out uint dpiX, out uint dpiY) == 0)
                {
                    return dpiX;
                }
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort DPI-for-point lookup; falls back to 96 (100% scaling).
                LogUtility.LogDebug($"{nameof(GetDpiForScreenPoint)}: DPI lookup for point ({x},{y}) failed, defaulting to 96 - {ex.Message}");
            }

            return 96.0;
        }

        private class DpiContextDisposer : IDisposable
        {
            private readonly IntPtr _oldContext;
            private bool _disposed = false;

            public DpiContextDisposer(IntPtr oldContext)
            {
                _oldContext = oldContext;
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (_disposed)
                    return;

                if (disposing)
                {
                    // Dispose managed resources here (none in this case)
                }

                // Dispose unmanaged resources
                if (_oldContext != IntPtr.Zero)
                {
                    try
                    {
                        SetThreadDpiAwarenessContext(_oldContext);
                    }
                    catch (Exception ex)
                    {
                        // Safe to ignore: best-effort restore of the prior DPI awareness context;
                        // cleanup-only code, not a functional failure.
                        LogUtility.LogDebug($"{nameof(DpiContextDisposer)}: failed to restore prior DPI awareness context - {ex.Message}");
                    }
                }

                _disposed = true;
            }

            // Note: No finalizer needed since we have no unmanaged resources to clean up
            // The IntPtr is just a handle we're restoring, not something we own
        }
    }
}
