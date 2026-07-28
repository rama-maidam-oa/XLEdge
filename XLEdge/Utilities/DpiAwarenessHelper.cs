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

        [DllImport("user32.dll")]
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
