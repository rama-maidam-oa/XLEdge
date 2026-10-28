using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace XLEdge.Views
{
    /// <summary>
    /// Removes the minimize/maximize buttons from a host window's title bar. WPF has no property for this
    /// (ResizeMode="NoResize" only greys them out), so the WS_MINIMIZEBOX / WS_MAXIMIZEBOX bits are cleared
    /// from the native window style once the HWND exists.
    /// </summary>
    internal static class WindowChromeHelper
    {
        private const int GWL_STYLE = -16;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int nIndex) =>
            IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, nIndex) : new IntPtr(GetWindowLong32(hWnd, nIndex));

        private static void SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
        {
            if (IntPtr.Size == 8) SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
            else SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32());
        }

        /// <summary>
        /// Call from the window's SourceInitialized handler (the HWND must exist). Leaves the title, icon, Close button and
        /// resize border alone. Never throws: the tweak is cosmetic and must not stop the window from opening.
        /// </summary>
        public static void RemoveMinimizeMaximizeButtons(Window window)
        {
            try
            {
                IntPtr hwnd = new WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero) return;

                long style = GetWindowLongPtr(hwnd, GWL_STYLE).ToInt64();
                style &= ~(WS_MINIMIZEBOX | WS_MAXIMIZEBOX);
                SetWindowLongPtr(hwnd, GWL_STYLE, new IntPtr(style));
            }
            catch
            {
                // Cosmetic only.
            }
        }
    }
}
