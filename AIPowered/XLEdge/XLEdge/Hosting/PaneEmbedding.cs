using System;
using System.Runtime.InteropServices;

namespace XLEdge.Hosting
{
    /// <summary>
    /// Turns the core's content window (a top-level window of the core's AppDomain; only its handle crosses) into a child of
    /// the task pane and keeps it sized to it. A dead or zero handle is ignored.
    /// </summary>
    internal static class PaneEmbedding
    {
        private const int GWL_STYLE = -16;
        private const long WS_CHILD = 0x40000000L;
        private const long WS_POPUP = 0x80000000L;
        private const long WS_CAPTION = 0x00C00000L;
        private const long WS_THICKFRAME = 0x00040000L;
        private const long WS_SYSMENU = 0x00080000L;
        private const long WS_MINIMIZEBOX = 0x00020000L;
        private const long WS_MAXIMIZEBOX = 0x00010000L;

        public static bool IsAlive(IntPtr content) => content != IntPtr.Zero && IsWindow(content);

        /// <summary>Re-parents <paramref name="content"/> into <paramref name="parent"/> as a borderless child of that size.</summary>
        public static bool Embed(IntPtr content, IntPtr parent, int width, int height) => Embed(content, parent, width, height, out _);

        /// <summary>
        /// As <see cref="Embed(IntPtr, IntPtr, int, int)"/>; on failure the content keeps its original styles (never a parentless
        /// WS_CHILD) and <paramref name="win32Error"/> holds the error of the refused SetParent - typically ERROR_INVALID_STATE
        /// when the two windows have different DPI awareness (see <see cref="CreateUnderDpiOf"/>).
        /// </summary>
        public static bool Embed(IntPtr content, IntPtr parent, int width, int height, out int win32Error)
        {
            win32Error = 0;
            if (!IsAlive(content) || !IsAlive(parent)) return false;
            long original = GetWindowLongPtr(content, GWL_STYLE).ToInt64();
            long style = original & ~(WS_POPUP | WS_CAPTION | WS_THICKFRAME | WS_SYSMENU | WS_MAXIMIZEBOX | WS_MINIMIZEBOX);
            style |= WS_CHILD;
            SetWindowLongPtr(content, GWL_STYLE, new IntPtr(style));
            SetParent(content, parent);
            if (GetParent(content) != parent)
            {
                win32Error = Marshal.GetLastWin32Error();
                if (win32Error == 0) win32Error = -1;
                SetWindowLongPtr(content, GWL_STYLE, new IntPtr(original));
                return false;
            }
            Resize(content, width, height);
            return true;
        }

        /// <summary>
        /// Runs <paramref name="create"/> (the call that makes the core create its content window) under the DPI awareness
        /// context of <paramref name="pane"/>, so the new window can be parented into it: Office switches the thread's context,
        /// and Windows refuses SetParent between windows of different awareness. Before Windows 10 1607 it just runs it.
        /// </summary>
        public static IntPtr CreateUnderDpiOf(IntPtr pane, Func<IntPtr> create)
        {
            IntPtr previous = IntPtr.Zero;
            try
            {
                IntPtr context = IsAlive(pane) ? GetWindowDpiAwarenessContext(pane) : IntPtr.Zero;
                if (context != IntPtr.Zero) previous = SetThreadDpiAwarenessContext(context);
            }
            catch (EntryPointNotFoundException)
            {
                previous = IntPtr.Zero;
            }
            try
            {
                return create();
            }
            finally
            {
                if (previous != IntPtr.Zero) SetThreadDpiAwarenessContext(previous);
            }
        }

        public static void Resize(IntPtr content, int width, int height)
        {
            if (IsAlive(content)) MoveWindow(content, 0, 0, Math.Max(0, width), Math.Max(0, height), true);
        }

        /// <summary>
        /// Takes the content out of the pane (before the core that owns it is unloaded). Hidden first: a child of the desktop
        /// would otherwise show at the screen origin until the core destroys it.
        /// </summary>
        public static void Release(IntPtr content)
        {
            if (!IsAlive(content)) return;
            ShowWindow(content, SW_HIDE);
            SetParent(content, IntPtr.Zero);
        }

        private const int SW_HIDE = 0;

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindow(IntPtr hWnd);

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

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDpiAwarenessContext(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
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
