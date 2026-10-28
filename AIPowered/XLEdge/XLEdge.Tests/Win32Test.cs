using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace XLEdge.Tests
{
    /// <summary>Runs window tests on an STA thread with a message pump, and reads window state through Win32.</summary>
    internal static class Win32Test
    {
        public const int GWL_STYLE = -16;
        public const int GWL_EXSTYLE = -20;
        public const long WS_CHILD = 0x40000000L;
        public const long WS_POPUP = 0x80000000L;
        public const long WS_CAPTION = 0x00C00000L;
        public const long WS_THICKFRAME = 0x00040000L;
        public const long WS_EX_TOOLWINDOW = 0x00000080L;
        public const long WS_EX_APPWINDOW = 0x00040000L;

        public static void OnSta(Action test)
        {
            ExceptionDispatchInfo failure = null;
            var thread = new Thread(() =>
            {
                try { test(); }
                catch (Exception ex) { failure = ExceptionDispatchInfo.Capture(ex); }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
            failure?.Throw();
        }

        /// <summary>Pumps this thread's messages until <paramref name="done"/> is true or the timeout passes; returns <paramref name="done"/>.</summary>
        public static bool PumpUntil(Func<bool> done, TimeSpan timeout)
        {
            var clock = Stopwatch.StartNew();
            while (!done())
            {
                if (clock.Elapsed > timeout) return false;
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(15);
            }
            return true;
        }

        public static void Pump(TimeSpan duration) => PumpUntil(() => false, duration);

        public static long Style(IntPtr hwnd) => GetWindowLongPtr(hwnd, GWL_STYLE).ToInt64();
        public static long ExStyle(IntPtr hwnd) => GetWindowLongPtr(hwnd, GWL_EXSTYLE).ToInt64();

        public static (int Width, int Height) ClientSize(IntPtr hwnd)
        {
            GetClientRect(hwnd, out RECT r);
            return (r.Right - r.Left, r.Bottom - r.Top);
        }

        public static (int Left, int Top) ScreenPosition(IntPtr hwnd)
        {
            GetWindowRect(hwnd, out RECT r);
            return (r.Left, r.Top);
        }

        /// <summary>Every window of this thread, top-level and child.</summary>
        public static HashSet<IntPtr> ThreadWindows()
        {
            var all = new HashSet<IntPtr>();
            EnumThreadWindows(GetCurrentThreadId(), (top, _) =>
            {
                all.Add(top);
                EnumChildWindows(top, (child, __) => { all.Add(child); return true; }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            return all;
        }

        public static string ClassName(IntPtr hwnd)
        {
            var text = new System.Text.StringBuilder(256);
            GetClassName(hwnd, text, text.Capacity);
            return text.ToString();
        }

        private delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll")] public static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll")] public static extern bool DestroyWindow(IntPtr hWnd);
        [DllImport("kernel32.dll")] private static extern int GetCurrentThreadId();
        [DllImport("user32.dll")] private static extern bool EnumThreadWindows(int threadId, EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder name, int max);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")] private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int index);
        [DllImport("user32.dll", EntryPoint = "GetWindowLong")] private static extern int GetWindowLong32(IntPtr hWnd, int index);

        private static IntPtr GetWindowLongPtr(IntPtr hWnd, int index) =>
            IntPtr.Size == 8 ? GetWindowLongPtr64(hWnd, index) : new IntPtr(GetWindowLong32(hWnd, index));
    }
}
