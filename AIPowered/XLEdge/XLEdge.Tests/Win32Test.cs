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

        /// <summary>Every window of this process (all threads), top-level and child, with its class name.</summary>
        public static Dictionary<IntPtr, string> ProcessWindows()
        {
            int pid = Process.GetCurrentProcess().Id;
            var all = new Dictionary<IntPtr, string>();
            EnumWindows((top, _) =>
            {
                GetWindowThreadProcessId(top, out int owner);
                if (owner != pid) return true;
                all[top] = ClassName(top);
                EnumChildWindows(top, (child, __) =>
                {
                    GetWindowThreadProcessId(child, out int childOwner);
                    if (childOwner == pid) all[child] = ClassName(child);
                    return true;
                }, IntPtr.Zero);
                return true;
            }, IntPtr.Zero);
            // Message-only windows (WinForms parking/marshaling windows, SystemEvents) are not enumerated above.
            IntPtr found = IntPtr.Zero;
            while ((found = FindWindowEx(HWND_MESSAGE, found, null, null)) != IntPtr.Zero)
            {
                GetWindowThreadProcessId(found, out int owner);
                if (owner == pid) all[found] = ClassName(found);
            }
            return all;
        }

        private static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);

        /// <summary>The native thread that owns the window, and whether it is the calling thread.</summary>
        public static string Owner(IntPtr hwnd)
        {
            int thread = GetWindowThreadProcessId(hwnd, out _);
            return thread == GetCurrentThreadId() ? $"thread {thread} (this thread)" : $"thread {thread}";
        }

        [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
        [DllImport("user32.dll")] private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr FindWindowEx(IntPtr parent, IntPtr after, string className, string windowName);

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
        [DllImport("user32.dll")] public static extern IntPtr GetFocus();

        /// <summary>A hidden window of a class registered here under <paramref name="className"/> (e.g. Excel's "EXCEL7").</summary>
        public static IntPtr CreateWindowOfClass(string className)
        {
            var wc = new WNDCLASS { lpfnWndProc = DefWindowProcPointer, hInstance = GetModuleHandle(null), lpszClassName = className };
            RegisterClass(ref wc); // fails harmlessly when already registered by an earlier test
            return CreateWindowEx(0, className, "", 0, 0, 0, 10, 10, IntPtr.Zero, IntPtr.Zero, wc.hInstance, IntPtr.Zero);
        }

        private static readonly IntPtr DefWindowProcPointer = GetProcAddress(GetModuleHandle("user32.dll"), "DefWindowProcW");

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WNDCLASS
        {
            public uint style;
            public IntPtr lpfnWndProc;
            public int cbClsExtra, cbWndExtra;
            public IntPtr hInstance, hIcon, hCursor, hbrBackground;
            public string lpszMenuName, lpszClassName;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern ushort RegisterClass(ref WNDCLASS wc);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr CreateWindowEx(int exStyle, string className, string title, int style, int x, int y, int w, int h,
            IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern IntPtr GetModuleHandle(string name);
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi)] private static extern IntPtr GetProcAddress(IntPtr module, string name);

        public static readonly IntPtr DPI_AWARENESS_CONTEXT_UNAWARE = new IntPtr(-1);
        public static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
        [DllImport("user32.dll")] public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
        [DllImport("user32.dll")] public static extern IntPtr GetThreadDpiAwarenessContext();
        [DllImport("user32.dll")] public static extern bool AreDpiAwarenessContextsEqual(IntPtr a, IntPtr b);
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
