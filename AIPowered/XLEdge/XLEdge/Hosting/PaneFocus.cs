using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace XLEdge.Hosting
{
    /// <summary>
    /// Gives the keyboard back to the worksheet when the user clicks a cell while it is in a task pane. The pane's WebView2
    /// takes the focus into a window of its browser process, and a single click on the grid does not take it back (Excel
    /// only does on entering edit mode), so keystrokes kept going to the pane. Called on SheetSelectionChange.
    /// </summary>
    internal static class PaneFocus
    {
        private const string WorksheetGridClass = "EXCEL7";

        /// <summary>On SheetSelectionChange: returns the focus to the grid under the mouse when it is in one of the panes.</summary>
        public static bool ReturnToGridAfterClick(IEnumerable<IntPtr> panes)
        {
            if (!GetCursorPos(out POINT cursor)) return false;
            return ReturnToGrid(panes, WindowFromPoint(cursor), IsWorksheetGrid);
        }

        /// <summary>
        /// Moves the focus to <paramref name="windowUnderCursor"/> only when it is a worksheet grid (the selection changed
        /// because the user clicked it) and this thread's focus is in one of <paramref name="panes"/> or a window inside
        /// one (the WebView2's browser window is a child of the pane, owned by another process). Otherwise - code selecting
        /// cells while the user types in the pane, focus already elsewhere - nothing changes.
        /// </summary>
        public static bool ReturnToGrid(IEnumerable<IntPtr> panes, IntPtr windowUnderCursor, Func<IntPtr, bool> isGrid)
        {
            if (windowUnderCursor == IntPtr.Zero || !isGrid(windowUnderCursor)) return false;
            IntPtr focus = FocusOfThisThread();
            if (focus == IntPtr.Zero) return false;
            if (!(panes ?? Enumerable.Empty<IntPtr>()).Any(pane => pane != IntPtr.Zero && (focus == pane || IsChild(pane, focus)))) return false;
            SetFocus(windowUnderCursor);
            return true;
        }

        public static bool IsWorksheetGrid(IntPtr window)
        {
            if (window == IntPtr.Zero) return false;
            var name = new StringBuilder(64);
            return GetClassName(window, name, name.Capacity) > 0 && name.ToString() == WorksheetGridClass;
        }

        /// <summary>The focus of this GUI thread, including a window of another process attached to its input (WebView2).</summary>
        private static IntPtr FocusOfThisThread()
        {
            var info = new GUITHREADINFO { cbSize = Marshal.SizeOf(typeof(GUITHREADINFO)) };
            return GetGUIThreadInfo(GetCurrentThreadId(), ref info) ? info.hwndFocus : IntPtr.Zero;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X, Y; }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        private struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive, hwndFocus, hwndCapture, hwndMenuOwner, hwndMoveSize, hwndCaret;
            public RECT rcCaret;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsChild(IntPtr parent, IntPtr window);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr window);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr window, StringBuilder name, int max);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetGUIThreadInfo(int threadId, ref GUITHREADINFO info);

        [DllImport("kernel32.dll")]
        private static extern int GetCurrentThreadId();
    }
}
