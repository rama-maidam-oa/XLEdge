using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    public class ExcelWindowHelper
    {
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, ref RECT rectangle);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        [DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool LockSetForegroundWindow(uint uLockCode);

        private const int SW_RESTORE = 9;
        private const uint GW_CHILD = 5;
        private const uint GW_HWNDNEXT = 2;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_SHOWWINDOW = 0x0040;
        private const uint LSFW_LOCK = 1;
        private const uint LSFW_UNLOCK = 2;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// Centers the WPF window exactly over the Excel worksheet area, 
        /// ensuring it doesn't go above the formula bar
        /// </summary>
        public static void CenterWindowOverExcel(Window window)
        {
            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
            {
                throw new InvalidOperationException("Excel application instance is not available.");
            }

            try
            {
                var dpiScale = GetWindowScale(window);

                // Get Excel main window handle
                IntPtr excelHandle = new IntPtr(excelApp.Hwnd);

                // Get Excel window rectangle
                RECT excelRect = new RECT();
                GetWindowRect(excelHandle, ref excelRect);

                // Estimate the top of the worksheet area (below formula bar and column headers)
                int titleBarHeight = 30;
                int menuBarHeight = 25;
                int formulaBarHeight = 30;
                int columnHeadersHeight = 25;

                double worksheetTop = (excelRect.Top + titleBarHeight + menuBarHeight + formulaBarHeight + columnHeadersHeight) / dpiScale;

                int statusBarHeight = 25;
                double worksheetBottom = (excelRect.Bottom - statusBarHeight) / dpiScale;
                double excelLeft = excelRect.Left / dpiScale;
                double excelRight = excelRect.Right / dpiScale;
                double excelWidth = excelRight - excelLeft;

                try
                {
                    var windowWidth = GetWindowDimension(window.ActualWidth, window.Width);
                    var windowHeight = GetWindowDimension(window.ActualHeight, window.Height);

                    Excel.Range activeCell = excelApp.ActiveCell;
                    if (activeCell != null)
                    {
                        Excel.Window activeWindow = excelApp.ActiveWindow;
                        double cellLeft = activeWindow.PointsToScreenPixelsX((int)activeCell.Left) / dpiScale;
                        double cellTop = activeWindow.PointsToScreenPixelsY((int)activeCell.Top) / dpiScale;
                        double cellWidth = activeWindow.PointsToScreenPixelsX((int)activeCell.Width) / dpiScale;
                        double cellHeight = activeWindow.PointsToScreenPixelsY((int)activeCell.Height) / dpiScale;

                        double centerX = cellLeft + (cellWidth / 2);
                        double centerY = cellTop + (cellHeight / 2);

                        window.Left = centerX - (windowWidth / 2);
                        window.Top = centerY - (windowHeight / 2);
                    }
                    else
                    {
                        window.Left = excelLeft + (excelWidth - windowWidth) / 2;
                        window.Top = worksheetTop + ((worksheetBottom - worksheetTop) - windowHeight) / 2;
                    }
                }
                catch (Exception cellCenterEx)
                {
                    LogUtility.LogDebug($"{nameof(CenterWindowOverExcel)}: failed to center over active cell, falling back to Excel-window centering - {cellCenterEx.Message}");
                    var windowWidth = GetWindowDimension(window.ActualWidth, window.Width);
                    var windowHeight = GetWindowDimension(window.ActualHeight, window.Height);
                    window.Left = excelLeft + (excelWidth - windowWidth) / 2;
                    window.Top = worksheetTop + ((worksheetBottom - worksheetTop) - windowHeight) / 2;
                }

                window.WindowStartupLocation = WindowStartupLocation.Manual;
                window.Top = Math.Max(window.Top, worksheetTop);

                if (window.Top + window.Height > worksheetBottom)
                {
                    window.Top = worksheetBottom - window.Height;
                }

                window.Left = Math.Max(window.Left, excelLeft);
                if (window.Left + window.Width > excelRight)
                {
                    window.Left = excelRight - window.Width;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error positioning window: {ex.Message}");
                LogUtility.LogWarn($"{nameof(CenterWindowOverExcel)}: failed to position window over Excel, falling back to CenterScreen - {ex.Message}");
                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
            }
        }

        /// <summary>
        /// Attempts to bring the Excel main window to the foreground, restoring it if minimized, AND
        /// hand real Win32 keyboard focus back to the worksheet grid.
        /// </summary>
        public static void ActivateExcelMainWindow(Excel.Application excelApp = null)
        {
            try
            {
                // GetActiveExcelApplication()/.Hwnd are Excel COM RCW access, and this method is
                // called from several places (some off the WPF/Excel STA dispatcher thread) - marshal
                // just that COM read onto the correct thread. UiDispatcher.Run is a no-op passthrough
                // when already called from that thread, so this doesn't change timing for the common case.
                IntPtr excelHandle = IntPtr.Zero;
                UiDispatcher.Run(() =>
                {
                    excelApp ??= ExcelApplicationHelper.GetActiveExcelApplication();
                    if (excelApp != null)
                    {
                        excelHandle = new IntPtr(excelApp.Hwnd);
                    }
                });
                if (excelApp == null || excelHandle == IntPtr.Zero) return;

                // Restore if minimized
                if (IsIconic(excelHandle))
                {
                    ShowWindow(excelHandle, SW_RESTORE);
                }

                // Set foreground window - this is the Windows API way to bring Excel to front
                SetForegroundWindow(excelHandle);
                System.Threading.Thread.Sleep(10);

                // Bring window to top
                BringWindowToTop(excelHandle);
                System.Threading.Thread.Sleep(10);

                // Set focus to the window
                SetFocus(excelHandle);
                System.Threading.Thread.Sleep(10);

                // Find the worksheet grid window and set focus to it
                IntPtr gridHwnd = FindWorksheetGridWindow(excelHandle);
                if (gridHwnd != IntPtr.Zero)
                {
                    SetFocus(gridHwnd);
                    System.Threading.Thread.Sleep(10);
                }

                // Deliberately not sending a dummy {F2}/{ESC} keystroke here (VB.NET's own
                // equivalent cleanup step has this same SendKeys call commented out). SendKeys goes
                // through the same low-level keyboard-injection path used to detect toggle-key state,
                // and was found to be flipping the user's NumLock on every report run. The
                // SetForegroundWindow/BringWindowToTop/SetFocus(gridHwnd) sequence above already gives
                // the worksheet grid real OS keyboard focus without it.
                LogUtility.LogDebug($"{nameof(ActivateExcelMainWindow)}: Successfully activated Excel window");
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"{nameof(ActivateExcelMainWindow)}: failed to activate Excel main window - {ex.Message}");
            }
        }

        /// <summary>
        /// Forces Excel window activation using multiple techniques
        /// This simulates the alt-tab behavior that fixes keyboard focus
        /// </summary>
        private static void ForceWindowActivation(IntPtr excelHandle, Excel.Application excelApp)
        {
            try
            {
                // Step 1: Get the foreground window
                IntPtr foregroundWindow = GetForegroundWindow();

                // Step 2: If Excel is already the foreground window, we still need to
                // force keyboard focus reset
                if (foregroundWindow == excelHandle || foregroundWindow == IntPtr.Zero)
                {
                    // Excel is already foreground, but keyboard focus might be lost
                    // We need to force a focus reset
                    ResetKeyboardFocus(excelApp);
                    return;
                }

                // Step 3: If a different window is foreground, use AttachThreadInput
                // to properly steal focus
                uint currentThreadId = GetWindowThreadProcessId(foregroundWindow, out _);
                uint excelThreadId = GetWindowThreadProcessId(excelHandle, out _);

                if (currentThreadId != excelThreadId && currentThreadId != 0 && excelThreadId != 0)
                {
                    // Attach threads to allow focus transfer
                    AttachThreadInput(currentThreadId, excelThreadId, true);

                    try
                    {
                        // Set foreground window
                        SetForegroundWindow(excelHandle);
                        System.Threading.Thread.Sleep(10);

                        // Bring window to top
                        BringWindowToTop(excelHandle);
                        System.Threading.Thread.Sleep(10);

                        // Set focus
                        SetFocus(excelHandle);
                        System.Threading.Thread.Sleep(10);
                    }
                    finally
                    {
                        // Detach threads
                        AttachThreadInput(currentThreadId, excelThreadId, false);
                    }
                }
                else
                {
                    // Same thread or we can't attach - use direct methods
                    SetForegroundWindow(excelHandle);
                    System.Threading.Thread.Sleep(10);
                    BringWindowToTop(excelHandle);
                    System.Threading.Thread.Sleep(10);
                    SetFocus(excelHandle);
                    System.Threading.Thread.Sleep(10);
                }

                // Step 4: Activate Excel COM window
                try
                {
                    excelApp.ActiveWindow?.Activate();
                }
                catch (Exception ex)
                {
                    LogUtility.LogError($"{nameof(ForceWindowActivation)}: ActiveWindow.Activate() failed - {ex.Message}");
                }

                // Step 5: Find and set focus to the worksheet grid
                IntPtr gridHwnd = FindWorksheetGridWindow(excelHandle);
                if (gridHwnd != IntPtr.Zero)
                {
                    SetFocus(gridHwnd);
                    System.Threading.Thread.Sleep(10);
                }

                // Step 6: Force a window position update (this triggers WM_ACTIVATE)
                SetWindowPos(excelHandle, IntPtr.Zero, 0, 0, 0, 0, SWP_NOSIZE | SWP_NOZORDER | SWP_SHOWWINDOW);
                System.Threading.Thread.Sleep(20);

                // Step 7: Reset keyboard focus using SendKeys
                ResetKeyboardFocus(excelApp);

                LogUtility.LogDebug($"{nameof(ForceWindowActivation)}: Completed window activation sequence");
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(ForceWindowActivation)}: Failed to force window activation - {ex.Message}");
                // Fallback: try simple method
                try
                {
                    SetForegroundWindow(excelHandle);
                    SetFocus(excelHandle);
                    ResetKeyboardFocus(excelApp);
                }
                catch (Exception innerEx)
                {
                    LogUtility.LogError($"{nameof(ForceWindowActivation)}: Fallback also failed - {innerEx.Message}");
                }
            }
        }

        /// <summary>
        /// Resets keyboard focus by forcing a real cell selection change (not simulated keystrokes).
        /// </summary>
        private static void ResetKeyboardFocus(Excel.Application excelApp)
        {
            try
            {
                if (excelApp == null) return;

                // The SendKeys {F2}/{ESC} and {DOWN}/{UP} "methods" that used to be tried here first
                // were removed - SendKeys/Application.SendKeys was found to be flipping the user's
                // NumLock state on every report run. The selection-change approach below (already
                // present as a third fallback) forces the same real keyboard-focus change through an
                // actual COM selection instead of a synthetic keystroke, without that side effect.

                // Force a selection change
                try
                {
                    if (excelApp.ActiveSheet != null && excelApp.ActiveCell != null)
                    {
                        var currentCell = excelApp.ActiveCell;
                        // Get a cell below and select it, then go back
                        Excel.Range targetCell = currentCell.Offset[1, 0];
                        if (targetCell != null)
                        {
                            targetCell.Select();
                            Thread.Sleep(20);
                            currentCell.Select();
                            Thread.Sleep(20);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogError($"{nameof(ResetKeyboardFocus)}: Selection change failed - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(ResetKeyboardFocus)}: Failed to reset keyboard focus - {ex.Message}");
            }
        }

        /// <summary>
        /// Finds the Excel worksheet grid window (the actual window that receives keyboard input)
        /// </summary>
        private static IntPtr FindWorksheetGridWindow(IntPtr parentHwnd)
        {
            try
            {
                string[] gridClassNames = new string[]
                {
                    "EXCEL7",
                    "XLDESK",
                    "Excel9",
                    "WorkbookWindow"
                };

                IntPtr mdiClient = FindWindowEx(parentHwnd, IntPtr.Zero, "MDIClient", null);
                if (mdiClient != IntPtr.Zero)
                {
                    IntPtr child = GetWindow(mdiClient, GW_CHILD);
                    while (child != IntPtr.Zero)
                    {
                        foreach (string className in gridClassNames)
                        {
                            System.Text.StringBuilder sb = new System.Text.StringBuilder(256);
                            GetClassName(child, sb, sb.Capacity);
                            if (sb.ToString().Contains(className))
                            {
                                IntPtr gridChild = FindWindowEx(child, IntPtr.Zero, "ExcelGrid", null);
                                if (gridChild != IntPtr.Zero)
                                {
                                    return gridChild;
                                }
                                return child;
                            }
                        }
                        child = GetWindow(child, GW_HWNDNEXT);
                    }
                }

                IntPtr current = GetWindow(parentHwnd, GW_CHILD);
                while (current != IntPtr.Zero)
                {
                    System.Text.StringBuilder sb = new System.Text.StringBuilder(256);
                    GetClassName(current, sb, sb.Capacity);
                    string className = sb.ToString();

                    foreach (string gridClass in gridClassNames)
                    {
                        if (className.Contains(gridClass))
                        {
                            return current;
                        }
                    }

                    current = GetWindow(current, GW_HWNDNEXT);
                }

                return IntPtr.Zero;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(FindWorksheetGridWindow)}: Failed to find grid window - {ex.Message}");
                return IntPtr.Zero;
            }
        }

        private static double GetWindowScale(Window window)
        {
            try
            {
                var source = System.Windows.PresentationSource.FromVisual(window);
                if (source?.CompositionTarget != null)
                {
                    var scale = source.CompositionTarget.TransformToDevice.M11;
                    if (scale > 0)
                        return scale;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(GetWindowScale)}: PresentationSource DPI lookup failed, trying fallback - {ex.Message}");
            }

            try
            {
                var dpi = System.Windows.Media.VisualTreeHelper.GetDpi(window);
                if (dpi.DpiScaleX > 0)
                    return dpi.DpiScaleX;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"{nameof(GetWindowScale)}: VisualTreeHelper.GetDpi lookup failed, defaulting to 1.0 - {ex.Message}");
            }

            return 1.0;
        }

        private static double GetWindowDimension(double actualValue, double fallbackValue)
        {
            if (actualValue > 0 && !double.IsNaN(actualValue))
                return actualValue;

            if (fallbackValue > 0 && !double.IsNaN(fallbackValue))
                return fallbackValue;

            return 0;
        }
    }
}