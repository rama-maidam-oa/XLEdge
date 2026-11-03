using System;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    public static class ExcelApplicationHelper
    {
        /// <summary>Twin (Stage 4): the workbook the running report/refresh started in (RunWorkbookScope), else the active one.</summary>
        public static Excel.Workbook RunWorkbook(Excel.Application app) =>
            RunWorkbookScope.CurrentWorkbook as Excel.Workbook ?? app?.ActiveWorkbook;

        /// <summary>
        /// Twin (Stage 4): activates <paramref name="sheet"/>, bringing its workbook to the front first - a run may be writing to a
        /// workbook the user switched away from, and Excel cannot activate a sheet of an inactive workbook.
        /// </summary>
        public static void ActivateInItsWorkbook(Excel.Worksheet sheet)
        {
            if (sheet == null) return;
            try
            {
                // Only when the user switched away: activating the already-active workbook would jump to its first window.
                if (sheet.Parent is Excel.Workbook book && NeedsWorkbookActivation(sheet.Application.ActiveWorkbook?.FullName, book.FullName))
                    book.Activate();
            }
            catch (Exception ex) { LogUtility.LogWarn("ExcelApplicationHelper.ActivateInItsWorkbook: its workbook could not be activated: " + ex.Message); }
            sheet.Activate();
        }

        /// <summary>True when the sheet's workbook is not the active one (full names, case-insensitive as Windows paths).</summary>
        internal static bool NeedsWorkbookActivation(string activeWorkbookFullName, string sheetWorkbookFullName) =>
            !string.Equals(activeWorkbookFullName, sheetWorkbookFullName, StringComparison.OrdinalIgnoreCase);

        public static Excel.Application GetActiveExcelApplication()
        {
            TryGetActiveExcelApplication(out Excel.Application excelApp);
            return excelApp;
        }

        public static bool TryGetActiveExcelApplication(out Excel.Application excelApp)
        {
            if (TryUseApplication(XLEdgeAppState.Instance.ExcelAppUnsafe, out excelApp))
            {
                return true;
            }

            if (TryGetActiveExcelApplicationInternal(out excelApp))
            {
                XLEdgeAppState.Instance.InitializeExcelApplication(excelApp);
                return true;
            }

            excelApp = null;
            return false;
        }

        // Core: the Excel the add-in runs in (the core's own RCW, see CoreServices). XLEdge 11.1.2 fell back to
        // AddinModule.HostApplication and then to Marshal.GetActiveObject("Excel.Application"), which can return another
        // Excel process when two are open.
        internal static bool TryGetActiveExcelApplicationInternal(out Excel.Application excelApp)
        {
            excelApp = null;

            try
            {
                if (TryUseApplication(XLEdgeAppState.Instance.ExcelApp, out excelApp))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }

            excelApp = null;
            return false;
        }

        public static IntPtr GetExcelWindowHandle()
        {
            return TryGetActiveExcelApplication(out Excel.Application excelApp)
                ? new IntPtr(excelApp.Hwnd)
                : IntPtr.Zero;
        }

        /// <summary>
        /// Resolves the active Excel Application, throwing if none is available.
        /// </summary>
        public static Excel.Application RequireActiveExcelApplication()
        {
            if (TryGetActiveExcelApplication(out Excel.Application excelApp))
            {
                return excelApp;
            }

            throw new InvalidOperationException("No active Excel application instance could be resolved.");
        }

        /// <summary>
        /// Detects whether the active worksheet is currently in cell-edit mode by toggling
        /// Application.Interactive. Excel throws a COMException while a cell is being edited,
        /// which this treats as "in edit mode". Interactive is always restored to true before
        /// returning, including via a finally block, so this never leaves Excel's input handling
        /// disabled.
        /// </summary>
        public static bool IsCellInEditMode()
        {
            Excel.Application excelApp;
            try
            {
                excelApp = RequireActiveExcelApplication();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(IsCellInEditMode));
                return false;
            }

            if (!excelApp.Interactive)
            {
                // Already stuck/disabled from a previous run of this check (or something else) -
                // don't report "in edit mode" for this, but do try to self-heal it here too.
                try
                {
                    excelApp.Interactive = true;
                }
                catch (Exception ex)
                {
                    // Safe to ignore/expected: genuinely still in edit mode right now - leave it,
                    // caller treats this as "not confirmed in edit mode" same as before; a later
                    // call will retry.
                    LogUtility.LogDebug($"{nameof(IsCellInEditMode)}: self-heal of Interactive=true failed (likely still in cell edit mode) - {ex.Message}");
                }

                return false;
            }

            bool inEditMode;
            try
            {
                excelApp.Interactive = false;
                excelApp.Interactive = true;
                inEditMode = false;
            }
            catch (Exception)
            {
                // Expected/routine: toggling Interactive fails while a cell is in edit mode - treat
                // as "in edit mode", not a real error.
                inEditMode = true;
            }
            finally
            {
                // Always make sure we leave Interactive back on. If Excel is genuinely still in
                // edit mode this will itself throw (harmless, swallowed) and self-correct on a
                // later call once the user exits edit mode.
                try
                {
                    if (!excelApp.Interactive)
                    {
                        excelApp.Interactive = true;
                    }
                }
                catch (Exception)
                {
                    // Safe to ignore/expected: still in edit mode - nothing more we can do right now.
                }
            }

            return inEditMode;
        }

        internal static bool TryUseApplication(Excel.Application excelApp, out Excel.Application resolvedApp)
        {
            resolvedApp = null;

            if (excelApp == null)
            {
                return false;
            }

            try
            {
                int hwnd = excelApp.Hwnd;
                if (hwnd != 0)
                {
                    resolvedApp = excelApp;
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }

            return false;
        }
    }
}