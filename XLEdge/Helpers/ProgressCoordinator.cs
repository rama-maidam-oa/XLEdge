using System;
using System.Runtime.InteropServices;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Resets shared drilldown/report state after an operation completes: clears
    /// <see cref="XLEdgeAppState"/> fields, hides the scratch "DummySheet" if visible, cleans up
    /// leftover temp report files, and refreshes the Excel tab label.
    /// </summary>
    public static class ProgressCoordinator
    {
        public static void ResetReportState()
        {
            try
            {
                var state = XLEdgeAppState.Instance;

                state.ActiveWorkbookOverride = null;
                state.FollowDrilldown = false;
                state.DrillPostData = string.Empty;
                state.ChildTableName = string.Empty;
                state.ChildShtName = string.Empty;
                state.ChildRptLabel = string.Empty;

                // Clears out this run's (and any other leftover) temp report CSVs now that the
                // operation has completed.
                XLEdgeTempFileCleaner.DeleteAllTempFiles();

                HideDummySheetIfVisible();

                try
                {
                    if (state.ActiveWorksheetOverride != null)
                    {
                        state.ActiveWorksheetOverride.Activate();
                        state.ActiveWorksheetOverride = null;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(ResetReportState));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ResetReportState));
            }
            finally
            {
                try
                {
                    if (ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelAppForLabel))
                    {
                        AddinModule.CurrentInstance?.UpdateTabLabel(excelAppForLabel.ActiveSheet as Excel.Worksheet);
                    }

                    XLEdgeAppState.Instance.ProcessRunning = false;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(ResetReportState));
                }
            }
        }

        private static void HideDummySheetIfVisible()
        {
            if (!ExcelSheetHelper.SheetExists("DummySheet"))
            {
                return;
            }

            Excel.Worksheet dummySheet = null;
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                dummySheet = (Excel.Worksheet)excelApp.ActiveWorkbook.Worksheets["DummySheet"];

                if (dummySheet.Visible == Excel.XlSheetVisibility.xlSheetVisible)
                {
                    dummySheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(HideDummySheetIfVisible));
            }
            finally
            {
                if (dummySheet != null)
                {
                    Marshal.ReleaseComObject(dummySheet);
                }
            }
        }
    }
}
