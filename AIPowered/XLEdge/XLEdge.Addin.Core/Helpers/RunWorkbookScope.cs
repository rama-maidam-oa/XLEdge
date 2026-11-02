using System;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Twin (Stage 4, the developer's request): the workbook a report or refresh run started in. 11.1.2 reads ActiveWorkbook
    /// after its HTTP awaits, so a user switching workbooks meanwhile got the report in the other one; 11.1.2 already has the
    /// field for it (XLEdgeAppState.ActiveWorkbookOverride), never set. The first scope of a run sets it to the workbook active
    /// now and owns it; nested scopes (each report of a multi-report run, each sheet of Refresh All) keep the outer one's.
    /// Read through ExcelApplicationHelper.RunWorkbook.
    /// </summary>
    internal sealed class RunWorkbookScope : IDisposable
    {
        private bool _owner;

        public static RunWorkbookScope Begin()
        {
            var scope = new RunWorkbookScope();
            try
            {
                XLEdgeAppState state = XLEdgeAppState.Instance;
                if (state.ActiveWorkbookOverride == null)
                {
                    Excel.Workbook active = ExcelApplicationHelper.GetActiveExcelApplication()?.ActiveWorkbook;
                    if (active != null)
                    {
                        state.ActiveWorkbookOverride = active;
                        scope._owner = true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn("RunWorkbookScope: the active workbook could not be read - this run follows the active workbook: " + ex.Message);
            }
            return scope;
        }

        public void Dispose()
        {
            if (!_owner) return;
            _owner = false;
            XLEdgeAppState.Instance.ActiveWorkbookOverride = null;
        }
    }
}
