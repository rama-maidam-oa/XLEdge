using System;
using System.Threading;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Twin (Stage 4, the developer's request): the workbook a report or refresh run started in. 11.1.2 reads ActiveWorkbook
    /// after its HTTP awaits, so a user switching workbooks meanwhile got the report in the other one. The first scope of a run
    /// takes the workbook active now and owns it; nested scopes (each report of a multi-report run, each sheet of Refresh All)
    /// keep the outer one's. Kept per run (AsyncLocal: it follows the run's own awaits and Task.Runs, not other runs or event
    /// handlers), so two runs at once keep their own workbooks and code outside a run sees none. Read through
    /// ExcelApplicationHelper.RunWorkbook.
    /// </summary>
    internal sealed class RunWorkbookScope : IDisposable
    {
        private static readonly AsyncLocal<object> Current = new AsyncLocal<object>();
        private bool _owner;

        /// <summary>The running run's workbook on this flow, or null outside a run.</summary>
        internal static object CurrentWorkbook => Current.Value;

        public static RunWorkbookScope Begin() => Begin(() => ExcelApplicationHelper.GetActiveExcelApplication()?.ActiveWorkbook);

        /// <summary><paramref name="activeWorkbook"/>: the workbook active now (read only when no run is open on this flow).</summary>
        internal static RunWorkbookScope Begin(Func<object> activeWorkbook)
        {
            var scope = new RunWorkbookScope();
            try
            {
                if (Current.Value == null)
                {
                    object active = activeWorkbook();
                    if (active != null)
                    {
                        Current.Value = active;
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
            Current.Value = null;
        }
    }
}
