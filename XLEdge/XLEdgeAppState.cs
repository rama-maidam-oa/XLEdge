using System;
using XLEdge.Helpers;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge
{
    public sealed class XLEdgeAppState
    {
        // Singleton
        private static readonly Lazy<XLEdgeAppState> _instance =
            new Lazy<XLEdgeAppState>(() => new XLEdgeAppState());

        public static XLEdgeAppState Instance => _instance.Value;

        private readonly object _excelSyncRoot = new object();
        private Excel.Application _excelApp;

        private XLEdgeAppState()
        {
        }
        internal Excel.Application ExcelAppUnsafe
        {
            get
            {
                lock (_excelSyncRoot)
                {
                    return _excelApp;
                }
            }
        }

        // Excel & Add-in
        public Excel.Application ExcelApp
        {
            get
            {
                lock (_excelSyncRoot)
                {
                    if (ExcelApplicationHelper.TryUseApplication(_excelApp, out Excel.Application validApp))
                    {
                        _excelApp = validApp;
                        return _excelApp;
                    }

                    if (ExcelApplicationHelper.TryGetActiveExcelApplicationInternal(out Excel.Application recoveredApp))
                    {
                        _excelApp = recoveredApp;
                        return _excelApp;
                    }

                    return null;
                }
            }
            set
            {
                lock (_excelSyncRoot)
                {
                    if (ExcelApplicationHelper.TryUseApplication(value, out Excel.Application validApp))
                    {
                        _excelApp = validApp;
                        return;
                    }

                    if (ExcelApplicationHelper.TryGetActiveExcelApplicationInternal(out Excel.Application recoveredApp))
                    {
                        _excelApp = recoveredApp;
                        return;
                    }

                    _excelApp = null;
                }
            }
        }
        public IntPtr ExcelHandle
        {
            get
            {
                Excel.Application app = ExcelApp;
                if (app == null)
                {
                    return IntPtr.Zero;
                }

                try
                {
                    return new IntPtr(app.Hwnd);
                }
                catch (Exception ex)
                {
                    // Safe to ignore: best-effort HWND lookup; treated as "no Excel window handle
                    // available" by callers.
                    LogUtility.LogDebug($"{nameof(ExcelHandle)}: failed to read Excel app HWND - {ex.Message}");
                    return IntPtr.Zero;
                }
            }

        }
        public void InitializeExcelApplication(Excel.Application excelApp)
        {
            lock (_excelSyncRoot)
            {
                if (ExcelApplicationHelper.TryUseApplication(excelApp, out Excel.Application validApp))
                {
                    _excelApp = validApp;
                    return;
                }

                if (ExcelApplicationHelper.TryGetActiveExcelApplicationInternal(out Excel.Application recoveredApp))
                {
                    _excelApp = recoveredApp;
                    return;
                }

                _excelApp = null;
            }
        }

        public void EnsureExcelApplication()
        {
            lock (_excelSyncRoot)
            {
                if (ExcelApplicationHelper.TryUseApplication(_excelApp, out Excel.Application validApp))
                {
                    _excelApp = validApp;
                    return;
                }

                if (ExcelApplicationHelper.TryGetActiveExcelApplicationInternal(out Excel.Application recoveredApp))
                {
                    _excelApp = recoveredApp;
                    return;
                }

                _excelApp = null;
            }
        }
        public void ResetExcelApplicationIfInvalid()
        {
            lock (_excelSyncRoot)
            {
                if (!ExcelApplicationHelper.TryUseApplication(_excelApp, out Excel.Application validApp))
                {
                    _excelApp = null;
                    EnsureExcelApplication();
                }
                else
                {
                    _excelApp = validApp;
                }
            }
        }
        public DateTime? GetDateFromCell(Excel.Range cell)
        {
            if (cell == null)
                return null;

            try
            {
                object value = cell.Value2;

                if (value == null)
                    return null;

                if (value is DateTime dateTime)
                    return dateTime.Date;

                if (value is double oaDate)
                {
                    try
                    {
                        return DateTime.FromOADate(oaDate).Date;
                    }
                    catch (Exception ex)
                    {
                        // Safe to ignore: the numeric value isn't a valid OLE Automation date; caller
                        // treats null as "no date in this cell".
                        LogUtility.LogDebug($"{nameof(GetDateFromCell)}: failed to convert OLE Automation date {oaDate} - {ex.Message}");
                        return null;
                    }
                }

                if (DateTime.TryParse(
                    Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture),
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None,
                    out DateTime parsed))
                {
                    return parsed.Date;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }

            return null;
        }
        public void WriteDateToCell(Excel.Range cell, DateTime dateValue)
        {
            if (cell == null)
                return;

            try
            {
                cell.NumberFormat = "dd-MMM-yyyy";
                cell.Value = dateValue.Date;
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error writing date to Excel cell: {ex.Message}");
            }
        }
        //XLEdge pane
        public ADXExcelTaskPane1 XLEdgePane {  get; set; }

        public string LoginUrl { get; set; }
        public string LoginUrlName { get; set; }
        public string LoginToken { get; set; }
        public string LoginUserName { get; set; }
        public bool LoginFromSense { get; set; }
        public bool IsLoginCompleted { get; set; }
        public bool EdgePaneShown { get; set; } = false;
        public bool LoginFromGLSense { get; set; } = false;
        /// <summary>Set once XLEdge has notified the sibling GLSense add-in of a login that originated
        /// in XLEdge itself, so a login isn't re-broadcast to GLSense on every subsequent WebView2
        /// navigation back to the "excel=Y#Home" landing page within the same session. See
        /// AddinModule.NotifyGLSenseOfLogin.</summary>
        public bool LoginSentToGLSense { get; set; } = false;
        public bool DebugLogs { get; set; }
        /// <summary>When true, the report engine includes response payload contents in debug logging.</summary>
        public bool DebugOutputData { get; set; }
        /// <summary>Set by UpdateTabLabel based on whether the active sheet has a valid,
        /// refreshable XLEdge table.</summary>
        public bool RefreshAll { get; set; }
        public bool ParamDataSameSheet { get; set; }
        public bool SchOutputsToSameSheet { get; set; }
        public bool RefreshSync { get; set; }
        public bool AllowSheetNameChanges { get; set; }
        public bool ShowCalendarControl { get; set; }
        public bool ShowSegmentSelectionWindow { get; set; }
        public bool OverrideFormats { get; set; }

        // --- Drilldown / report-navigation state ---
        public bool FollowDrilldown { get; set; }
        public string DrillPostData { get; set; } = string.Empty;
        public string ChildTableName { get; set; } = string.Empty;
        public string ChildShtName { get; set; } = string.Empty;
        public string ChildRptLabel { get; set; } = string.Empty;

        /// <summary>
        /// Sticky workbook reference used while a report/progress operation is in flight.
        /// </summary>
        public Excel.Workbook ActiveWorkbookOverride { get; set; }

        /// <summary>
        /// Sticky worksheet reference to reactivate once a long-running operation completes.
        /// </summary>
        public Excel.Worksheet ActiveWorksheetOverride { get; set; }

        /// <summary>Indicates whether a long-running process is currently in progress.</summary>
        public bool ProcessRunning { get; set; }

        // --- Cached param/meta data from control sheet edits ---
        // Set by BuildRefreshParamsPayload and consumed by RefreshListObjectAsync to avoid
        // making an extra API call when control-sheet edits are present.

        private string _updatedParamData;
        private string _updatedMetaData;

        /// <summary>
        /// Cached parameter data from control sheet edits.
        /// Set by BuildRefreshParamsPayload before a refresh, cleared after use.
        /// </summary>
        public string UpdatedParamData
        {
            get => _updatedParamData;
            set => _updatedParamData = value;
        }

        /// <summary>
        /// Cached meta data from control sheet edits.
        /// Set by BuildRefreshParamsPayload before a refresh, cleared after use.
        /// </summary>
        public string UpdatedMetaData
        {
            get => _updatedMetaData;
            set => _updatedMetaData = value;
        }

        /// <summary>
        /// Clears both cached param and meta data. Called after a refresh completes
        /// to prevent stale data from being reused for subsequent refreshes.
        /// </summary>
        public void ClearCachedRefreshData()
        {
            _updatedParamData = null;
            _updatedMetaData = null;
        }
    }
}
