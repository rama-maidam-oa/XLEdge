using System;
using System.Runtime.InteropServices;
using XLEdge.Addin.Core.Infrastructure;
using XLEdge.Contracts;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge
{
    /// <summary>
    /// Today's XLEdge session state (same members), for code ported into the core. Differences forced by the AppDomain
    /// boundary: Excel is the core's own RCW (CoreServices; a stale one is dropped and re-requested from the shell instead
    /// of searched for in the running-object table), DebugLogs is the shell's Debug toggle, and XLEdgePane is the core's
    /// content of the shell's pane (Stage 2e). State lives in the core's AppDomain, so a reload starts from defaults.
    /// </summary>
    public sealed class XLEdgeAppState
    {
        // Singleton
        private static readonly Lazy<XLEdgeAppState> _instance =
            new Lazy<XLEdgeAppState>(() => new XLEdgeAppState());

        public static XLEdgeAppState Instance => _instance.Value;

        private readonly object _excelSyncRoot = new object();

        private XLEdgeAppState()
        {
        }

        internal Excel.Application ExcelAppUnsafe
        {
            get
            {
                lock (_excelSyncRoot)
                {
                    // Never creates one: today's callers (the shutdown sequence) only want an existing reference.
                    return CoreServices.ExcelAppIfCreated as Excel.Application;
                }
            }
        }

        // Excel & Add-in
        /// <summary>The Excel Application, or null. Setting it is accepted for today's callers but changes nothing:
        /// the core always works on the Excel the shell runs in.</summary>
        public Excel.Application ExcelApp
        {
            get
            {
                lock (_excelSyncRoot)
                {
                    Excel.Application app = CoreServices.ExcelApp as Excel.Application;
                    if (app == null) return null;
                    switch (Probe(app))
                    {
                        case ExcelProbe.Usable:
                            return app;
                        case ExcelProbe.Disconnected:
                            // Dropped (not released: callers may hold it) and asked for again from the shell.
                            CoreServices.ResetExcel();
                            app = CoreServices.ExcelApp as Excel.Application;
                            return app != null && Probe(app) == ExcelProbe.Usable ? app : null;
                        default:
                            // Busy (cell edit, dialog) or no window: today's TryUseApplication answer - no usable Excel now.
                            return null;
                    }
                }
            }
            set
            {
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

        /// <summary>Today's call at add-in start; the core's Excel comes from the shell, so there is nothing to set.</summary>
        public void InitializeExcelApplication(Excel.Application excelApp)
        {
        }

        public void EnsureExcelApplication()
        {
            _ = ExcelApp;
        }

        public void ResetExcelApplicationIfInvalid()
        {
            _ = ExcelApp;
        }

        /// <summary>
        /// At shutdown, after today's caller released the RCW itself: stops using it without releasing it again (CoreServices
        /// releases whatever is left at the core's Shutdown).
        /// </summary>
        internal void ClearExcelApplicationUnsafe()
        {
            lock (_excelSyncRoot)
            {
                CoreServices.ResetExcel();
            }
        }

        private enum ExcelProbe { Usable, Unavailable, Disconnected }

        /// <summary>Today's ExcelApplicationHelper.TryUseApplication probe (a live Excel has a window handle), telling a
        /// disconnected RCW apart from a busy Excel.</summary>
        private static ExcelProbe Probe(Excel.Application app)
        {
            try
            {
                return app.Hwnd != 0 ? ExcelProbe.Usable : ExcelProbe.Unavailable;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return CoreServices.IsDisconnected(ex) ? ExcelProbe.Disconnected : ExcelProbe.Unavailable;
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

        /// <summary>The active Excel window's pane as XLEdgeLogin last found it (11.1.2's member).</summary>
        public ADXExcelTaskPane1 XLEdgePane { get; set; }
        /// <summary>
        /// The shell's Debug toggle (IXLEdgeContext.DebugMode), so the ribbon button, the shell's logger and ported code
        /// always agree. False when the core is not attached.
        /// </summary>
        public bool DebugLogs
        {
            get
            {
                try { return CoreServices.Context?.DebugMode ?? false; }
                catch (Exception) { return false; }
            }
            set
            {
                try
                {
                    IXLEdgeContext context = CoreServices.Context;
                    if (context != null) context.DebugMode = value;
                }
                catch (Exception) { /* the shell is gone (closing session) */ }
            }
        }
        /// <summary>When true, the report engine includes response payload contents in debug logging.</summary>
        public bool DebugOutputData { get; set; }
        /// <summary>Set by UpdateTabLabel based on whether the active sheet has a valid,
        /// refreshable XLEdge table.</summary>
        public bool RefreshAll { get; set; }
        public bool ParamDataSameSheet { get; set; } = true;
        public bool SchOutputsToSameSheet { get; set; }
        public bool RefreshSync { get; set; }
        public bool AllowSheetNameChanges { get; set; }
        public bool ShowCalendarControl { get; set; }
        /// <summary>Defaults to true; the JSON preferences backfill in
        /// XLEdgePreferencesManager only applies this default when the entry is missing from
        /// the user's saved preferences file, so existing users who already have this set to
        /// false keep their choice.</summary>
        public bool ShowSegmentSelectionWindow { get; set; } = true;
        /// <summary>Backs the "Preserve user applied formats on refresh or run" option: true means
        /// preserve (skip re-stamping format), false means apply the report's default format.
        /// Renamed from OverrideFormats (was inverted from its own name) to match VB.NET's
        /// PreserveFormats rename (OISR-22179) - this is a pre-delivery product name/JSON-key change,
        /// not just a code-side rename, so the preferences file key changes too (see
        /// XLEdgeUserPreferences.PreserveFormats). Defaults to false, matching VB.NET.</summary>
        public bool PreserveFormats { get; set; } = false;

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
