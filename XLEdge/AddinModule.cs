
using AddinExpress.MSO;
using MahApps.Metro.IconPacks;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using XLEdge.Helpers;
using XLEdge.Models;
using XLEdge.Utilities;
using XLEdge.Views;
using Excel = Microsoft.Office.Interop.Excel;


namespace XLEdge
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("80B0FB76-A5F4-41E0-B283-C175B7374B60"), ProgId("XLEdge.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public static XLEdgeRibbonHelper RibbonHelper { get; private set; }
        private static XLEdgeRibbonHelper _ribbonHelper;
        private static string _pendingRibbonState;
        public static NLog.Config.LoggingConfiguration LoggerConfiguration { get; set; }
        public static NLog.Logger Logger { get; set; }

        private bool _isCalendarOpen;
        private bool _isSegmentWindowOpen;

        // Cached reference to the sibling GLSense add-in's COM object, ported from VB's module-level
        // "addinInstance" field (resolved once via GetGLSenseAddinObject, then reused - e.g. by
        // InvokedFromGLSense and the GLSense session-sync call).
        private object _glSenseAddinInstance;

        // Deferred sheet-delete cleanup (ported from AddinModule.vb's "SheetsToDelete"/"DeleteTimer").
        // Excel does not allow certain operations (e.g. deleting a companion parameter sheet) from
        // directly inside SheetBeforeDelete without risking a crash, so the actual deletes are queued
        // and performed a moment later on a timer tick instead.
        private readonly List<string> _sheetsToDelete = new List<string>();
        private readonly System.Windows.Forms.Timer _deleteTimer = new System.Windows.Forms.Timer();

        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
            _deleteTimer.Interval = 200;
            _deleteTimer.Tick += DeleteTimer_Tick;
        }
 
        #region Add-in Express automatic code
 
        // Required by Add-in Express - do not modify
        // the methods within this region
 
        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }
 
        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }
 
        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }
 
        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance 
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        /// <summary>
        /// Ported from XLEdgeProcedures.vb's RibbonInitialize: resets the ribbon to its logged-out
        /// default state (Login enabled, every other report-action button disabled). Called by
        /// XLEdgeServerConfiguration (the Server Configuration window) after deleting an instance,
        /// matching VB's FormConfiguration.CmdDelete_Click - only when the ribbon is currently showing
        /// "Login" (i.e. the user isn't logged in), a no-op reset otherwise makes no visible difference.
        /// </summary>
        public void RibbonInitialize()
        {
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                if (excelApp == null)
                {
                    return;
                }

                RibEdgeLogin.Enabled = true;
                RibEdgeRefresh.Enabled = false;
                RibEdgeRefreshAll.Enabled = false;
                RibEdgeShowHide.Enabled = false;
                RibEdgeOptions.Enabled = false;
                RibEdgeHelp.Enabled = false;
                RibEdgeParamRefresh.Enabled = false;
                RibEdgeParamRefreshBook.Enabled = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RibbonInitialize));
            }
        }

        public Excel._Application ExcelApp
        {
            get
            {
                return (HostApplication as Excel._Application);
            }
        }
        public bool loginButtonVisibility()
        {
           return RibEdgeLogin.Visible;
        }
        public ADXExcelTaskPane1 GetPaneInstance()
        {

            try
            {
                return (ADXExcelTaskPane1)adxExcelTaskPanesCollectionItem1.TaskPaneInstance;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error In getting pane instance.");
                return null;
            }
        }
        private void LaunchExcelPane()
        {
            const string MethodName = "LaunchExcelPane";
            ADXExcelTaskPane1 EdgeExcelPane;

            EdgeExcelPane = GetPaneInstance();
            try
            {
                if (EdgeExcelPane != null)
                {
                    EdgeExcelPane.Visible = !EdgeExcelPane.Visible;
                }
                else
                {
                    adxExcelTaskPanesCollectionItem1.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                    EdgeExcelPane = (ADXExcelTaskPane1)adxExcelTaskPanesCollectionItem1.CreateTaskPaneInstance();
                    if (EdgeExcelPane != null)
                    {
                        EdgeExcelPane.Width = 500;
                        EdgeExcelPane.Show();
                        EdgeExcelPane.Visible = true;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Exception encountered while initializing excel task pane.|{ex.Message}");
            }
        }
        private void XLEdgeLogin(bool togglePane)
        {
            const string MethodName = "LaunchExcelPane";
            ADXExcelTaskPane1 EdgeExcelPane;

            XLEdgeAppState.Instance.XLEdgePane = GetPaneInstance();
            EdgeExcelPane = XLEdgeAppState.Instance.XLEdgePane;

            if (togglePane)
            {
                if (EdgeExcelPane != null)
                {
                    // Set EdgePaneShown before flipping Visible to true so the
                    // ADXBeforeTaskPaneShow handler (which hides the pane again while
                    // EdgePaneShown is false, to keep it hidden until WebView2 finishes
                    // initializing) doesn't re-hide it here.
                    bool willShow = !EdgeExcelPane.Visible;

                    try
                    {
                        if (willShow)
                        {
                            XLEdgeAppState.Instance.EdgePaneShown = true;
                        }

                        EdgeExcelPane.Visible = willShow;
                    }
                    finally
                    {
                        if (willShow)
                        {
                            XLEdgeAppState.Instance.EdgePaneShown = false;
                        }
                    }
                }

                return;
            }

            try
            {
                XLEdgeAppState.Instance.EdgePaneShown = true;

                if (EdgeExcelPane != null)
                {
                    if (!EdgeExcelPane.Visible)
                    {
                        EdgeExcelPane.Visible = true;
                    }
                    EdgeExcelPane.Activate();
                    _ = EdgeExcelPane.RefreshLoginNavigationAsync();
                    return;
                }

                adxExcelTaskPanesCollectionItem1.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                EdgeExcelPane = (ADXExcelTaskPane1)adxExcelTaskPanesCollectionItem1.CreateTaskPaneInstance();
                if (EdgeExcelPane != null)
                {
                    EdgeExcelPane.Width = 600;
                    EdgeExcelPane.Show();
                    EdgeExcelPane.Visible = true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Exception encountered while initializing excel task pane.|{ex.Message}");
            }
            finally
            {
                XLEdgeAppState.Instance.EdgePaneShown = false;
            }
        }
        /// <summary>
        /// Entry point invoked by the sibling GLSense add-in (via COM reflection: InvokeMember("InvokedFromGLSense", ...))
        /// after a GLSense login completes, to hand the resulting session over to XLEdge.
        /// Ported from VB AddinModule.vb's InvokedFromGLSense (lines 192-281).
        /// </summary>
        public void InvokedFromGLSense(string eeName, string eeUrl, string eeToken, string eeUser, bool hasXLEdgePermission = true)
        {
            const string MethodName = "InvokedFromGLSense";
            LogUtility.LogDebug($"{MethodName}|Login invoked from GLSense. URL: {eeUrl}. User: {eeUser}.");

            SafeInvokeWpf(() =>
            {
                ADXExcelTaskPane1 edgeExcelPane = GetPaneInstance();

                if (!hasXLEdgePermission)
                {
                    LogUtility.LogWarn($"{MethodName}|Login from GLSense! User {eeUser} does not have access to XLEdge.");
                    RibEdgeLogin.Enabled = false;
                    RibControlSheet.Enabled = false;
                    RibEdgeDebug.Enabled = false;
                    RibEdgeIncludeOutputData.Enabled = false;

                    if (edgeExcelPane != null && edgeExcelPane.Visible)
                    {
                        edgeExcelPane.Activate();
                        _ = edgeExcelPane.NavigateBlankAsync();
                        edgeExcelPane.Visible = false;
                    }
                    return;
                }

                try
                {
                    if (_glSenseAddinInstance == null)
                    {
                        try
                        {
                            _glSenseAddinInstance = GetGLSenseAddinObject();
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, $"{MethodName}|Failed to resolve GLSense addin object.");
                            _glSenseAddinInstance = null;
                        }
                    }

                    XLEdgeAppState.Instance.LoginUrlName = eeName;
                    XLEdgeAppState.Instance.LoginUrl = eeUrl;
                    XLEdgeAppState.Instance.LoginToken = eeToken;
                    XLEdgeAppState.Instance.LoginUserName = eeUser;
                    XLEdgeAppState.Instance.LoginFromGLSense = true;
                    XLEdgeAppState.Instance.IsLoginCompleted = true;

                    RibEdgeLogin.Visible = false;
                    RibEdgeLogout.Visible = true;
                    RibEdgeLogout.Caption = eeName;

                    RibEdgeDialogBoxLauncher.Enabled = !string.IsNullOrWhiteSpace(eeUser);

                    RibEdgeDebug.Enabled = true;
                    RibEdgeDebug.Pressed = false;
                    RibEdgeIncludeOutputData.Enabled = true;
                    RibEdgeIncludeOutputData.Pressed = false;
                    RibEdgeHelp.Enabled = true;
                    RibEdgeShowHide.Enabled = true;
                    RibEdgeOptions.Enabled = true;
                    RibEdgeRefresh.Enabled = true;
                    RibEdgeRefreshAll.Enabled = true;
                    RibEdgeParamRefresh.Enabled = true;
                    RibEdgeParamRefreshBook.Enabled = true;

                    if (edgeExcelPane != null && edgeExcelPane.Visible && !string.IsNullOrWhiteSpace(eeToken))
                    {
                        edgeExcelPane.Activate();
                        edgeExcelPane.Text = eeUrl;
                        _ = edgeExcelPane.RefreshLoginNavigationAsync();
                    }

                    RibLoginURL.Caption = eeUrl;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{MethodName}|{ex.Message}");
                }
                finally
                {
                    XLEdgeAppState.Instance.EdgePaneShown = false;
                }
            });
        }

        /// <summary>
        /// Ported from ADXExcelTaskPane1.vb's WebCtrl_SourceChanged - the reverse direction of
        /// InvokedFromGLSense: when a login completes directly through XLEdge's own WebView2 (NOT one
        /// that originated from GLSense calling InvokedFromGLSense), notify the sibling GLSense add-in
        /// via reflection so its own session/ribbon can stay in sync with this one. VB calls this
        /// "GetGLCubeInformation" - the method name is misleading (it doesn't return cube info to
        /// XLEdge; it hands GLSense the credentials it needs to load its own cube list), kept as-is to
        /// match the sibling add-in's actual method name.
        ///
        /// Guarded, matching VB's own combined condition exactly: only fires once per login (via
        /// XLEdgeAppState.LoginSentToGLSense, ported from VB's module-level "LoginSentToGLSense" flag)
        /// and never for a login that originated FROM GLSense in the first place (LoginFromGLSense) -
        /// GLSense already knows about its own-initiated logins.
        /// </summary>
        public void NotifyGLSenseOfLogin(string authToken, string loginUrl, string userName)
        {
            const string MethodName = nameof(NotifyGLSenseOfLogin);

            if (XLEdgeAppState.Instance.LoginFromGLSense ||
                string.IsNullOrWhiteSpace(authToken) ||
                XLEdgeAppState.Instance.LoginSentToGLSense)
            {
                return;
            }

            try
            {
                if (_glSenseAddinInstance == null)
                {
                    try
                    {
                        _glSenseAddinInstance = GetGLSenseAddinObject();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"{MethodName}|Failed to resolve GLSense addin object.");
                        _glSenseAddinInstance = null;
                    }
                }

                if (_glSenseAddinInstance == null)
                {
                    return;
                }

                // Matches VB's ordering exactly: the flag is set BEFORE attempting the reflection call,
                // so a failed/unavailable GLSense instance doesn't cause this to retry on every
                // subsequent navigation within the same session.
                XLEdgeAppState.Instance.LoginSentToGLSense = true;

                try
                {
                    _glSenseAddinInstance.GetType().InvokeMember(
                        "GetGLCubeInformation",
                        System.Reflection.BindingFlags.InvokeMethod,
                        null,
                        _glSenseAddinInstance,
                        new object[] { authToken, loginUrl, userName });
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{MethodName}|Failed to invoke GetGLCubeInformation on GLSense addin.");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Unexpected error notifying GLSense of login.");
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's LogOffAllTaskPanesAsync - the symmetric logout counterpart of
        /// NotifyGLSenseOfLogin: tells the sibling GLSense add-in (via the same cached COM object /
        /// reflection call) that this XLEdge session has logged out, by invoking its "LogoutSession"
        /// method. Called from LogoffFromXLEdgeAddin's state-reset block.
        /// </summary>
        private void NotifyGLSenseOfLogout()
        {
            const string MethodName = nameof(NotifyGLSenseOfLogout);

            if (_glSenseAddinInstance == null)
            {
                return;
            }

            try
            {
                _glSenseAddinInstance.GetType().InvokeMember(
                    "LogoutSession",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    _glSenseAddinInstance,
                    Array.Empty<object>());
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Failed to invoke LogoutSession on GLSense addin.");
            }
        }

        public bool NavigateReportsToAddress(string name, string address)
        {
            try
            {
                return WpfAppManager.InvokeOnWpfThread(() =>
                {
                    string normalizedAddress = NormalizeAddress(address);
                    if (string.IsNullOrWhiteSpace(normalizedAddress))
                    {
                        return false;
                    }
                    XLEdgeAppState.Instance.LoginUrl = normalizedAddress;
                    XLEdgeAppState.Instance.LoginUrlName = name;
                    XLEdgeLogin(false);
                    return true;
                });
            }
            catch (Exception ex)
            {
                SafeLogException(ex, "Exception occured while navigating reports task pane");
                return false;
            }
        }

        private static string NormalizeAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
            {
                return string.Empty;
            }

            string normalized = address.Trim().TrimEnd('\r', '\n', '/', '\\', ' ');
            string[] patterns = { "/bypass-saml-login-flow", "/bypass-sso-login-flow" };

            foreach (string pattern in patterns)
            {
                normalized = Regex.Replace(normalized, Regex.Escape(pattern), string.Empty, RegexOptions.IgnoreCase);
            }

            return normalized;
        }

        // Centralized helpers to reduce duplication and improve error handling
        private static void SafeInvokeWpf(System.Action action)
        {
            if (action == null) return;
            try
            {
                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        action();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex);
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private static void SafeLogException(Exception ex, string context = null)
        {
            if (ex == null) return;
            if (!string.IsNullOrEmpty(context))
                LogUtility.LogError($"{context}: {ex.Message}");
            LogUtility.LogException(ex);
        }

        private void AddinModule_OnRibbonLoaded(object sender, IRibbonUI ribbon)
        {
            try
            {
                LogHelper.InitializeLogger();

                XLApp.Initialize(this.HostApplication as Excel.Application);

                XLEdgePreferencesManager.Instance.Initialize();

                // Ported from AddinModule.vb's AddinModule_OnRibbonLoaded -> EEDeleteAllFiles() call -
                // clears out any temp report CSVs left over from a previous Excel session/crash.
                try
                {
                    Helpers.XLEdgeTempFileCleaner.DeleteAllTempFiles();
                }
                catch (Exception cleanupEx)
                {
                    SafeLogException(cleanupEx, "Exception occured cleaning up temp files on ribbon load");
                }

                _ribbonHelper = new XLEdgeRibbonHelper(AddinModule.CurrentInstance, ribbon);
                RibbonHelper = _ribbonHelper; // Expose it globally

                if (!string.IsNullOrWhiteSpace(_pendingRibbonState))
                {
                    _ribbonHelper.ApplyState(_pendingRibbonState);
                    _pendingRibbonState = null;
                }
                else if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    _ribbonHelper.ApplyState("LoggedOut");
                }
                else
                {
                    _ribbonHelper.ApplyState("LoggedIn");
                }

                WpfUiBootstrapper.Init(XLEdgeAppConstants.GLAccentHex, XLEdgeAppConstants.GLTheme);
                WpfUiBootstrapper.PreloadResources();
            }
            catch (Exception ex)
            {
                SafeLogException(ex, "Exception occured in AddinModule_OnRibbonLoaded");
            }
            
        }
        private void AddinModule_OnError(AddinExpress.MSO.ADXErrorEventArgs e)
        {
            e.Handled = true;
            MessageFunctions.XLEdgeMessage("Error: " + e.ADXError.ToString(), MessageBoxIcon.Error, MessageBoxButtons.OK);
        }
        private void RibLogin_OnClick(object sender, IRibbonControl control, bool pressed) //RibLogin Code
        {
            try
            {
                IntPtr excelHandle = ExcelApplicationHelper.GetExcelWindowHandle();

                SafeInvokeWpf(() =>
                {
                    var win = new XLEdgeServerConfiguration();
                    if (excelHandle != IntPtr.Zero)
                    {
                        win.ShowDialogWithOwner(excelHandle);
                    }
                    else
                    {
                        win.ShowDialog();
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
        {
            if (!XLEdgeAppState.Instance.IsLoginCompleted)
                return;

            if (XLApp.App == null)
            {
                LogUtility.LogError("Excel application instance is not available in SheetSelectionChange event.");
                return;
            }

            if (sheet is not Excel.Worksheet selectedSheet)
                return;

            if (range is not Excel.Range selectedRange)
                return;

            if (selectedRange.Cells.Count != 1)
                return;

            // Each of these is independently gated by its own Options checkbox, so both can run off
            // the same selection-changed event without interfering with one another.
            TryShowCalendarControl(selectedSheet, selectedRange);
            TryShowSegmentSelectionWindow(selectedSheet, selectedRange);
        }

        private void TryShowCalendarControl(Excel.Worksheet selectedSheet, Excel.Range selectedRange)
        {
            if (_isCalendarOpen || !XLEdgeAppState.Instance.ShowCalendarControl)
                return;

            IntPtr excelHandle = XLApp.Handle;
            _isCalendarOpen = true;
            try
            {
                // The calendar must only ever appear for a genuine Date parameter cell inside the
                // orb_params_control table. Previously, when the active sheet had no ListObjects at
                // all (e.g. any blank/empty worksheet, or any sheet before the Parameters Control
                // Sheet has even been created), the "if (selectedSheet.ListObjects.Count > 0)" guard
                // below was simply false, so every one of its internal validation checks was skipped
                // entirely and execution fell straight through to showing the calendar for ANY
                // single-cell selection. That's also why it was popping up immediately after login:
                // the post-login focus-release fix (ReportGenerator.
                // ReleaseKeyboardFocusFromTaskPaneAsync) deliberately reselects a cell to nudge
                // keyboard focus back to Excel, which fires this same SheetSelectionChange handler on
                // whatever sheet/cell happened to be active at login time - almost never the
                // Parameters Control Sheet. Rewritten to look up the orb_params_control table by name
                // (mirroring TryShowSegmentSelectionWindow below) and return immediately if it isn't
                // present on the active sheet at all, instead of only checking it once one exists.
                Excel.ListObject tableObj = null;
                foreach (Excel.ListObject lo in selectedSheet.ListObjects)
                {
                    if (lo.Name.Equals("orb_params_control", StringComparison.OrdinalIgnoreCase))
                    {
                        tableObj = lo;
                        break;
                    }
                }

                if (tableObj == null ||
                    XLApp.App.Application.Intersect(selectedRange, tableObj.DataBodyRange) == null)
                {
                    return;
                }

                int row = selectedRange.Row;
                int col = selectedRange.Column;
                if (row < 4 || (col != 10 && col != 11))
                    return;

                Excel.Range dataTypeCell = (Excel.Range)selectedSheet.Cells[row, 8];
                Excel.Range operatorCell = (Excel.Range)selectedSheet.Cells[row, 9];

                string dataType = Convert.ToString(dataTypeCell.Value).ToUpper();
                string operatorValue = Convert.ToString(operatorCell.Value).ToUpper();

                if (!dataType.Contains("DATE"))
                    return;

                // Value1 (column J, col 10) always gets the calendar for a Date parameter,
                // regardless of the operator. Value2 (column K, col 11) only participates in a
                // range, so it should only get the calendar when the operator is actually
                // "between". This was previously backwards: "!hVal.Contains("BETWEEN") && col !=
                // 11" blocked J whenever the operator wasn't "between" (it should never be gated
                // on the operator at all) and let K through regardless of the operator (it should
                // require "between").
                if (col == 11 && !operatorValue.Contains("BETWEEN"))
                    return;

                DateTime initialDate = XLApp.GetDateFromCell(selectedRange) ?? DateTime.Today;
                DateTime? selectedDate = null;

                double explicitLeft = XLApp.App.ActiveWindow.PointsToScreenPixelsX((int)Math.Round(Convert.ToDouble(selectedRange.Left)));
                double explicitTop = XLApp.App.ActiveWindow.PointsToScreenPixelsY((int)Math.Round(Convert.ToDouble(selectedRange.Top) + Convert.ToDouble(selectedRange.Height)));

                SafeInvokeWpf(() =>
                {
                    var calendarForm = new XLEdgeCalendar(initialDate);
                    if (calendarForm.ShowDialogWithOwner(excelHandle, explicitLeft, explicitTop) == true)
                    {
                        selectedDate = calendarForm.SelectedDate;
                    }
                });

                if (selectedDate.HasValue)
                {
                    XLApp.WriteDateToCell(selectedRange, selectedDate.Value);
                }
            }
            finally
            {
                _isCalendarOpen = false;
            }
        }

        // Ported from the former adxExcelAppEvents1_SheetBeforeDoubleClick handler: previously this
        // GL segment picker only opened on a double-click of column J in the control table. Per
        // request, it's now gated by its own Options checkbox (ShowSegmentSelectionWindow) and, when
        // enabled, opens on simple selection - consistent with how the calendar control works, and
        // extensible the same way if more "show a picker window" options are added later.
        private void TryShowSegmentSelectionWindow(Excel.Worksheet selectedSheet, Excel.Range selectedRange)
        {
            if (_isSegmentWindowOpen || !XLEdgeAppState.Instance.ShowSegmentSelectionWindow)
                return;

            if (!selectedSheet.Name.Equals("Parameters Control Sheet", StringComparison.OrdinalIgnoreCase))
                return;

            // Only column J (10) - the Value1 column - triggers the segment picker.
            if (selectedRange.Column != 10)
                return;

            _isSegmentWindowOpen = true;
            try
            {
                Excel.ListObject controlTable = null;
                foreach (Excel.ListObject lo in selectedSheet.ListObjects)
                {
                    if (lo.Name.Equals("orb_params_control", StringComparison.OrdinalIgnoreCase))
                    {
                        controlTable = lo;
                        break;
                    }
                }

                if (controlTable == null)
                    return;

                if (selectedRange.Row < controlTable.HeaderRowRange.Row ||
                    selectedRange.Row > controlTable.DataBodyRange.Rows.Count + controlTable.HeaderRowRange.Row)
                    return;

                Excel.Range paramTypeCell = selectedSheet.Cells[selectedRange.Row, 4] as Excel.Range;
                string paramType = paramTypeCell?.Value2 as string ?? string.Empty;
                if (!paramType.Equals("extraParameters", StringComparison.OrdinalIgnoreCase))
                    return;

                Excel.Range paramNameCell = selectedSheet.Cells[selectedRange.Row, 5] as Excel.Range;
                string paramName = paramNameCell?.Value2 as string ?? string.Empty;
                if (!paramName.Equals("ORACLE_GL_SEGMENT_VALUES", StringComparison.OrdinalIgnoreCase))
                    return;

                Excel.Range displayValueCell = selectedSheet.Cells[selectedRange.Row, 235] as Excel.Range;
                string displayValues = displayValueCell?.Value2 as string ?? string.Empty;

                if (!string.IsNullOrEmpty(displayValues))
                {
                    int rowNumber = selectedRange.Row;
                    SafeInvokeWpf(() =>
                    {
                        var window = new XLEdgeGLAccountsWindow(selectedSheet, rowNumber, displayValues);
                        window.ShowDialog();
                    });
                }
                else
                {
                    MessageFunctions.XLEdgeMessage(
                        "No segment display values found for this row.\n" +
                        "Please ensure the GL Accounts data is properly loaded.",
                        System.Windows.Forms.MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SheetSelectionChange - GL Accounts segment window");
            }
            finally
            {
                _isSegmentWindowOpen = false;
            }
        }

        private void RibEdgeOptions_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    MessageFunctions.XLEdgeMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                IntPtr excelHandle = IntPtr.Zero;
                if (ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
                {
                    excelHandle = new IntPtr(excelApp.Hwnd);
                }

                SafeInvokeWpf(() =>
                {
                    var win = new XLEdgeOptions();
                    if (excelHandle != IntPtr.Zero)
                    {
                        win.ShowDialogWithOwner(excelHandle);
                    }
                    else
                    {
                        win.ShowDialog();
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void RibEdgeDialogBoxLauncher_OnAction(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    MessageFunctions.XLEdgeMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                IntPtr excelHandle = IntPtr.Zero;
                if (ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
                {
                    excelHandle = new IntPtr(excelApp.Hwnd);
                }

                SafeInvokeWpf(() =>
                {
                    var win = new XLEdgeLoginDetails();
                    if (excelHandle != IntPtr.Zero)
                    {
                        win.ShowDialogWithOwner(excelHandle);
                    }
                    else
                    {
                        win.ShowDialog();
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void RibEdgeAbout_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            IntPtr excelHandle = IntPtr.Zero;
            if (ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
            {
                excelHandle = new IntPtr(excelApp.Hwnd);
            }

            SafeInvokeWpf(() =>
            {
                var win = new XLEdgeAbout();
                if (excelHandle != IntPtr.Zero)
                {
                    win.ShowDialogWithOwner(excelHandle);
                }
                else
                {
                    win.ShowDialog();
                }
            });
        }

        private void RibEdgeHelp_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    MessageFunctions.XLEdgeMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                if (!string.IsNullOrEmpty(XLEdgeAppState.Instance.LoginToken))
                {
                    string helpUrl = XLEdgeAppState.Instance.LoginUrl + "/web/public/redirect-help/Excel_XLEdge.htm?jwtParam=" + XLEdgeAppState.Instance.LoginToken;
                    Process.Start(helpUrl);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }
        private async Task LogOffAllTaskPanesAsync(CancellationToken token)
        {
            if (adxExcelTaskPanesCollectionItem1.TaskPaneInstances.Count <= 0)
            {
                return;
            }

            foreach (ADXExcelTaskPane1 xlTaskpane in adxExcelTaskPanesCollectionItem1.TaskPaneInstances)
            {
                if (token.IsCancellationRequested)
                {
                    break;
                }

                try
                {
                    await xlTaskpane.LogoutAsync(XLEdgeAppState.Instance.LoginUrl, token);
                    await Task.Delay(200, token);
                    xlTaskpane.HidePaneSafe();
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Error while logging off task pane.");
                }
            }
        }
        private async Task LogoffFromXLEdgeAddin()
        {
            XLEdgeWaitWindow waitWindow = null;
            CancellationTokenSource linkedCts = null;

            try
            {
                // Uses the single shared app-wide WPF dispatcher (UiDispatcher.Current, same as every
                // other wait/busy window in this add-in - see ReportGenerator.CreateAndShowWaitWindow)
                // rather than an arbitrary open workbook's task pane dispatcher. Picking "the first
                // available" task pane here was fragile with multiple workbooks open - if that
                // particular pane's dispatcher wasn't ready/valid for any reason, the wait window
                // silently never showed. Showing the wait window is a UX nicety only - any failure to
                // show it is caught and logged on its own here so it can never block the actual logoff
                // work below.
                try
                {
                    await UiDispatcher.RunAsync(() =>
                    {
                        waitWindow = new XLEdgeWaitWindow();

                        PackIconFontAwesomeKind icon = PackIconFontAwesomeKind.DoorOpenSolid;
                        waitWindow.SetProcessTitle("Logging Off", icon);
                        waitWindow.SetProcessMessage("Please wait...");
                        waitWindow.Show();
                    });
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "LogoffFromAddin|Failed to show wait window - continuing logoff without it.");
                    waitWindow = null;
                }

                linkedCts = waitWindow != null
                    ? CancellationTokenSource.CreateLinkedTokenSource(waitWindow.Token)
                    : new CancellationTokenSource();

                using (linkedCts)
                {
                    try
                    {
                        await LogOffAllTaskPanesAsync(linkedCts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        LogUtility.LogWarn("LogoffFromAddin|Cancelled by user.");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "LogoffFromAddin|Error during logoff.");
            }
            finally
            {
                if (waitWindow != null)
                {
                    try
                    {
                        await waitWindow.Dispatcher.InvokeAsync(() =>
                        {
                            if (waitWindow.IsVisible)
                            {
                                waitWindow.RequestClose();
                            }
                        });

                        await Task.Delay(150);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "Error closing XLEdgeWaitWindow.");
                    }
                }
            }

            // Ported from AddinModule.vb's LogOffAllTaskPanesAsync - tell the sibling GLSense add-in
            // this session logged out before clearing our own login state below.
            try
            {
                NotifyGLSenseOfLogout();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "LogoffFromXLEdgeAddin|Failed to notify GLSense of logout.");
            }

            XLEdgeAppState.Instance.IsLoginCompleted = false;
            XLEdgeAppState.Instance.LoginToken = string.Empty;
            XLEdgeAppState.Instance.LoginUserName = string.Empty;
            XLEdgeAppState.Instance.LoginUrl = string.Empty;
            XLEdgeAppState.Instance.LoginUrlName = string.Empty;
            XLEdgeAppState.Instance.LoginToken = string.Empty;
            XLEdgeAppState.Instance.LoginFromGLSense = false;
            XLEdgeAppState.Instance.LoginSentToGLSense = false;
            XLEdgeAppState.Instance.DebugLogs = false;
            XLEdgeAppState.Instance.EdgePaneShown = false;

            RibSheetLabel.Caption = " ";

            XLEdgePreferencesManager.Instance.ResetRuntimeFromSaved();
        }

        /// <summary>
        /// Entry point invoked by the sibling GLSense add-in (via COM reflection:
        /// InvokeMember("LogoffFromAddin", ...)) when the user logs out of GLSense, so this
        /// XLEdge session gets logged out in lockstep - the reverse direction of
        /// InvokedFromGLSense above. GLSense's caller (AddinModule.RibLogout_OnClick) invokes
        /// this synchronously via late-bound COM reflection and discards the result (same
        /// fire-and-forget pattern it already uses), so this is a thin public wrapper around
        /// the existing LogoffFromXLEdgeAddin/ApplyRibbonState sequence already used by this
        /// add-in's own Logout ribbon button (RibEdgeLogout_OnClick) - no change to that
        /// existing logic, just exposing it under the name GLSense already expects to call.
        /// Previously GLSense called InvokeMember("LogoffFromAddin", ...), but no method with
        /// that exact name existed on this class (the real method is the private
        /// LogoffFromXLEdgeAddin) - IDispatch could only report DISP_E_UNKNOWNNAME ("COM
        /// object that has been separated from its underlying RCW cannot be used" as
        /// GLSense's wrapper reported it), silently skipping the entire logoff so GLSense
        /// logged itself out while this XLEdge session stayed logged in.
        /// </summary>
        public void LogoffFromAddin()
        {
            const string MethodName = "LogoffFromAddin";
            LogUtility.LogDebug($"{MethodName}|Logout invoked from GLSense.");

            _ = LogoffFromAddinAsync();
        }

        private async Task LogoffFromAddinAsync()
        {
            try
            {
                await LogoffFromXLEdgeAddin();
                ApplyRibbonState("LoggedOut");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "LogoffFromAddin|Error during logoff invoked from GLSense.");
            }
        }

        private async void RibEdgeLogout_OnClick(object sender, IRibbonControl control, bool pressed)
        {

            await LogoffFromXLEdgeAddin();

            ApplyRibbonState("LoggedOut");
        }
        
        private void RibEdgeShowHide_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    MessageFunctions.XLEdgeMessage("Please log in to the instance.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                XLEdgeLogin(true);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        /// <summary>
        /// Refreshes the "RibSheetLabel" ribbon caption to describe the active sheet's XLEdge table
        /// (scheduled output / data report / drilldown+attachment column summary), warns if a required
        /// parameter sheet is missing, and cleans up orphaned named ranges via DeleteNamedCache.
        /// Ribbon button enable/disable state itself is delegated to ApplyRibbonState /
        /// XLEdgeRibbonHelper.ProcessActiveWorkbook rather than duplicated here.
        /// </summary>
        public void UpdateTabLabel(Excel.Worksheet workSheet)
        {
            RibSheetLabel.Caption = " ";

            if (workSheet == null)
            {
                return;
            }

            try
            {
                

                if (workSheet.ListObjects.Count == 0)
                {
                    return;
                }

                Excel.ListObject tableObj = workSheet.ListObjects[1];
                string tableName = tableObj.Name ?? string.Empty;

                if (tableName.Length <= 3 || !tableName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                if (tableName.EndsWith("_P", StringComparison.Ordinal))
                {
                    RibSheetLabel.Caption = "This sheet has a scheduled output.";
                    XLEdgeAppState.Instance.RefreshAll = false;
                    return;
                }

                if (!tableName.EndsWith("_E", StringComparison.Ordinal))
                {
                    XLEdgeAppState.Instance.RefreshAll = false;
                    return;
                }

                XLEdgeAppState.Instance.RefreshAll = true;

                // A companion "P_" parameter sheet only exists for reports generated in
                // separate-sheet mode (table starts at row 1). Same-sheet mode reports never have
                // one, so only check for it when the table actually starts at row 1.
                bool expectsCompanionParamSheet = tableObj.HeaderRowRange != null && tableObj.HeaderRowRange.Row <= 1;

                if (expectsCompanionParamSheet)
                {
                    string paramSheetName = $"P_{workSheet.Name}";
                    Excel.Worksheet paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableName);
                    if (paramSheet == null)
                    {
                        MessageFunctions.XLEdgeMessage(
                            "Reports parameters information worksheet missing. Please rerun the report to generate.",
                            MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    }
                }

                if (tableObj.DataBodyRange != null && tableObj.DataBodyRange.Rows.Count >= 1 && tableObj.HeaderRowRange != null)
                {
                    var drillCols = new List<string>();
                    var attCols = new List<string>();

                    // Resize[1, colCount] (not Rows[1]) - the Rows collection's integer indexer is
                    // declared to return object and has been observed to throw InvalidCastException at
                    // runtime, the same category of Excel Interop early-bound-cast gotcha as
                    // Worksheets/Sheets fixed earlier in this migration. Resize is a plain property
                    // returning Excel.Range directly, so it's safe.
                    Excel.Range firstDataRow = tableObj.DataBodyRange.Resize[1, tableObj.DataBodyRange.Columns.Count];
                    try
                    {
                        foreach (Excel.Range cell in firstDataRow.Columns)
                        {
                            try
                            {
                                if (cell.Hyperlinks.Count == 0)
                                {
                                    continue;
                                }

                                Excel.Hyperlink hyperlink = cell.Hyperlinks[1];
                                string screenTip = hyperlink.ScreenTip ?? string.Empty;

                                Excel.Range headerCell = (Excel.Range)tableObj.HeaderRowRange.Cells[1, cell.Column - tableObj.HeaderRowRange.Column + 1];
                                string headerText = Convert.ToString(headerCell.Value) ?? string.Empty;

                                if (screenTip.IndexOf("DRILLDOWN", StringComparison.OrdinalIgnoreCase) >= 0 && !drillCols.Contains(headerText))
                                {
                                    drillCols.Add(headerText);
                                }

                                if (screenTip.IndexOf("ATTACHMENT", StringComparison.OrdinalIgnoreCase) >= 0 && !attCols.Contains(headerText))
                                {
                                    attCols.Add(headerText);
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(cell);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(firstDataRow);
                    }

                    ApplyRibbonState(workSheet.Parent as Excel.Workbook);

                    if (drillCols.Count > 0 && attCols.Count > 0)
                    {
                        RibSheetLabel.Caption = $"This sheet has drilldown and attachment links: Drilldowns on column(s): {string.Join(", ", drillCols)}  |  Attachments on column(s): {string.Join(", ", attCols)}. By default all attachments are saved in the downloads folder.";
                    }
                    else if (drillCols.Count > 0)
                    {
                        RibSheetLabel.Caption = $"This sheet has drilldown links on column(s): {string.Join(", ", drillCols)}";
                    }
                    else if (attCols.Count > 0)
                    {
                        RibSheetLabel.Caption = $"This sheet has attachment links on column(s): {string.Join(", ", attCols)}. By default all attachments are saved in the downloads folder.";
                    }
                    else
                    {
                        RibSheetLabel.Caption = "This sheet has a data report.";
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(UpdateTabLabel));
            }
            finally
            {
                DeleteNamedCache();
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's RibEdgeLogout_PropertyChanging (Handles RibEdgeLogout.PropertyChanging).
        /// Whenever the Logout button's visibility changes (i.e. right after login/logout swaps which
        /// button is shown), refresh the active sheet's tab label to match the current sheet.
        ///
        /// NOTE: could not verify the exact Add-in Express delegate/event-args type name for ribbon
        /// control PropertyChanging from this environment - "ADXRibbonControlPropertyChanging_EventHandler"
        /// / "ADXRibbonPropertyChangingEventArgs" below are best-guess names following ADX's naming
        /// convention seen elsewhere in this file (e.g. ADXError_EventHandler/ADXErrorEventArgs). Confirm
        /// in Visual Studio and adjust the wiring in AddinModule.Designer.cs if it doesn't compile.
        /// </summary>
        private void RibEdgeLogout_PropertyChanging(object sender, AddinExpress.MSO.ADXRibbonPropertyChangingEventArgs e)
        {
            try
            {
                if (e.PropertyType == AddinExpress.MSO.ADXRibbonControlPropertyType.Visible &&
                    ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
                {
                    UpdateTabLabel(excelApp.ActiveSheet as Excel.Worksheet);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RibEdgeLogout_PropertyChanging));
            }
        }

        /// <summary>Ported from AddinModule.vb's RibEdgeIncludeOutputData_OnClick.</summary>
        private void RibEdgeIncludeOutputData_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                XLEdgeAppState.Instance.DebugOutputData = pressed;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RibEdgeIncludeOutputData_OnClick));
            }
        }

        /// <summary>Ported from AddinModule.vb's RibEdgeDebug_OnClick.</summary>
        private void RibEdgeDebug_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                XLEdgeAppState.Instance.DebugLogs = pressed;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RibEdgeDebug_OnClick));
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's AdxExcelAppEvents1_SheetBeforeDelete. Rather than deleting a
        /// table's companion parameter sheet immediately (which can crash Excel when done from directly
        /// inside a BeforeDelete handler), the companion sheet's name is queued and removed a moment
        /// later by DeleteTimer_Tick.
        /// </summary>
        private void AdxExcelAppEvents1_SheetBeforeDelete(object sender, object sheet)
        {
            const string MethodName = "AdxExcelAppEvents1_SheetBeforeDelete";

            if (!(sheet is Excel.Worksheet worksheet))
            {
                return;
            }

            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
            {
                return;
            }

            try
            {
                if (worksheet.ListObjects.Count > 0)
                {
                    Excel.ListObject tableObj = worksheet.ListObjects[1];
                    string tableName = tableObj.Name ?? string.Empty;

                    if ((tableName.EndsWith("_P", StringComparison.Ordinal) || tableName.EndsWith("_E", StringComparison.Ordinal)))
                    {
                        string paramSheetName = $"P_{worksheet.Name}";
                        Excel.Worksheet paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableName);
                        if (paramSheet != null && !_sheetsToDelete.Contains(paramSheet.Name))
                        {
                            _sheetsToDelete.Add(paramSheet.Name);
                        }
                    }
                }
                else
                {
                    string tableName = Convert.ToString(worksheet.Range["IT2"]?.Value) ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(tableName) &&
                        tableName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) &&
                        (tableName.EndsWith("_P", StringComparison.Ordinal) || tableName.EndsWith("_E", StringComparison.Ordinal)))
                    {
                        Excel.Worksheet sheetToDelete = FindSheetWithTableName(excelApp.ActiveWorkbook, tableName);
                        if (sheetToDelete != null && !_sheetsToDelete.Contains(sheetToDelete.Name))
                        {
                            _sheetsToDelete.Add(sheetToDelete.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Exception queueing companion worksheet for deferred delete.");
            }

            if (_sheetsToDelete.Count > 0)
            {
                _deleteTimer.Stop();
                _deleteTimer.Start();
            }
        }

        /// <summary>Ported from AddinModule.vb's DeleteTimer_Tick - performs the deferred deletes queued by
        /// AdxExcelAppEvents1_SheetBeforeDelete.</summary>
        private void DeleteTimer_Tick(object sender, EventArgs e)
        {
            _deleteTimer.Stop();

            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp) || excelApp.ActiveWorkbook == null)
            {
                _sheetsToDelete.Clear();
                return;
            }

            bool restoreAlerts = excelApp.DisplayAlerts;
            bool restoreEvents = excelApp.EnableEvents;

            try
            {
                excelApp.DisplayAlerts = false;
                excelApp.EnableEvents = false;

                foreach (string sheetName in _sheetsToDelete.Distinct(StringComparer.OrdinalIgnoreCase).ToList())
                {
                    Excel.Worksheet sht = null;
                    try
                    {
                        foreach (Excel.Worksheet ws in excelApp.ActiveWorkbook.Worksheets)
                        {
                            if (sht == null && string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            {
                                sht = ws;
                            }
                            else
                            {
                                Marshal.ReleaseComObject(ws);
                            }
                        }

                        sht?.Delete();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"DeleteTimer_Tick|Deferred delete failed for sheet '{sheetName}'.");
                    }
                    finally
                    {
                        if (sht != null)
                        {
                            Marshal.ReleaseComObject(sht);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DeleteTimer_Tick|Exception during deferred delete.");
            }
            finally
            {
                _sheetsToDelete.Clear();
                try
                {
                    excelApp.DisplayAlerts = restoreAlerts;
                    excelApp.EnableEvents = restoreEvents;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "DeleteTimer_Tick|Exception restoring Excel alert/event state.");
                }
            }
        }

        private static Excel.Worksheet FindSheetWithTableName(Excel.Workbook workbook, string tableName)
        {
            if (workbook == null)
            {
                return null;
            }

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                bool release = true;
                try
                {
                    if (ws.ListObjects.Count > 0 && string.Equals(ws.ListObjects[1].Name, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        release = false;
                        return ws;
                    }
                }
                finally
                {
                    if (release)
                    {
                        Marshal.ReleaseComObject(ws);
                    }
                }
            }

            return null;
        }

        /// <summary>Ported from AddinModule.vb's DeleteHashNamedRanges - removes "ORB_"-prefixed named
        /// ranges whose RefersTo has become a broken #REF! reference.</summary>
        private void DeleteHashNamedRanges()
        {
            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp) || excelApp.ActiveWorkbook == null)
            {
                return;
            }

            try
            {
                if (excelApp.ActiveWorkbook.Names.Count == 0)
                {
                    return;
                }

                foreach (Excel.Name nm in excelApp.ActiveWorkbook.Names)
                {
                    try
                    {
                        string refersTo = SafeRefersTo(nm);
                        if (nm.Name.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) &&
                            refersTo.IndexOf("#REF", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            nm.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "DeleteHashNamedRanges|Exception inspecting named range.");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(nm);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(DeleteHashNamedRanges));
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's DeleteNamedCache. Cleans up two categories of orphaned "ORB_"
        /// named ranges left behind once their owning sheet is gone: "_ChildReport"/"_Instance" caches
        /// (checked via NmSheetExists) and generic broken #REF! ranges.
        /// </summary>
        private void DeleteNamedCache()
        {
            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp) || excelApp.ActiveWorkbook == null)
            {
                return;
            }

            try
            {
                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook.Names.Count == 0)
                {
                    return;
                }

                foreach (Excel.Name nm in workbook.Names)
                {
                    try
                    {
                        string nmName = nm.Name ?? string.Empty;

                        if (nmName.Length > 5 && nmName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) &&
                            (nmName.IndexOf("_ChildReport", StringComparison.OrdinalIgnoreCase) >= 0 ||
                             nmName.IndexOf("_Instance", StringComparison.OrdinalIgnoreCase) >= 0))
                        {
                            string shtName = nmName.Substring(4);
                            int idx = shtName.LastIndexOf('_');
                            if (idx > 0)
                            {
                                shtName = shtName.Substring(0, idx);
                                if (!NmSheetExists(shtName, workbook))
                                {
                                    nm.Delete();
                                }
                            }
                        }
                        else if (nmName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase))
                        {
                            string refersTo = SafeRefersTo(nm);
                            if (refersTo.IndexOf("#REF", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                nm.Delete();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "DeleteNamedCache|Exception inspecting named range.");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(nm);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(DeleteNamedCache));
            }
        }

        private static string SafeRefersTo(Excel.Name nm)
        {
            try
            {
                return nm.RefersTo?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort read of a named range's RefersTo formula; caller treats
                // an empty string as "unresolvable"/no match.
                LogUtility.LogDebug($"{nameof(SafeRefersTo)}: failed to read RefersTo - {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>Ported from AddinModule.vb's NmSheetExists - checks whether a worksheet whose
        /// "cleaned" name (alphanumeric/underscore only, see ExcelSheetHelper.CleanUpName) matches
        /// the given name still exists in the workbook.</summary>
        private static bool NmSheetExists(string sheetName, Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return false;
            }

            try
            {
                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    try
                    {
                        if (string.Equals(ExcelSheetHelper.CleanUpName(ws.Name), sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(ws);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(NmSheetExists));
                return false;
            }
        }

        private void adxExcelAppEvents1_SheetActivate(object sender, object hostObj)
        {
            // Wrapped in try/catch since this is a raw Excel COM event sink - any exception here
            // would otherwise surface as an unhandled error straight out of the COM callback.
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    return;
                }
                ApplyRibbonState("ApplySheetActiveState");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(adxExcelAppEvents1_SheetActivate));
            }
        }

        private void adxExcelAppEvents1_WorkbookActivate(object sender, object hostObj)
        {
            // Wrapped in try/catch for the same reason as adxExcelAppEvents1_SheetActivate above.
            try
            {
                if (!XLEdgeAppState.Instance.IsLoginCompleted)
                {
                    return;
                }

                if (hostObj is Excel.Workbook workbook)
                {
                    ApplyRibbonState(workbook);
                    return;
                }

                ApplyRibbonState("ApplySheetActiveState");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(adxExcelAppEvents1_WorkbookActivate));
            }
        }

        public static void ApplyRibbonState(string stateName)
        {
            if (RibbonHelper != null)
            {
                RibbonHelper.ApplyState(stateName);
                return;
            }

            _pendingRibbonState = stateName;
            LogUtility.LogDebug($"[AddinModule] Ribbon helper is not ready. Queued ribbon state '{stateName}'.");
        }

        public static void ApplyRibbonState(Excel.Workbook workbook)
        {
            if (RibbonHelper != null)
            {
                RibbonHelper.ApplyWorkbookActiveState(workbook);
                return;
            }

            _pendingRibbonState = "ApplySheetActiveState";
            LogUtility.LogDebug("[AddinModule] Ribbon helper is not ready. Queued workbook ribbon state refresh.");
        }

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            const string MethodName = "AddinModule_AddinStartupComplete";
            try
            {
                XLApp.Initialize(this.HostApplication as Excel.Application);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
            }
        }

        private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
        {
            // Releases cached active-sheet/workbook COM overrides on add-in shutdown. Does not
            // call Excel.Quit() - disabling/unloading the add-in should not close Excel itself.
            try
            {
                Excel.Worksheet worksheetOverride = XLEdgeAppState.Instance.ActiveWorksheetOverride;
                if (worksheetOverride != null)
                {
                    Marshal.FinalReleaseComObject(worksheetOverride);
                    XLEdgeAppState.Instance.ActiveWorksheetOverride = null;
                }

                Excel.Workbook workbookOverride = XLEdgeAppState.Instance.ActiveWorkbookOverride;
                if (workbookOverride != null)
                {
                    Marshal.FinalReleaseComObject(workbookOverride);
                    XLEdgeAppState.Instance.ActiveWorkbookOverride = null;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(AddinModule_AddinBeginShutdown));
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private async void RibEdgeRefresh_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            const string MethodName = "RibEdgeRefresh_OnClick";

            if (string.IsNullOrEmpty(XLEdgeAppState.Instance.LoginUrl))
            {
                return;
            }

            if (ExcelApplicationHelper.IsCellInEditMode())
            {
                LogUtility.LogWarn($"{MethodName}|Active cell is in edit mode. Exit edit mode and try again.");
                return;
            }

            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                Excel.Worksheet activeSheet = excelApp.ActiveSheet as Excel.Worksheet;

                if (activeSheet == null || activeSheet.ListObjects.Count == 0)
                {
                    MessageFunctions.XLEdgeMessage("No reports in the sheet to refresh", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                Excel.ListObject tableObj = activeSheet.ListObjects[1];
                if (tableObj == null || !tableObj.Name.EndsWith("_E", StringComparison.Ordinal))
                {
                    MessageFunctions.XLEdgeMessage("Active sheet does not contain a data report.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                // The VB original additionally resolved a separate parameter sheet here, validated
                // its IT5 (logged-in instance) / IT1 ("Child Report") markers, and built request
                // post-data via XLEdgeParamsData.BuildParamData so Refresh would pick up any
                // parameter edits made in the "orb_params_control" sheet - not just blindly re-fetch
                // with the report's original parameters. RefreshListObjectAsync re-derives the
                // runId/columns it needs on its own, so the instance/child-report validation isn't
                // needed anymore, but building and passing the edited-parameters payload is real,
                // reachable functionality - restored below via BuildRefreshParamsPayload.
                string paramsPayload = BuildRefreshParamsPayload(excelApp.ActiveWorkbook, activeSheet, tableObj);

                await ReportGenerator.RefreshListObjectAsync(tableObj.Name, paramsJsonPayload: paramsPayload);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
            }
        }

        private async void RibEdgeRefreshAll_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            const string MethodName = "RibEdgeRefreshAll_OnClick";

            if (string.IsNullOrEmpty(XLEdgeAppState.Instance.LoginUrl))
            {
                return;
            }

            if (ExcelApplicationHelper.IsCellInEditMode())
            {
                LogUtility.LogWarn($"{MethodName}|Active cell is in edit mode. Exit edit mode and try again.");
                return;
            }

            // Collects (sheetName, message) for every sheet that fails, so the loop can continue
            // processing the remaining sheets and report all failures together at the end.
            // RefreshListObjectAsync is called below with collectErrors: true so its validation
            // failures throw instead of showing their own UI, letting them be caught here uniformly.
            var errors = new List<(string SheetName, string Message)>();

            // A single wait window covers the whole book-refresh operation; its label is updated
            // per sheet, and each RefreshListObjectAsync call below passes useWaitWindow: false so
            // it doesn't create its own.
            XLEdgeWaitWindow waitWindow = null;

            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                var tableNames = new List<string>();
                var tableSheetsByName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                // NOTE: Workbook.Worksheets actually returns an Excel.Sheets COM object at runtime
                // (not Excel.Worksheets - that interface exists in the interop assembly but isn't what
                // this property implements), so casting/declaring it as Excel.Worksheets throws
                // InvalidCastException (E_NOINTERFACE) at runtime even though it compiles fine.
                Excel.Sheets sheets = excelApp.ActiveWorkbook.Worksheets;
                try
                {
                    foreach (Excel.Worksheet ws in sheets)
                    {
                        try
                        {
                            if (ws.ListObjects.Count > 0 && ws.ListObjects[1].Name.EndsWith("_E", StringComparison.Ordinal))
                            {
                                tableNames.Add(ws.ListObjects[1].Name);
                                tableSheetsByName[ws.ListObjects[1].Name] = ws.Name;
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(ws);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(sheets);
                }

                if (tableNames.Count == 0)
                {
                    MessageFunctions.XLEdgeMessage("No reports in the workbook to refresh!", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                LogUtility.LogDebug($"{MethodName}|Workbook refresh started for {tableNames.Count} report(s).");

                var waitCancelHelper = new CancellationHelper();
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        waitWindow = new XLEdgeWaitWindow(waitCancelHelper);
                        waitWindow.SetProcessTitle("Refreshing all reports", MahApps.Metro.IconPacks.PackIconFontAwesomeKind.SpinnerSolid);
                        waitWindow.SetProcessMessage($"Refreshing {tableNames.Count} report(s)...");
                        waitWindow.StartMonitoring();
                        waitWindow.Show();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "Failed to show wait window for book refresh");
                        waitWindow = null;
                    }
                });

                bool cancelled = false;

                foreach (string tableName in tableNames)
                {
                    if (cancelled)
                    {
                        break;
                    }

                    tableSheetsByName.TryGetValue(tableName, out string sheetName);
                    string displaySheetName = sheetName ?? tableName;

                    try
                    {
                        if (waitWindow != null)
                        {
                            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => waitWindow.SetProcessMessage($"Refreshing '{displaySheetName}'..."));
                        }

                        string paramsPayload = null;
                        if (sheetName != null)
                        {
                            Excel.Worksheet sht = null;
                            try
                            {
                                sht = excelApp.ActiveWorkbook.Worksheets[sheetName] as Excel.Worksheet;
                                if (sht != null && sht.ListObjects.Count > 0)
                                {
                                    // RefreshListObjectAsync resolves the sheet to use via
                                    // excelApp.ActiveSheet, so activate this sheet before calling it.
                                    sht.Activate();

                                    Excel.ListObject sheetTableObj = sht.ListObjects[1];

                                    // Skip reports executed under a different logged-in XLEdge instance
                                    // and skip child/drilldown reports - RefreshAll must not silently
                                    // pull someone else's session data or independently re-run a child
                                    // report.
                                    if (!TryResolveInstanceAndChildFlag(sht, sheetTableObj, tableName, out bool isChild))
                                    {
                                        LogUtility.LogDebug($"{MethodName}|{tableName}|Skipped - executed under a different logged-in instance.");
                                        continue;
                                    }

                                    if (isChild)
                                    {
                                        LogUtility.LogDebug($"{MethodName}|{tableName}|Skipped - child/drilldown report.");
                                        continue;
                                    }

                                    paramsPayload = BuildRefreshParamsPayload(excelApp.ActiveWorkbook, sht, sheetTableObj);
                                }
                            }
                            finally
                            {
                                if (sht != null)
                                {
                                    Marshal.ReleaseComObject(sht);
                                }
                            }
                        }

                        await ReportGenerator.RefreshListObjectAsync(tableName, useWaitWindow: false, paramsJsonPayload: paramsPayload, collectErrors: true);
                    }
                    catch (OperationCanceledException)
                    {
                        // Cancelling stops the whole book-wide refresh rather than just the current
                        // sheet, unlike a genuine per-sheet failure (see the catch below).
                        LogUtility.LogWarn($"{MethodName}|Book refresh cancelled by user.");
                        cancelled = true;
                    }
                    catch (Exception ex)
                    {
                        errors.Add((displaySheetName, ex.Message));
                        LogUtility.LogException(ex, $"{MethodName}|{tableName}");
                    }
                }

                // Close the wait window before showing the summary message so the summary doesn't
                // render behind it.
                if (waitWindow != null)
                {
                    var wwToClose = waitWindow;
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try { wwToClose.RequestClose(); } catch (Exception ex) { LogUtility.LogException(ex, "Failed to close book-refresh wait window"); }
                    });
                }

                if (!cancelled && errors.Count > 0)
                {
                    string summary = string.Join(Environment.NewLine, errors.Select(e => $"Sheet '{e.SheetName}' failed: {e.Message}"));
                    MessageFunctions.XLEdgeMessage(summary, MessageBoxIcon.Error, MessageBoxButtons.OK);
                }

                // RefreshListObjectAsync's own focus-release logic is bypassed by collectErrors:true,
                // so reclaim keyboard focus from XLEdgeCTP's WebView2 here for the whole operation.
                await ReportGenerator.ReleaseKeyboardFocusFromTaskPaneAsync();
            }
            catch (Exception ex)
            {
                if (waitWindow != null)
                {
                    var wwToClose = waitWindow;
                    try
                    {
                        await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            // Non-fatal - log and continue, since the real failure is logged below.
                            try { wwToClose.RequestClose(); } catch (Exception closeEx) { LogUtility.LogException(closeEx, $"{MethodName}: failed to close wait window while handling another exception"); }
                        });
                    }
                    catch (Exception dispatchEx)
                    {
                        LogUtility.LogException(dispatchEx, $"{MethodName}: failed to dispatch wait window close while handling another exception");
                    }
                }
                LogUtility.LogException(ex, MethodName);

                await ReportGenerator.ReleaseKeyboardFocusFromTaskPaneAsync();
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's RibEdgeParamRefreshBook_OnClick. Scans the workbook for all
        /// non-child "_E" report tables belonging to the currently logged-in instance, collects their
        /// run ids, and asks the hosted web app (via a DOM hook it exposes) to re-run all of them with
        /// their current parameters.
        /// </summary>
        private void RibEdgeParamRefreshBook_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            const string MethodName = "RibEdgeParamRefreshBook_OnClick";
            try
            {
                if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp) || excelApp.ActiveWorkbook == null)
                {
                    return;
                }

                var runIds = new List<string>();

                foreach (Excel.Worksheet ws in excelApp.ActiveWorkbook.Worksheets)
                {
                    try
                    {
                        if (ws.ListObjects.Count == 0)
                        {
                            continue;
                        }

                        Excel.ListObject tableObj = ws.ListObjects[1];
                        string tableName = tableObj.Name ?? string.Empty;

                        if (!tableName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) ||
                            !tableName.EndsWith("_E", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        if (!TryResolveInstanceAndChildFlag(ws, tableObj, tableName, out bool isChild))
                        {
                            continue; // Belongs to a different logged-in instance - skip it.
                        }

                        if (isChild)
                        {
                            continue;
                        }

                        string[] parts = tableName.Split('_');
                        if (parts.Length > 2)
                        {
                            runIds.Add(parts[2]);
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(ws);
                    }
                }

                if (runIds.Count == 0)
                {
                    return;
                }

                string refreshRunIds = string.Join("-", runIds);
                ADXExcelTaskPane1 paneInst = GetPaneInstance();
                if (paneInst == null)
                {
                    return;
                }

                try
                {
                    if (!paneInst.Visible)
                    {
                        XLEdgeAppState.Instance.EdgePaneShown = true;
                        LaunchExcelPane();
                    }

                    _ = RefreshBookParametersAsync(refreshRunIds, paneInst);
                }
                finally
                {
                    XLEdgeAppState.Instance.EdgePaneShown = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
            }
        }

        private static async Task RefreshBookParametersAsync(string runId, ADXExcelTaskPane1 taskPane)
        {
            try
            {
                if (taskPane == null)
                {
                    return;
                }

                await taskPane.ExecuteScriptAsync(
                    $"document.querySelector('[reruntype=xledgeworkbookrerun]').setAttribute('runIds', '{runId}');");
                await taskPane.ExecuteScriptAsync(
                    "document.querySelector('[reruntype=xledgeworkbookrerun]').click();");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RefreshBookParametersAsync));
            }
        }

        /// <summary>
        /// Ported from AddinModule.vb's RibEdgeParamRefresh_OnClick. Validates the active sheet's report
        /// table belongs to the current instance and isn't a child (drilldown) report, then asks the
        /// hosted web app to re-run it with its current parameters.
        /// </summary>
        private void RibEdgeParamRefresh_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            const string MethodName = "RibEdgeParamRefresh_OnClick";
            try
            {
                if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp))
                {
                    return;
                }

                Excel.Worksheet sht = excelApp.ActiveSheet as Excel.Worksheet;
                ADXExcelTaskPane1 paneInst = GetPaneInstance();

                if (sht == null || sht.ListObjects.Count == 0)
                {
                    if (RibEdgeLogout.Visible && paneInst != null)
                    {
                        XLEdgeAppState.Instance.EdgePaneShown = true;
                        LaunchExcelPane();
                    }
                    return;
                }

                Excel.ListObject tableObj = sht.ListObjects[1];
                string tableName = tableObj.Name ?? string.Empty;

                if (!tableName.EndsWith("_E", StringComparison.Ordinal))
                {
                    return;
                }

                if (!TryResolveInstanceAndChildFlag(sht, tableObj, tableName, out bool isChild, showMismatchMessage: true))
                {
                    return; // Instance mismatch - message already shown.
                }

                if (isChild)
                {
                    MessageFunctions.XLEdgeMessage("Re-Run wont work for the child reports.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    return;
                }

                string[] parts = tableName.Split('_');
                if (parts.Length <= 2)
                {
                    return;
                }

                string refreshRunId = parts[2];

                if (paneInst == null)
                {
                    return;
                }

                try
                {
                    if (!paneInst.Visible)
                    {
                        XLEdgeAppState.Instance.EdgePaneShown = true;
                        LaunchExcelPane();
                    }

                    _ = RefreshParametersAsync(refreshRunId, paneInst);
                }
                finally
                {
                    XLEdgeAppState.Instance.EdgePaneShown = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
            }
        }

        private static async Task RefreshParametersAsync(string runId, ADXExcelTaskPane1 taskPane)
        {
            try
            {
                if (taskPane == null)
                {
                    return;
                }

                await taskPane.ExecuteScriptAsync(
                    $"document.getElementById('XLEdgeParamRefresh').setAttribute('runId', '{runId}');");
                await taskPane.ExecuteScriptAsync(
                    "document.getElementById('XLEdgeParamRefresh').click();");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RefreshParametersAsync));
            }
        }

        /// <summary>
        /// Shared helper for RibEdgeParamRefreshBook_OnClick/RibEdgeParamRefresh_OnClick: reads the
        /// "IT5" (executed-instance login URL) and "IT1" ("Child Report" marker) cells from the table's
        /// companion parameter sheet (or the table's own sheet, if there's no separate parameter sheet),
        /// matching the VB original's inline logic in both handlers. Returns false if the report was
        /// executed against a different XLEdge instance than the current login (the caller should skip
        /// it silently in the "book" handler, or show a message in the single-sheet handler).
        /// </summary>
        private static bool TryResolveInstanceAndChildFlag(Excel.Worksheet sheet, Excel.ListObject tableObj, string tableName, out bool isChild, bool showMismatchMessage = false)
        {
            isChild = false;
            Excel.Worksheet sourceSheet = sheet;
            bool releaseSourceSheet = false;

            try
            {
                if (tableObj.HeaderRowRange != null && tableObj.HeaderRowRange.Offset[1, 0].Row == 2)
                {
                    string paramSheetName = $"P_{sheet.Name}";
                    Excel.Worksheet paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableName);
                    if (paramSheet == null)
                    {
                        return true;
                    }

                    sourceSheet = paramSheet;
                    releaseSourceSheet = true;
                }

                object it5 = null;
                try { it5 = sourceSheet.Range["IT5"]?.Value; }
                catch (Exception ex)
                {
                    // Safe to ignore/expected: IT5 cell may not exist on older/differently-shaped
                    // sheets; treated the same as "no instance mismatch info available".
                    LogUtility.LogDebug($"{nameof(TryResolveInstanceAndChildFlag)}: failed to read IT5 cell - {ex.Message}");
                }

                if (it5 != null && !string.Equals(Convert.ToString(it5), XLEdgeAppState.Instance.LoginUrl, StringComparison.Ordinal))
                {
                    if (showMismatchMessage)
                    {
                        MessageFunctions.XLEdgeMessage(
                            $"Current user logged-in instance and report executed instance, both are different.{Environment.NewLine}" +
                            $"Current logged instance : {XLEdgeAppState.Instance.LoginUrl}{Environment.NewLine}" +
                            $"Report executed instance : {it5}",
                            MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                    }

                    return false;
                }

                try
                {
                    object it1 = sourceSheet.Range["IT1"]?.Value;
                    isChild = it1 != null && string.Equals(Convert.ToString(it1), "Child Report", StringComparison.OrdinalIgnoreCase);
                }
                catch (Exception ex)
                {
                    // Safe to ignore/expected: IT1 cell may not exist on older/differently-shaped
                    // sheets; treated the same as "not a child report".
                    LogUtility.LogDebug($"{nameof(TryResolveInstanceAndChildFlag)}: failed to read IT1 cell - {ex.Message}");
                    isChild = false;
                }

                return true;
            }
            finally
            {
                if (releaseSourceSheet && sourceSheet != null)
                {
                    Marshal.ReleaseComObject(sourceSheet);
                }
            }
        }

        /// <summary>
        /// Ported from the parameter-resolution half of AddinModule.vb's RibEdgeRefresh_OnClick:
        /// resolves the table's companion parameter sheet (or the table's own sheet, if there's no
        /// separate one) and builds the edited-parameters JSON payload from the shared
        /// "orb_params_control" sheet via XLEdgeParamsBuilder.BuildParamData, so refreshing a report
        /// picks up any filter-value edits the user made there instead of always re-fetching with the
        /// report's original parameters. Returns null/empty if there's nothing to send (falls back to
        /// a plain refresh).
        /// </summary>
        private static string BuildRefreshParamsPayload(Excel.Workbook workbook, Excel.Worksheet sheet, Excel.ListObject tableObj)
        {
            if (workbook == null || sheet == null || tableObj == null)
            {
                return null;
            }

            Excel.Worksheet paramSheet = null;
            bool releaseParamSheet = false;

            try
            {
                if (tableObj.HeaderRowRange != null && tableObj.HeaderRowRange.Offset[1, 0].Row == 2)
                {
                    string paramSheetName = $"P_{sheet.Name}";
                    paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableObj.Name);
                    releaseParamSheet = true;
                }
                else
                {
                    paramSheet = sheet;
                }

                if (paramSheet == null)
                {
                    LogUtility.LogWarn("BuildRefreshParamsPayload|Parameter sheet not found");
                    return null;
                }

                // Build the params payload

                // Clear cached data to ensure fresh read from control sheet
                XLEdgeAppState.Instance.ClearCachedRefreshData();

                (string paramsPayload, string mergedParamsJson) = XLEdgeParamsBuilder.BuildParamData(workbook, paramSheet, tableObj);

                // Store the richly-merged params JSON (not the bare CSV-API request payload) as the
                // persisted/display source, and cache the report's meta JSON alongside it.
                if (!string.IsNullOrEmpty(paramsPayload))
                {
                    XLEdgeAppState.Instance.UpdatedParamData = mergedParamsJson;

                    if (ReportGenerator.TryGetStoredReportXml(workbook, tableObj.Name, out _, out string metaJson, out _))
                    {
                        XLEdgeAppState.Instance.UpdatedMetaData = metaJson;
                    }
                }

                return paramsPayload;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(BuildRefreshParamsPayload));
                return null;
            }
            finally
            {
                if (releaseParamSheet && paramSheet != null)
                {
                    Marshal.ReleaseComObject(paramSheet);
                }
            }
        }

        private void RibControlSheet_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            try
            {
                ParamsControlSheetBuilder.ShowOrRebuild();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RibControlSheet_OnClick));
            }
        }

        private async void adxExcelAppEvents1_SheetFollowHyperlink(object sender, object sheet, object hyperlink)
        {
            const string MethodName = "adxExcelAppEvents1_SheetFollowHyperlink";

            if (!(sheet is Excel.Worksheet dataSheet) || !(hyperlink is Excel.Hyperlink hyperLink))
            {
                return;
            }

            string cellText;
            try { cellText = hyperLink.Range.Text as string ?? string.Empty; }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort read of the hyperlink cell's display text; treated as
                // "not the Goto-Report-Data link" if unreadable.
                LogUtility.LogDebug($"{MethodName}: failed to read hyperlink cell text - {ex.Message}");
                cellText = string.Empty;
            }

            if (cellText == "Goto Report Data")
            {
                NavigateToReportDataSheet(dataSheet);
                return;
            }

            if (dataSheet.ListObjects.Count == 0)
            {
                return;
            }

            Excel.ListObject tableObj = dataSheet.ListObjects[1];
            string tableObjName = tableObj.Name;

            if (tableObjName.IndexOf("ORB_", StringComparison.Ordinal) < 0 ||
                tableObjName.IndexOf("ORB_DD_", StringComparison.Ordinal) >= 0 ||
                tableObjName.IndexOf("ORB_XLDD_", StringComparison.Ordinal) >= 0)
            {
                return;
            }

            if (!XLEdgeAppState.Instance.IsLoginCompleted)
            {
                MessageFunctions.XLEdgeMessage("Please login to the instance and try again!", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            string screenTip;
            try { screenTip = hyperLink.ScreenTip?.ToLowerInvariant() ?? string.Empty; }
            catch (Exception ex)
            {
                // Worth investigating if this recurs - ScreenTip is how this handler determines what
                // action to take (attachment/drilldown); failing to read it means the click is a no-op.
                LogUtility.LogWarn($"{MethodName}: failed to read hyperlink ScreenTip - {ex.Message}");
                screenTip = string.Empty;
            }

            if (string.IsNullOrEmpty(screenTip))
            {
                LogUtility.LogWarn($"{MethodName}|Hyperlink does not contain screen tip information. Cannot determine action to take.");
                return;
            }

            if (screenTip.Contains("attachment|"))
            {
                await HandleAttachmentDownloadAsync(hyperLink.ScreenTip);
                return;
            }

            if (!screenTip.Contains("drilldown|"))
            {
                return;
            }

            // Ported from AddinModule.vb's SheetFollowHyperlink: resolve the companion parameter sheet
            // (or the data sheet itself in same-sheet mode) and refuse to drill down if the report was
            // executed under a different logged-in XLEdge instance than the current login - reuses the
            // same IT5/IT1 helper the Refresh handlers already use (isChild is irrelevant here and discarded).
            if (!TryResolveInstanceAndChildFlag(dataSheet, tableObj, tableObjName, out _, showMismatchMessage: true))
            {
                return;
            }

            Excel.Workbook workbook = dataSheet.Parent as Excel.Workbook;

            if (!ReportGenerator.TryGetStoredReportXml(workbook, tableObjName, out _, out string metaJson, out string storedParamsJson))
            {
                MessageFunctions.XLEdgeMessage("Reports metadata information missing. Please rerun the report to generate.", MessageBoxIcon.Exclamation, MessageBoxButtons.OK);
                return;
            }

            ReportMeta reportMeta;
            try
            {
                reportMeta = System.Text.Json.JsonSerializer.Deserialize<ReportMeta>(metaJson, JsonGlobals.Options);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
                return;
            }

            if (reportMeta?.Drilldowns == null || reportMeta.Drilldowns.Length == 0)
            {
                return;
            }

            string columnName;
            Excel.Range clickedRange;
            try
            {
                clickedRange = dataSheet.Range[hyperLink.SubAddress];
                Excel.Range headerCell = (Excel.Range)dataSheet.Cells[tableObj.HeaderRowRange.Row, clickedRange.Column];
                columnName = Convert.ToString(headerCell.Value) ?? string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
                return;
            }

            if (string.IsNullOrEmpty(columnName))
            {
                LogUtility.LogWarn($"{MethodName}|Column name is empty. Cannot proceed with drilldown.");
                return;
            }

            List<string> candidates = reportMeta.Drilldowns
                .Where(d => string.Equals(d.DrillColumnName?.Trim(), columnName.Trim(), StringComparison.OrdinalIgnoreCase))
                .Select(d => $"DRILLDOWN|{d.DrillReportId}|{d.DrillReportName}|{reportMeta.ReportId}")
                .Distinct()
                .ToList();

            if (candidates.Count == 0)
            {
                LogUtility.LogWarn($"{MethodName}|No drilldown information found for column '{columnName}'.");
                return;
            }

            string selected;
            if (candidates.Count > 1)
            {
                string pickedResult = null;
                SafeInvokeWpf(() =>
                {
                    var picker = new XLEdgeDrilldownReports { DrillRptsList = candidates };
                    picker.ShowDialog();
                    pickedResult = picker.DrillSelRpt;
                });
                selected = pickedResult;
            }
            else
            {
                selected = candidates[0];
            }

            if (string.IsNullOrEmpty(selected))
            {
                return;
            }

            string[] parts = selected.Split('|');
            if (parts.Length < 3)
            {
                return;
            }

            string childReportId = parts[1];
            string childReportName = parts[2];
            string childTitle = $"Edge|{childReportId}|{childReportId}|{childReportName}";

            // Ported from AddinModule.vb's "Strs.Length >= 4" block: resolve the same companion
            // parameter sheet used for the IT5 check above (or the data sheet itself in same-sheet
            // mode) so the "Responsibility" (IT4/ORACLE_RESP_ID) extra parameter can be attached, then
            // build the scoped drilldown request from the clicked row's PARAM/STATIC/cell-value
            // parameter mapping - matching the VB original instead of re-running the child unfiltered.
            Excel.Worksheet parameterSheetForExtras = dataSheet;
            bool releaseParamSheetForExtras = false;
            try
            {
                if (tableObj.HeaderRowRange != null && tableObj.HeaderRowRange.Offset[1, 0].Row == 2)
                {
                    Excel.Worksheet resolvedParamSheet = ExcelSheetHelper.GetParameterSheet($"P_{dataSheet.Name}", tableObjName);
                    if (resolvedParamSheet != null)
                    {
                        parameterSheetForExtras = resolvedParamSheet;
                        releaseParamSheetForExtras = true;
                    }
                }

                string drilldownPayload = DrilldownRequestBuilder.BuildDrilldownRequestJson(
                    reportMeta, storedParamsJson, childReportId, columnName,
                    tableObj.HeaderRowRange, dataSheet, clickedRange, parameterSheetForExtras);

                await ReportGenerator.CreateReportFromTitleAsync(childTitle, useWaitWindow: true, paramsJsonPayload: drilldownPayload);
            }
            finally
            {
                if (releaseParamSheetForExtras && parameterSheetForExtras != null)
                {
                    Marshal.ReleaseComObject(parameterSheetForExtras);
                }
            }
        }

        /// <summary>
        /// Ported from the attachment-click branch of AddinModule.vb's SheetFollowHyperlink
        /// (ScreenTip.IndexOf("attachment") >= 0 -> DownloadFile). Downloads the file to the user's
        /// Downloads folder and confirms with a message box, matching the VB original.
        /// </summary>
        private static async Task HandleAttachmentDownloadAsync(string screenTip)
        {
            const string MethodName = "HandleAttachmentDownloadAsync";
            try
            {
                string url = AttachmentLinkHelper.BuildDownloadUrl(screenTip, XLEdgeAppState.Instance.LoginUrl);
                if (string.IsNullOrWhiteSpace(url))
                {
                    LogUtility.LogWarn($"{MethodName}|Could not build a download URL from ScreenTip '{screenTip}'.");
                    return;
                }

                string savedPath = await ApiHelper.DownloadFileAsync(url);
                if (string.IsNullOrWhiteSpace(savedPath))
                {
                    MessageFunctions.XLEdgeMessage("Failed to download the attachment.", MessageBoxIcon.Error, MessageBoxButtons.OK);
                    return;
                }

                string fileName = System.IO.Path.GetFileName(savedPath);
                MessageFunctions.XLEdgeMessage(
                    $"Attachment has been saved to the downloads folder and the file name is \"{fileName}\"",
                    MessageBoxIcon.Information, MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, MethodName);
            }
        }

        private static void NavigateToReportDataSheet(Excel.Worksheet dataSheet)
        {
            try
            {
                object it2 = dataSheet.Range["IT2"].Value;
                if (it2 == null)
                {
                    return;
                }

                string tableName = Convert.ToString(it2);
                Excel.Workbook workbook = dataSheet.Parent as Excel.Workbook;
                if (workbook == null)
                {
                    return;
                }

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    bool release = true;
                    try
                    {
                        if (ws.ListObjects.Count > 0 && string.Equals(ws.ListObjects[1].Name, tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws.Activate();
                            release = false;
                            return;
                        }
                    }
                    finally
                    {
                        if (release)
                        {
                            Marshal.ReleaseComObject(ws);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(NavigateToReportDataSheet));
            }
        }

        /// <summary>
        /// Locates the sibling "GLSense" Add-in Express add-in (a related product) via Excel's
        /// COMAddIns collection and returns its automation object, so XLEdge can call into it
        /// directly. Releases every COMAddIn wrapper encountered while iterating, not just the
        /// one that matched.
        /// </summary>
        public object GetGLSenseAddinObject()
        {
            Excel._Application app;
            try
            {
                app = Marshal.GetActiveObject("Excel.Application") as Excel._Application;
            }
            catch (Exception ex)
            {
                // Safe to ignore/expected: no running Excel instance registered in the ROT, or COM
                // interop failure resolving it - caller treats null as "GLSense add-in not available".
                LogUtility.LogDebug($"{nameof(GetGLSenseAddinObject)}: failed to get active Excel Application via ROT - {ex.Message}");
                return null;
            }

            if (app == null)
            {
                return null;
            }

            try
            {
                Microsoft.Office.Core.COMAddIns addins = app.COMAddIns;
                if (addins == null)
                {
                    return null;
                }

                try
                {
                    foreach (Microsoft.Office.Core.COMAddIn candidate in addins)
                    {
                        try
                        {
                            if (candidate.ProgId != "GLSense.AddinModule")
                            {
                                continue;
                            }

                            Microsoft.Office.Core.COMAddIn addin = null;
                            try
                            {
                                addin = addins.Item("GLSense.AddinModule");
                            }
                            catch (Exception ex)
                            {
                                // Safe to ignore/expected: GLSense COM add-in is not registered on this
                                // machine - caller treats null as "GLSense add-in not available".
                                LogUtility.LogDebug($"{nameof(GetGLSenseAddinObject)}: GLSense.AddinModule not registered - {ex.Message}");
                            }

                            if (addin != null)
                            {
                                try
                                {
                                    if (addin.Connect)
                                    {
                                        return addin.Object;
                                    }
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(addin);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(candidate);
                        }
                    }

                    return null;
                }
                finally
                {
                    Marshal.ReleaseComObject(addins);
                }
            }
            catch (Exception ex)
            {
                // Worth investigating if this recurs - a real failure walking Excel's COMAddIns
                // collection; caller treats null as "GLSense add-in not available".
                LogUtility.LogWarn($"{nameof(GetGLSenseAddinObject)}: failed to resolve GLSense add-in via COMAddIns - {ex.Message}");
                return null;
            }
            finally
            {
                Marshal.ReleaseComObject(app);
            }
        }

    }
}

