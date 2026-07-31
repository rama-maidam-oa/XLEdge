using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    public sealed class XLEdgeRibbonHelper
    {
        private const string RibEdgeDialogBoxLauncher = "RibEdgeDialogBoxLauncher";
        private const string RibEdgeRefresh = "RibEdgeRefresh";
        private const string RibEdgeRefreshAll = "RibEdgeRefreshAll";
        private const string RibEdgeParamRefresh = "RibEdgeParamRefresh";
        private const string RibEdgeShowHide = "RibEdgeShowHide";
        private const string RibEdgeIncludeOutputData = "RibEdgeIncludeOutputData";
        private const string RibEdgeOptions = "RibEdgeOptions";
        private const string RibEdgeHelp = "RibEdgeHelp";
        private const string RibEdgeLogin = "RibEdgeLogin";
        private const string EibEdgeLogout = "RibEdgeLogout";
        private const string RibEdgeDebug = "RibEdgeDebug";
        private const string RibControlSheet = "RibControlSheet";
        private const string RibEdgeAbout = "RibEdgeAbout";

        private static readonly string[] LoggedOutDisabledControls =
        {
            RibEdgeDialogBoxLauncher,
            RibEdgeRefresh,
            RibEdgeRefreshAll,
            RibEdgeParamRefresh,
            RibEdgeShowHide,
            RibEdgeIncludeOutputData,
            RibEdgeOptions,
            RibEdgeHelp
        };

        private static readonly string[] LoggedOutEnabledControls =
        {
            RibEdgeLogin,
            RibEdgeDebug,
            RibControlSheet,
            RibEdgeAbout
        };

        private static readonly string[] LoggedInEnabledControls =
        {
            RibEdgeDialogBoxLauncher,
            RibEdgeRefresh,
            RibEdgeRefreshAll,
            RibEdgeParamRefresh,
            RibEdgeShowHide,
            RibEdgeIncludeOutputData,
            RibEdgeOptions,
            RibEdgeHelp
        };

        private static readonly string[] NoPermissionDisabledControls =
        {
            RibEdgeDialogBoxLauncher,
            RibEdgeRefresh,
            RibEdgeRefreshAll,
            RibEdgeParamRefresh,
            RibEdgeShowHide,
            RibEdgeIncludeOutputData,
            RibEdgeOptions,
            RibEdgeHelp,
            RibEdgeLogin,
            RibEdgeDebug,
            RibControlSheet,
            RibEdgeAbout
        };

        private const string StateLoggedOut = "LoggedOut";
        private const string StateLoggedIn = "LoggedIn";
        private const string StateSheetActive = "ApplySheetActiveState";
        private const string StateNoPermission = "NoXLEdePermission";

        private readonly AddinModule _addinModule;
        private readonly AddinExpress.MSO.IRibbonUI _ribbon;
        private readonly Dictionary<string, bool> _enabledStates;

        public static XLEdgeRibbonHelper Current { get; private set; }
        public XLEdgeRibbonHelper(AddinModule addinModule, AddinExpress.MSO.IRibbonUI ribbon)
        {
            _addinModule = addinModule;
            _ribbon = ribbon;
            _enabledStates = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            Current = this;
        }

        public void SetControlPressed(string controlName, bool pressed)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo propPressed = ctrl.GetType().GetProperty("Pressed");
                if (propPressed != null)
                {
                    propPressed.SetValue(ctrl, pressed, null);
                    return;
                }

                PropertyInfo propChecked = ctrl.GetType().GetProperty("Checked");
                if (propChecked != null)
                {
                    propChecked.SetValue(ctrl, pressed, null);
                    return;
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Pressed/Checked property on control '{controlName}'.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] SetControlPressed: {controlName}");
            }
        }
        public bool GetControlPressed(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return false;

                // Try Pressed first
                PropertyInfo propPressed = ctrl.GetType().GetProperty("Pressed");
                if (propPressed != null)
                {
                    return (bool)propPressed.GetValue(ctrl);
                }

                // Try Checked as fallback
                PropertyInfo propChecked = ctrl.GetType().GetProperty("Checked");
                if (propChecked != null)
                {
                    return (bool)propChecked.GetValue(ctrl);
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Pressed/Checked property on control '{controlName}'.");
                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetControlPressed: {controlName}");
                return false;
            }
        }
        public void SetControlEnabled(string controlName, bool enabled)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo prop = ctrl.GetType().GetProperty("Enabled");
                if (prop != null)
                {
                    prop.SetValue(ctrl, enabled, null);
                    _enabledStates[controlName] = enabled;
                    return;
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Enabled property on control '{controlName}'.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] SetControlEnabled: {controlName}");
            }
        }
        public bool GetControlEnabled(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return false;

                PropertyInfo prop = ctrl.GetType().GetProperty("Enabled");
                if (prop != null)
                {
                    return (bool)prop.GetValue(ctrl);
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Enabled property on control '{controlName}'.");
                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetControlEnabled: {controlName}");
                return false;
            }
        }
        public void SetControlCaption(string controlName, string caption)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo prop = ctrl.GetType().GetProperty("Caption");
                if (prop != null)
                {
                    prop.SetValue(ctrl, caption, null);
                    return;
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Caption property on control '{controlName}'.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] SetControlVisible: {controlName}");
            }
        }
        public string GetControlCaption(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return string.Empty;

                PropertyInfo prop = ctrl.GetType().GetProperty("Caption");
                if (prop != null)
                {
                    return prop.GetValue(ctrl)?.ToString() ?? string.Empty;
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Caption property on control '{controlName}'.");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetControlCaption: {controlName}");
                return string.Empty;
            }
        }
        public void SetControlVisible(string controlName, bool visible)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return;

                PropertyInfo prop = ctrl.GetType().GetProperty("Visible");
                prop?.SetValue(ctrl, visible, null);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] SetControlVisible: {controlName}");
            }
        }
        public bool GetControlVisible(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return false;

                PropertyInfo prop = ctrl.GetType().GetProperty("Visible");
                if (prop != null)
                {
                    return (bool)prop.GetValue(ctrl);
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Visible property on control '{controlName}'.");
                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetControlVisible: {controlName}");
                return false;
            }
        }
        public string GetControlLabel(string controlName)
        {
            try
            {
                object ctrl = GetRibbonControl(controlName);
                if (ctrl == null)
                    return string.Empty;

                PropertyInfo prop = ctrl.GetType().GetProperty("Label");
                if (prop != null)
                {
                    return prop.GetValue(ctrl)?.ToString() ?? string.Empty;
                }

                LogUtility.LogWarn($"[XLEdgeRibbonHelper] No Label property on control '{controlName}'.");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetControlLabel: {controlName}");
                return string.Empty;
            }
        }
        public void EnableControls(IEnumerable<string> controlNames)
        {
            if (controlNames == null)
                return;

            foreach (string name in controlNames)
                SetControlEnabled(name, true);
        }

        public void DisableControls(IEnumerable<string> controlNames)
        {
            if (controlNames == null)
                return;

            foreach (string name in controlNames)
                SetControlEnabled(name, false);
        }

        public void RestorePreviousState()
        {
            foreach (KeyValuePair<string, bool> kvp in _enabledStates)
                SetControlEnabled(kvp.Key, kvp.Value);
        }

        private void RefreshRibbon()
        {
            try
            {
                _ribbon?.Invalidate();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "[XLEdgeRibbonHelper] RefreshRibbon");
            }
        }

        public void ApplyState(string stateName)
        {
            if (string.IsNullOrWhiteSpace(stateName))
            {
                LogUtility.LogWarn("[XLEdgeRibbonHelper] ApplyState called with an empty state name.");
                return;
            }

            switch (stateName)
            {
                case StateLoggedOut:
                    ApplyLoggedOutState();
                    break;
                case StateLoggedIn:
                    ApplyLoggedInState();
                    break;

                case StateSheetActive:
                    ApplySheetActiveState();
                    break;

                case StateNoPermission:
                    ApplyNoPermissionState();
                    break;

                default:
                    LogUtility.LogWarn($"[XLEdgeRibbonHelper] Unknown state: {stateName}");
                    break;
            }

            RefreshRibbon();
        }

        public void ApplyWorkbookActiveState(Excel.Workbook workbook)
        {
            if (!XLEdgeAppState.Instance.IsLoginCompleted)
                return;

            try
            {
                ProcessActiveWorkbook(workbook);
            }
            catch (Exception ex)
            {
                LogUtility.LogError("[XLEdgeRibbonHelper] ApplyWorkbookActiveState: " + ex.Message);
            }

            RefreshRibbon();
        }

        private void ApplyLoggedOutState()
        {
            try
            {
                DisableControls(LoggedOutDisabledControls);
                EnableControls(LoggedOutEnabledControls);

                SetControlVisible(RibEdgeLogin, true);
                SetControlVisible(EibEdgeLogout, false);
                SetControlPressed(RibEdgeDebug, false);
                SetControlPressed(RibEdgeIncludeOutputData, false);

                // Login always shows the plain "Login" caption, never an instance name.
                SetControlCaption(RibEdgeLogin, "Login");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void ApplyNoPermissionState()
        {
            try
            {
                SetControlPressed(RibEdgeIncludeOutputData, false);
                SetControlPressed(RibEdgeDebug, false);
                DisableControls(NoPermissionDisabledControls);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void ApplyLoggedInState()
        {
            try
            {
                EnableControls(LoggedInEnabledControls);

                SetControlVisible(RibEdgeLogin, false);
                SetControlVisible(EibEdgeLogout, true);
                SetControlPressed(RibEdgeIncludeOutputData, false);

                // Logout shows the selected instance name.
                SetControlCaption(EibEdgeLogout, XLEdgeAppState.Instance.LoginUrlName);

                ApplySheetActiveState();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void ApplySheetActiveState()
        {
            if (!XLEdgeAppState.Instance.IsLoginCompleted)
                return;

            try
            {
                ProcessActiveWorkbook(null);
            }
            catch (Exception ex)
            {
                LogUtility.LogError("[XLEdgeRibbonHelper] ApplySheetActiveState: " + ex.Message);
            }
        }

        private void ProcessActiveWorkbook(Excel.Workbook workbook)
        {
            if (!ExcelApplicationHelper.TryGetActiveExcelApplication(out Excel.Application excelApp) || excelApp == null)
                return;

            Excel.Workbook activeWorkbook = workbook ?? GetActiveWorkbook(excelApp);
            if (activeWorkbook == null)
                return;

            Excel.Worksheet activeSheet = GetActiveWorksheet(activeWorkbook);
            if (activeSheet == null)
                return;

            bool bookHasReport = BookHasEdgeReport(activeWorkbook);

            // RibEdgeParamRefresh always mirrors RibEdgeRefresh's enabled state - both act on the active sheet's report table.
            if (activeSheet.ListObjects.Count == 0)
            {
                DisableControls([RibEdgeRefresh, RibEdgeParamRefresh, RibEdgeRefreshAll]);
                if (bookHasReport)
                {
                    EnableControls([RibEdgeRefreshAll]);
                }

                return;
            }

            string listObjectName = activeSheet.ListObjects[1].Name;

            // Only a real, refreshable report table - "ORB_{reportId}_{runId}_E" - enables Refresh/Param Refresh.
            bool isRefreshableReportTable = !string.IsNullOrWhiteSpace(listObjectName) &&
                listObjectName.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) &&
                listObjectName.EndsWith("_E", StringComparison.Ordinal) &&
                !listObjectName.Equals("orb_params_control", StringComparison.OrdinalIgnoreCase);

            if (!isRefreshableReportTable)
            {
                DisableControls([RibEdgeRefresh, RibEdgeParamRefresh]);
                if (bookHasReport)
                {
                    EnableControls([RibEdgeRefreshAll]);
                }
                else
                {
                    DisableControls([RibEdgeRefreshAll]);
                }

                return;
            }

            // A drilldown-generated ("Child Report") sheet must never allow Refresh/Param Refresh on
            // itself - matches the same IT1 check RibEdgeRefresh_OnClick/RibEdgeParamRefresh_OnClick
            // already enforce reactively (after the click). This makes the ribbon reflect that
            // up front instead of only blocking the action after the user has already clicked it.
            if (IsChildReportSheet(activeSheet, activeSheet.ListObjects[1], listObjectName))
            {
                DisableControls([RibEdgeRefresh, RibEdgeParamRefresh]);
                if (bookHasReport)
                {
                    EnableControls([RibEdgeRefreshAll]);
                }
                else
                {
                    DisableControls([RibEdgeRefreshAll]);
                }

                return;
            }

            EnableControls([RibEdgeRefresh, RibEdgeParamRefresh, RibEdgeRefreshAll]);
        }

        /// <summary>
        /// Read-only check for whether <paramref name="sheet"/>'s report table is a drilldown-generated
        /// "Child Report" - mirrors the sheet-resolution and IT1 read in AddinModule's
        /// TryResolveInstanceAndChildFlag, but without that method's instance-mismatch check or
        /// MessageBox popup, since this runs silently on every sheet/workbook activation and after
        /// every report run - a popup here would fire far more often than the user clicking Refresh.
        /// </summary>
        private static bool IsChildReportSheet(Excel.Worksheet sheet, Excel.ListObject tableObj, string tableName)
        {
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
                        return false;
                    }

                    sourceSheet = paramSheet;
                    releaseSourceSheet = true;
                }

                try
                {
                    object it1 = sourceSheet.Range["IT1"]?.Value;
                    return it1 != null && string.Equals(Convert.ToString(it1), "Child Report", StringComparison.OrdinalIgnoreCase);
                }
                catch (Exception ex)
                {
                    // Safe to ignore/expected: IT1 cell may not exist on older/differently-shaped
                    // sheets; treated the same as "not a child report".
                    LogUtility.LogDebug($"{nameof(IsChildReportSheet)}: failed to read IT1 cell - {ex.Message}");
                    return false;
                }
            }
            finally
            {
                if (releaseSourceSheet && sourceSheet != null)
                {
                    Marshal.ReleaseComObject(sourceSheet);
                }
            }
        }

        private static bool BookHasEdgeReport(Excel.Workbook workbook)
        {
            if (workbook == null)
                return false;

            // Releases each iterated Worksheet COM object explicitly since this runs on every sheet/workbook activation.
            Excel.Sheets sheets = workbook.Worksheets;
            try
            {
                foreach (Excel.Worksheet sheet in sheets)
                {
                    try
                    {
                        if (sheet != null &&
                            sheet.ListObjects.Count > 0 &&
                            sheet.ListObjects[1].Name.StartsWith("ORB_", StringComparison.OrdinalIgnoreCase) &&
                            sheet.ListObjects[1].Name.EndsWith("_E", StringComparison.Ordinal) &&
                            !sheet.ListObjects[1].Name.Equals("orb_params_control", StringComparison.OrdinalIgnoreCase) &&
                            !IsChildReportSheet(sheet, sheet.ListObjects[1], sheet.ListObjects[1].Name))
                        {
                            // RibEdgeRefreshAll_OnClick only ever collects "_E" (live Edge) tables to
                            // refresh - it ignores "_P" (scheduled/Process) tables entirely - so a book
                            // made up only of scheduled-output sheets, or only of child (drilldown)
                            // reports, has nothing it would actually refresh; those sheets don't count
                            // towards "this book has a refreshable report" here either.
                            return true;
                        }
                    }
                    finally
                    {
                        if (sheet != null)
                        {
                            Marshal.ReleaseComObject(sheet);
                        }
                    }
                }

                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(sheets);
            }
        }

        private static Excel.Workbook GetActiveWorkbook(Excel.Application excelApp)
        {
            return excelApp?.ActiveWorkbook;
        }

        private static Excel.Worksheet GetActiveWorksheet(Excel.Workbook workbook)
        {
            return workbook?.ActiveSheet as Excel.Worksheet;
        }

        private object GetRibbonControl(string controlName)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(controlName) || _addinModule == null)
                    return null;

                return XLEdgeRibbonReflectionHelper.GetRibbonControl(_addinModule, controlName);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"[XLEdgeRibbonHelper] GetRibbonControl: {controlName}");
                return null;
            }
        }
    }
}
