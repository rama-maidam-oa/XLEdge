using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge
{
    /// <summary>
    /// The ribbon actions and Excel events the shell forwards, mapped to 11.1.2's handlers (AddinModule.cs). The shell sends
    /// primitives - sheet and workbook names, cell addresses, hyperlink texts - so the COM objects 11.1.2's handlers received from
    /// Add-in Express are looked up again here, through the core's own Excel RCW.
    /// </summary>
    public partial class AddinModule
    {
        /// <summary>Ribbon actions whose handlers need the task pane or the login (Stage 2e); the AddinEntry spike handles Login/ShowHide.</summary>
        private static readonly HashSet<string> NotYetInTheCore = new HashSet<string>(StringComparer.Ordinal)
        {
            "RibEdgeLogout", "RibEdgeParamRefresh", "RibEdgeParamRefreshBook", "RibEdgeAbout",
        };

        /// <summary>A forwarded ribbon click; false when no 11.1.2 handler takes it.</summary>
        internal bool OnRibbonAction(string action, object parameter)
        {
            bool pressed = parameter is bool b && b;
            switch (action)
            {
                case "RibEdgeRefresh": RibEdgeRefresh_OnClick(null, null, pressed); return true;
                case "RibEdgeRefreshAll": RibEdgeRefreshAll_OnClick(null, null, pressed); return true;
                case "RibEdgeOptions": RibEdgeOptions_OnClick(null, null, pressed); return true;
                case "RibEdgeDialogBoxLauncher": RibEdgeDialogBoxLauncher_OnAction(null, null, pressed); return true;
                case "RibEdgeHelp": RibEdgeHelp_OnClick(null, null, pressed); return true;
                case "RibControlSheet": RibControlSheet_OnClick(null, null, pressed); return true;
                case "RibEdgeIncludeOutputData": RibEdgeIncludeOutputData_OnClick(null, null, pressed); return true;
                case "RibEdgeDebug": RibEdgeDebug_OnClick(null, null, pressed); return true;
            }
            if (NotYetInTheCore.Contains(action ?? string.Empty))
                LogUtility.LogInfo($"AddinModule: ribbon action '{action}' is not in the core yet (Stage 2e).");
            else
                LogUtility.LogWarn($"AddinModule: no handler for ribbon action '{action}'.");
            return false;
        }

        /// <summary>
        /// A forwarded Excel event. Arguments (from the shell): SheetSelectionChange {sheet, address, cell count, workbook},
        /// SheetActivate / SheetBeforeDelete {sheet, workbook}, SheetFollowHyperlink {sheet, cell, screen tip, sub-address, text,
        /// workbook}, WorkbookActivate {workbook}. The workbook comes last, so an older shell that does not send it still works
        /// (the active workbook is used then).
        /// </summary>
        internal void OnExcelEvent(string eventName, object[] args)
        {
            args = args ?? new object[0];
            switch (eventName)
            {
                case "SheetSelectionChange":
                {
                    Excel.Worksheet sheet = Sheet(Arg(args, 0), Arg(args, 3));
                    Excel.Range range = Range(sheet, Arg(args, 1));
                    if (sheet != null && range != null) adxExcelAppEvents1_SheetSelectionChange(null, sheet, range);
                    return;
                }
                case "SheetActivate":
                {
                    Excel.Worksheet sheet = Sheet(Arg(args, 0), Arg(args, 1));
                    if (sheet != null) adxExcelAppEvents1_SheetActivate(null, sheet);
                    return;
                }
                case "SheetBeforeDelete":
                {
                    Excel.Worksheet sheet = Sheet(Arg(args, 0), Arg(args, 1));
                    if (sheet != null) AdxExcelAppEvents1_SheetBeforeDelete(null, sheet);
                    return;
                }
                case "SheetFollowHyperlink":
                {
                    Excel.Worksheet sheet = Sheet(Arg(args, 0), Arg(args, 5));
                    Excel.Hyperlink hyperlink = Hyperlink(Range(sheet, Arg(args, 1)));
                    if (sheet != null && hyperlink != null) adxExcelAppEvents1_SheetFollowHyperlink(null, sheet, hyperlink);
                    return;
                }
                case "WorkbookActivate":
                {
                    Excel.Workbook workbook = Workbook(Arg(args, 0));
                    if (workbook != null) WorkbookActivated(workbook);
                    return;
                }
            }
            LogUtility.LogWarn($"AddinModule: no handler for Excel event '{eventName}'.");
        }

        /// <summary>
        /// 11.1.2's adxExcelAppEvents1_WorkbookActivate without its first block, which re-navigates the task pane (Stage 2e); the
        /// ribbon-state block is 11.1.2's.
        /// </summary>
        private void WorkbookActivated(object hostObj)
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
                LogUtility.LogException(ex, "adxExcelAppEvents1_WorkbookActivate");
            }
        }

        private static string Arg(object[] args, int index) => index < args.Length ? args[index] as string : null;

        /// <summary>The named workbook, or the active one when no name was sent; null when Excel or the workbook is not there.</summary>
        private Excel.Workbook Workbook(string name)
        {
            try
            {
                Excel._Application excel = ExcelApp;
                if (excel == null) return null;
                return string.IsNullOrEmpty(name) ? excel.ActiveWorkbook : excel.Workbooks[name];
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: workbook '{name}' not found - {ex.Message}");
                return null;
            }
        }

        private Excel.Worksheet Sheet(string sheetName, string workbookName)
        {
            if (string.IsNullOrEmpty(sheetName)) return null;
            Excel.Workbook workbook = Workbook(workbookName);
            if (workbook == null) return null;
            try
            {
                return workbook.Worksheets[sheetName] as Excel.Worksheet;
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: sheet '{sheetName}' not found - {ex.Message}");
                return null;
            }
        }

        private static Excel.Range Range(Excel.Worksheet sheet, string address)
        {
            if (sheet == null || string.IsNullOrEmpty(address)) return null;
            try
            {
                return sheet.Range[address];
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: range '{address}' not found - {ex.Message}");
                return null;
            }
        }

        private static Excel.Hyperlink Hyperlink(Excel.Range cell)
        {
            try
            {
                return cell != null && cell.Hyperlinks.Count > 0 ? cell.Hyperlinks[1] : null;
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: hyperlink not readable - {ex.Message}");
                return null;
            }
        }
    }
}
