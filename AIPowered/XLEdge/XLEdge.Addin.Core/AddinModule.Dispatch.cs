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
        /// <summary>Ribbon action -> 11.1.2 handler (the ribbon-control argument is never used; pressed is the toggle state).</summary>
        internal static readonly Dictionary<string, Action<AddinModule, bool>> RibbonHandlers = new Dictionary<string, Action<AddinModule, bool>>(StringComparer.Ordinal)
        {
            ["RibEdgeRefresh"] = (m, pressed) => m.RibEdgeRefresh_OnClick(null, null, pressed),
            ["RibEdgeRefreshAll"] = (m, pressed) => m.RibEdgeRefreshAll_OnClick(null, null, pressed),
            ["RibEdgeOptions"] = (m, pressed) => m.RibEdgeOptions_OnClick(null, null, pressed),
            ["RibEdgeDialogBoxLauncher"] = (m, pressed) => m.RibEdgeDialogBoxLauncher_OnAction(null, null, pressed),
            ["RibEdgeHelp"] = (m, pressed) => m.RibEdgeHelp_OnClick(null, null, pressed),
            ["RibControlSheet"] = (m, pressed) => m.RibControlSheet_OnClick(null, null, pressed),
            ["RibEdgeIncludeOutputData"] = (m, pressed) => m.RibEdgeIncludeOutputData_OnClick(null, null, pressed),
            ["RibEdgeDebug"] = (m, pressed) => m.RibEdgeDebug_OnClick(null, null, pressed),
        };

        /// <summary>Ribbon actions whose handlers need the task pane or the login (Stage 2e); the AddinEntry spike handles Login/ShowHide.</summary>
        private static readonly HashSet<string> NotYetInTheCore = new HashSet<string>(StringComparer.Ordinal)
        {
            "RibEdgeLogout", "RibEdgeParamRefresh", "RibEdgeParamRefreshBook", "RibEdgeAbout", "RibEdgeLogout_PropertyChanging",
        };

        /// <summary>A forwarded ribbon click; false when no 11.1.2 handler takes it.</summary>
        internal bool OnRibbonAction(string action, object parameter)
        {
            if (action != null && RibbonHandlers.TryGetValue(action, out Action<AddinModule, bool> handler))
            {
                handler(this, parameter is bool pressed && pressed);
                return true;
            }
            if (NotYetInTheCore.Contains(action ?? string.Empty))
                LogUtility.LogInfo($"AddinModule: ribbon action '{action}' is not in the core yet (Stage 2e).");
            else
                LogUtility.LogWarn($"AddinModule: no handler for ribbon action '{action}'.");
            return false;
        }

        /// <summary>What a forwarded Excel event names: its sheet, workbook, cell or range address, and hyperlink shape.</summary>
        internal sealed class EventTarget
        {
            public string Sheet, Workbook, Address, Shape;
            public override string ToString() => $"sheet={Sheet ?? "null"} workbook={Workbook ?? "null"} address={Address ?? "null"} shape={Shape ?? "null"}";
        }

        /// <summary>
        /// The names in a forwarded event's arguments, as the shell sends them: SheetSelectionChange {sheet, address, cell count,
        /// workbook}, SheetActivate / SheetBeforeDelete {sheet, workbook}, SheetFollowHyperlink {sheet, cell, screen tip,
        /// sub-address, text, workbook, shape (a hyperlink on a shape has no cell)}, WorkbookActivate {workbook}. New arguments are
        /// only ever appended, so a shell that sends fewer still works (a missing workbook means the active one). Null for an
        /// event the core does not handle.
        /// </summary>
        internal static EventTarget Target(string eventName, object[] args)
        {
            args = args ?? new object[0];
            string Arg(int index) => index < args.Length ? args[index] as string : null;
            switch (eventName)
            {
                case "SheetSelectionChange": return new EventTarget { Sheet = Arg(0), Address = Arg(1), Workbook = Arg(3) };
                case "SheetActivate":
                case "SheetBeforeDelete": return new EventTarget { Sheet = Arg(0), Workbook = Arg(1) };
                case "SheetFollowHyperlink": return new EventTarget { Sheet = Arg(0), Address = Arg(1), Workbook = Arg(5), Shape = Arg(6) };
                case "WorkbookActivate": return new EventTarget { Workbook = Arg(0) };
                default: return null;
            }
        }

        /// <summary>A forwarded Excel event, given to 11.1.2's handler with the COM objects it received from Add-in Express.</summary>
        internal void OnExcelEvent(string eventName, object[] args)
        {
            EventTarget target = Target(eventName, args);
            if (target == null)
            {
                LogUtility.LogWarn($"AddinModule: no handler for Excel event '{eventName}'.");
                return;
            }
            switch (eventName)
            {
                case "SheetSelectionChange":
                {
                    // The handler's own first checks, before any lookup: this event fires on every click.
                    if (!XLEdgeAppState.Instance.IsLoginCompleted || XLEdgeAppState.Instance.ProcessRunning) return;
                    Excel.Worksheet sheet = Sheet(target.Sheet, target.Workbook);
                    Excel.Range range = Range(sheet, target.Address);
                    if (sheet != null && range != null) adxExcelAppEvents1_SheetSelectionChange(null, sheet, range);
                    return;
                }
                case "SheetActivate":
                {
                    Excel.Worksheet sheet = Sheet(target.Sheet, target.Workbook);
                    if (sheet != null) adxExcelAppEvents1_SheetActivate(null, sheet);
                    return;
                }
                case "SheetBeforeDelete":
                {
                    Excel.Worksheet sheet = Sheet(target.Sheet, target.Workbook);
                    if (sheet != null) AdxExcelAppEvents1_SheetBeforeDelete(null, sheet);
                    return;
                }
                case "SheetFollowHyperlink":
                {
                    Excel.Worksheet sheet = Sheet(target.Sheet, target.Workbook);
                    Excel.Hyperlink hyperlink = string.IsNullOrEmpty(target.Address) ? ShapeHyperlink(sheet, target.Shape) : Hyperlink(Range(sheet, target.Address));
                    if (sheet != null && hyperlink != null) adxExcelAppEvents1_SheetFollowHyperlink(null, sheet, hyperlink);
                    return;
                }
                case "WorkbookActivate":
                {
                    Excel.Workbook workbook = Workbook(target.Workbook);
                    if (workbook != null) WorkbookActivated(workbook);
                    return;
                }
            }
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

        // The lookups below release the collections they go through; the objects they return go to the handlers, which (like
        // 11.1.2's, given them by Add-in Express) may keep using them after an await.

        /// <summary>The named workbook, or the active one when no name was sent; null when Excel or the workbook is not there.</summary>
        private Excel.Workbook Workbook(string name)
        {
            Excel.Workbooks workbooks = null;
            try
            {
                Excel._Application excel = ExcelApp;
                if (excel == null) return null;
                if (string.IsNullOrEmpty(name)) return excel.ActiveWorkbook;
                workbooks = excel.Workbooks;
                return workbooks[name];
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: workbook '{name}' not found - {ex.Message}");
                return null;
            }
            finally
            {
                if (workbooks != null) Marshal.ReleaseComObject(workbooks);
            }
        }

        private Excel.Worksheet Sheet(string sheetName, string workbookName)
        {
            if (string.IsNullOrEmpty(sheetName)) return null;
            Excel.Workbook workbook = Workbook(workbookName);
            if (workbook == null) return null;
            Excel.Sheets sheets = null;
            try
            {
                sheets = workbook.Worksheets;
                return sheets[sheetName] as Excel.Worksheet;
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: sheet '{sheetName}' not found - {ex.Message}");
                return null;
            }
            finally
            {
                if (sheets != null) Marshal.ReleaseComObject(sheets);
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
            if (cell == null) return null;
            Excel.Hyperlinks hyperlinks = null;
            try
            {
                hyperlinks = cell.Hyperlinks;
                return hyperlinks.Count > 0 ? hyperlinks[1] : null;
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: hyperlink not readable - {ex.Message}");
                return null;
            }
            finally
            {
                if (hyperlinks != null) Marshal.ReleaseComObject(hyperlinks);
            }
        }

        /// <summary>The hyperlink of a shape (it has no cell); 11.1.2 handled those too (an attachment ScreenTip works on any link).</summary>
        private static Excel.Hyperlink ShapeHyperlink(Excel.Worksheet sheet, string shapeName)
        {
            if (sheet == null || string.IsNullOrEmpty(shapeName)) return null;
            Excel.Shapes shapes = null;
            Excel.Shape shape = null;
            try
            {
                shapes = sheet.Shapes;
                shape = shapes.Item(shapeName);
                return shape.Hyperlink;
            }
            catch (COMException ex)
            {
                LogUtility.LogDebug($"AddinModule: shape '{shapeName}' hyperlink not readable - {ex.Message}");
                return null;
            }
            finally
            {
                if (shape != null) Marshal.ReleaseComObject(shape);
                if (shapes != null) Marshal.ReleaseComObject(shapes);
            }
        }
    }
}
