using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Ported from AddinModule.vb's RibControlSheet_OnClick / CreateControlSheet / ProcessSheetLogic /
    /// ApplyDV / ConvertToArray / SheeNametExists (VB lines ~2698-3173). Builds the "Parameters
    /// Control Sheet" - a table listing every report's parameters (name, current value, operator)
    /// with an editable Value1/Value2/Operator column, so a user can change filter values and
    /// resubmit the report.
    /// </summary>
    public static class ParamsControlSheetBuilder
    {
        private const string ControlSheetName = "Parameters Control Sheet";
        private const string ControlTableName = "orb_params_control";

        // Store display values reference for extra parameters
        private static Dictionary<string, string> _extraParamDisplayValues = new Dictionary<string, string>();

        public static void ShowOrRebuild()
        {
            try
            {

                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                if (excelApp == null)
                {
                    LogUtility.LogDebug($"{nameof(ShowOrRebuild)}: Failed to get Excel application");
                    return;
                }

                Excel.Workbook workbook = excelApp.ActiveWorkbook;
                if (workbook == null)
                {
                    LogUtility.LogDebug($"{nameof(ShowOrRebuild)}: No active workbook");
                    return;
                }

                Excel.ListObject existing = FindControlTable(workbook);
                Excel.Worksheet controlSheet = null;

                if (existing != null)
                {
                    LogUtility.LogDebug($"{nameof(ShowOrRebuild)}: Found existing control table: {existing.Name}");

                    var result = MessageFunctions.XLEdgeMessage(
                        "Parameters Control Sheet already exists, this will wipe out the data and recreate a new one." + Environment.NewLine + "Do you want to continue?",
                        System.Windows.Forms.MessageBoxIcon.Question,
                        System.Windows.Forms.MessageBoxButtons.YesNoCancel);

                    if (result != System.Windows.MessageBoxResult.Yes)
                    {
                        LogUtility.LogDebug($"{nameof(ShowOrRebuild)}: User cancelled rebuild");
                        return;
                    }

                    controlSheet = existing.Parent as Excel.Worksheet;
                }
                else if (ExcelSheetHelper.SheetExists(ControlSheetName, workbook))
                {
                    controlSheet = (Excel.Worksheet)workbook.Worksheets[ControlSheetName];
                }

                CreateControlSheet(workbook, ref controlSheet);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(ShowOrRebuild)}: Unexpected error");
            }
        }

        internal static Excel.ListObject FindControlTable(Excel.Workbook workbook)
        {
            try
            {

                foreach (Excel.Worksheet ws in workbook.Worksheets)
                {
                    bool release = true;
                    try
                    {
                        if (ws == null)
                        {
                            continue;
                        }

                        foreach (Excel.ListObject lo in ws.ListObjects)
                        {
                            try
                            {
                                if (lo == null)
                                {
                                    continue;
                                }

                                if (string.Equals(lo.Name, ControlTableName, StringComparison.OrdinalIgnoreCase))
                                {
                                    LogUtility.LogDebug($"{nameof(FindControlTable)}: Found control table: {lo.Name} in worksheet: {ws.Name}");
                                    release = false;
                                    return lo;
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(FindControlTable)}: Error iterating ListObjects - {ex.Message}");
                            }
                        }
                    }
                    finally
                    {
                        if (release && ws != null)
                        {
                            Marshal.ReleaseComObject(ws);
                        }
                    }
                }

                LogUtility.LogDebug($"{nameof(FindControlTable)}: Control table not found");
                return null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(FindControlTable)}: Unexpected error");
                return null;
            }
        }

        private static void CreateControlSheet(Excel.Workbook workbook, ref Excel.Worksheet ctrlSheet)
        {
            try
            {

                var data = new List<object[]>();
                _extraParamDisplayValues.Clear();

                if (!CollectParameterData(workbook, data))
                {
                    return;
                }

                if (!ResolveOrCreateControlSheet(workbook, ref ctrlSheet))
                {
                    return;
                }

                Excel.Range operatorsRange = ClearControlSheetRangeAndWriteOperatorKeys(workbook, ctrlSheet);

                int lastRow = WriteParameterRows(ctrlSheet, data);

                WriteControlSheetHeaders(ctrlSheet);

                CreateAndConfigureParameterTable(ctrlSheet, lastRow);

                // Populate extra-parameter display values (column IA) before the lock pass below.
                try
                {
                    AddExtraParameterDisplayValues(ctrlSheet, lastRow);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to add extra parameter display values");
                }

                // Operator validation on Column I (9)
                try
                {
                    if (operatorsRange != null)
                    {
                        string formula1 = operatorsRange.Address;
                        Excel.Range validationRange = ctrlSheet.Range[ctrlSheet.Cells[4, 9], ctrlSheet.Cells[lastRow, 9]];
                        if (validationRange != null)
                        {
                            validationRange.Validation.Delete();
                            validationRange.Validation.Add(
                                Excel.XlDVType.xlValidateList,
                                Excel.XlDVAlertStyle.xlValidAlertStop,
                                Excel.XlFormatConditionOperator.xlBetween,
                                "=" + formula1,
                                Type.Missing);
                            validationRange.Validation.IgnoreBlank = false;
                            validationRange.Validation.InCellDropdown = true;
                            validationRange.Validation.ErrorTitle = "Parameter Operators";
                            validationRange.Validation.InputMessage = "Select operator.";
                            validationRange.Validation.ErrorMessage = "Operator should be from the given list.";
                            validationRange.Validation.ShowInput = true;
                            validationRange.Validation.ShowError = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to add operator validation");
                }

                // Lock non-editable columns to their current values
                try
                {
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[3, 1], ctrlSheet.Cells[3, 11]]); // Header row
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 1], ctrlSheet.Cells[lastRow, 1]]); // Column A - Report Name
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 2], ctrlSheet.Cells[lastRow, 2]]); // Column B - Report ID
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 4], ctrlSheet.Cells[lastRow, 4]]); // Column D - Parameter Type
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 5], ctrlSheet.Cells[lastRow, 5]]); // Column E - Parameter Name
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 6], ctrlSheet.Cells[lastRow, 6]]); // Column F - Parameter Display Name
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 7], ctrlSheet.Cells[lastRow, 7]]); // Column G - Is Required
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 8], ctrlSheet.Cells[lastRow, 8]]); // Column H - Data Type

                    // Lock column IA (235), which holds the GL Accounts display values
                    LockToCurrentValue(ctrlSheet.Range[ctrlSheet.Cells[4, 235], ctrlSheet.Cells[lastRow, 235]]);

                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to lock columns");
                }

                // Highlight required parameters in Column E (name) and F (display name)
                try
                {
                    Excel.Range requiredColumn = ctrlSheet.Range[ctrlSheet.Cells[4, 7], ctrlSheet.Cells[lastRow, 7]];
                    if (requiredColumn != null)
                    {
                        foreach (Excel.Range cell in requiredColumn.Cells)
                        {
                            try
                            {
                                bool isRequired = cell.Value != null && string.Equals(cell.Value.ToString().Trim(), "yes", StringComparison.OrdinalIgnoreCase);
                                if (isRequired)
                                {
                                    Excel.Range nameCell = ctrlSheet.Range["E" + cell.Row];
                                    if (nameCell != null) nameCell.Font.Color = ColorTranslator.ToOle(Color.Red);

                                    Excel.Range displayCell = ctrlSheet.Range["F" + cell.Row];
                                    if (displayCell != null) displayCell.Font.Color = ColorTranslator.ToOle(Color.Red);
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(CreateControlSheet)}: Failed to highlight cell row {cell.Row} - {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(CreateControlSheet)}: failed to highlight required-parameter cells - {ex.Message}");
                }

                // Setup validation for ORACLE_GL_SEGMENT_VALUES to prevent direct editing
                try
                {
                    SetupGLSegmentValidation(ctrlSheet, lastRow);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to setup GL segment validation");
                }

                // Setup Value2 (Column K) validation - only editable when operator contains "Between"
                try
                {
                    SetupValue2Validation(ctrlSheet, lastRow);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to setup Value2 validation");
                }

                // Setup the title and other UI elements
                try
                {

                    Excel.Range mergeRange = ctrlSheet.Range["A1:J1"];
                    if (mergeRange != null) mergeRange.UnMerge();

                    Excel.Range clearRange = ctrlSheet.Range["A1:K1"];
                    if (clearRange != null) clearRange.Clear();

                    Excel.Range titleRange = ctrlSheet.Range[ctrlSheet.Cells[1, 1], ctrlSheet.Cells[1, 11]];
                    if (titleRange != null) titleRange.Merge();

                    var title1 = ctrlSheet.Range["A1"];
                    if (title1 != null)
                    {
                        title1.Value2 = "Parameters Control Sheet";
                        title1.Font.Bold = true;
                        title1.Font.Italic = true;
                        title1.Font.Size = 11;
                        title1.Font.ColorIndex = 2;
                        title1.Interior.Color = Rgb(21, 96, 130);
                    }

                    var genLabel = ctrlSheet.Range["M1"];
                    if (genLabel != null)
                    {
                        genLabel.Value2 = "Generated On : ";
                        genLabel.Font.Bold = true;
                        genLabel.Font.Italic = true;
                        genLabel.Font.Size = 11;
                        genLabel.Font.ColorIndex = 2;
                        genLabel.Interior.Color = Rgb(241, 169, 131);
                    }

                    var genValue = ctrlSheet.Range["N1"];
                    if (genValue != null)
                    {
                        genValue.NumberFormat = "dd-mmm-yyyy hh:mm:ss";
                        genValue.Value2 = DateTime.Now;
                        genValue.Font.Bold = false;
                        genValue.Font.Italic = true;
                        genValue.Font.Size = 10;
                        genValue.Font.Color = Rgb(21, 96, 130);
                        genValue.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    }

                    var tzLabel = ctrlSheet.Range["M2"];
                    if (tzLabel != null)
                    {
                        tzLabel.Value2 = "Time Zone  : ";
                        tzLabel.Font.Bold = true;
                        tzLabel.Font.Italic = true;
                        tzLabel.Font.Size = 11;
                        tzLabel.Font.ColorIndex = 2;
                        tzLabel.Interior.Color = Rgb(241, 169, 131);
                    }

                    var tzValue = ctrlSheet.Range["N2"];
                    if (tzValue != null)
                    {
                        tzValue.Value2 = TimeZoneInfo.Local.DisplayName;
                        tzValue.Font.Bold = false;
                        tzValue.Font.Italic = true;
                        tzValue.Font.Size = 10;
                        tzValue.Font.Color = Rgb(21, 96, 130);
                        tzValue.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    }

                    try
                    {
                        Excel.Range autofitRange = ctrlSheet.Range["M1:N2"];
                        if (autofitRange != null)
                        {
                            autofitRange.Columns.EntireColumn.AutoFit();
                        }

                        ctrlSheet.Tab.Color = Rgb(0, 176, 240);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(CreateControlSheet)}: failed to autofit generated-on columns / set tab color - {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to setup UI elements");
                }

                try
                {
                    Excel.Range gotoRange = ctrlSheet.Range["A1"];
                    if (gotoRange != null)
                    {
                        workbook.Application.Goto(gotoRange, true);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(CreateControlSheet)}: Failed to goto A1");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(CreateControlSheet));
            }
        }

        /// <summary>
        /// Walks every "_E" report table in the workbook and collects its parameters into
        /// <paramref name="data"/> via <see cref="ProcessSheetParams"/>. Returns false only if
        /// the workbook is null or an unexpected error stops collection entirely.
        /// </summary>
        private static bool CollectParameterData(Excel.Workbook workbook, List<object[]> data)
        {
            try
            {
                if (workbook == null)
                {
                    LogUtility.LogDebug($"{nameof(CollectParameterData)}: workbook is null");
                    return false;
                }

                // Workbook.Worksheets actually returns an Excel.Sheets COM object at runtime, not
                // Excel.Worksheets - casting/declaring it as Excel.Worksheets throws
                // InvalidCastException (E_NOINTERFACE) at runtime even though it compiles fine.
                Excel.Sheets allSheets = workbook.Worksheets;
                try
                {
                    foreach (Excel.Worksheet ws in allSheets)
                    {
                        try
                        {
                            if (ws.ListObjects.Count == 0)
                            {
                                continue;
                            }

                            Excel.ListObject tableObj = ws.ListObjects[1];
                            if (!tableObj.Name.EndsWith("_E", StringComparison.Ordinal))
                            {
                                continue;
                            }

                            string[] strList = tableObj.Name.Split('_');
                            string reportId = strList.Length > 1 ? strList[1] : string.Empty;

                            if (!ReportGenerator.TryGetStoredReportXml(workbook, tableObj.Name, out string title, out _, out string paramsJson))
                            {
                                continue;
                            }

                            if (string.IsNullOrWhiteSpace(paramsJson))
                            {
                                continue;
                            }

                            string reportName = title != null && title.Split('|').Length >= 4 ? title.Split('|')[3] : string.Empty;
                            if (string.IsNullOrWhiteSpace(reportName))
                            {
                                continue;
                            }

                            ProcessSheetParams(ws, data, paramsJson, reportId, reportName);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(CollectParameterData)}: Failed processing worksheet {ws?.Name} - {ex.Message}");
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(ws);
                        }
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(allSheets);
                }

                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(CollectParameterData));
                return false;
            }
        }

        /// <summary>
        /// Reuses the caller-supplied control sheet if one was resolved already, otherwise creates
        /// a new worksheet and names it. Returns false only if sheet creation itself fails.
        /// </summary>
        private static bool ResolveOrCreateControlSheet(Excel.Workbook workbook, ref Excel.Worksheet ctrlSheet)
        {
            try
            {
                if (ctrlSheet == null)
                {
                    try
                    {
                        ctrlSheet = (Excel.Worksheet)workbook.Worksheets.Add(workbook.Sheets[1], Type.Missing, Type.Missing, Type.Missing);
                    }
                    catch
                    {
                        ctrlSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count], Type.Missing, Type.Missing);
                    }

                    if (!ExcelSheetHelper.SheetExists(ControlSheetName, workbook))
                    {
                        ctrlSheet.Name = ControlSheetName;
                    }
                }

                return ctrlSheet != null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ResolveOrCreateControlSheet));
                return false;
            }
        }

        /// <summary>
        /// Clears the sheet's working range and (re)writes the operator dropdown's source list
        /// into column IT, returning that range so the caller can build a validation formula from it.
        /// </summary>
        private static Excel.Range ClearControlSheetRangeAndWriteOperatorKeys(Excel.Workbook workbook, Excel.Worksheet ctrlSheet)
        {
            ctrlSheet.Range["A1:M10000"].Cells.Clear();

            string[] operatorKeys = XLEdgeOperatorMappings.Map.Keys.ToArray();
            Excel.Range operatorsRange = ctrlSheet.Range[$"IT1:IT{operatorKeys.Length}"];
            operatorsRange.Value2 = workbook.Application.WorksheetFunction.Transpose(operatorKeys);

            return operatorsRange;
        }

        /// <summary>
        /// Bulk-writes every collected parameter row starting at row 4 (columns A-K: Report Name,
        /// Report ID, Worksheet, Parameter Type, Parameter Name, Parameter Display Name, Is Required,
        /// Data Type, Operator, Value1, Value2), forcing Report ID/Value1/Value2 to text format so
        /// Excel doesn't silently reinterpret them as numbers. Returns the last row written to (or
        /// row 4 if there was no data to write).
        /// </summary>
        private static int WriteParameterRows(Excel.Worksheet ctrlSheet, List<object[]> data)
        {
            const int columnCount = 11;
            int lastRow;

            if (data.Count > 0)
            {
                object[,] outputArray = new object[data.Count, columnCount];
                for (int r = 0; r < data.Count; r++)
                {
                    for (int c = 0; c < columnCount; c++)
                    {
                        outputArray[r, c] = data[r][c];
                    }
                }

                Excel.Range outputRange = ctrlSheet.Range[ctrlSheet.Cells[4, 1], ctrlSheet.Cells[3 + data.Count, columnCount]];
                ctrlSheet.Range[ctrlSheet.Cells[4, 2], ctrlSheet.Cells[3 + data.Count, 2]].NumberFormat = "@";
                ctrlSheet.Range[ctrlSheet.Cells[4, 10], ctrlSheet.Cells[3 + data.Count, 11]].NumberFormat = "@";
                outputRange.Value2 = outputArray;
                lastRow = 3 + data.Count;
            }
            else
            {
                lastRow = 4;
            }

            return lastRow;
        }

        /// <summary>Writes the row-3 column headers matching <see cref="WriteParameterRows"/>'s 11-column layout.</summary>
        private static void WriteControlSheetHeaders(Excel.Worksheet ctrlSheet)
        {
            string[] headers =
            {
                "Report Name", "Report ID", "Worksheet", "Parameter Type", "Parameter Name",
                "Parameter Display Name", "Is Required", "Data Type", "Operator", "Value1", "Value2"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                ctrlSheet.Cells[3, i + 1] = headers[i];
            }
        }

        /// <summary>Turns the header/data range into the "orb_params_control" Excel table and applies its styling.</summary>
        private static void CreateAndConfigureParameterTable(Excel.Worksheet ctrlSheet, int lastRow)
        {
            Excel.Range tableRange = ctrlSheet.Range[ctrlSheet.Cells[3, 1], ctrlSheet.Cells[lastRow, 11]];
            Excel.ListObject tableObj2 = null;
            try
            {
                tableObj2 = ctrlSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(CreateAndConfigureParameterTable)}: Failed to create parameters control table");
            }

            if (tableObj2 != null)
            {
                tableObj2.Name = ControlTableName;
                tableObj2.TableStyle = "TableStyleLight9";
                tableObj2.HeaderRowRange.Font.Size = 10;
                if (tableObj2.DataBodyRange != null)
                {
                    tableObj2.DataBodyRange.Font.Size = 9;
                    tableObj2.DataBodyRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                }

                foreach (Excel.ListColumn col in tableObj2.ListColumns)
                {
                    try
                    {
                        col.Range.EntireColumn.AutoFit();
                        if ((int)col.Range.ColumnWidth > 35)
                        {
                            col.Range.ColumnWidth = 35;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(CreateAndConfigureParameterTable)}: Failed to autofit/resize column {col.Name} - {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Safe COM cast helper to trace InvalidCastException errors
        /// </summary>
        private static T SafeCast<T>(object obj, string context) where T : class
        {
            try
            {
                if (obj == null)
                {
                    LogUtility.LogDebug($"{context}: Object is null");
                    return null;
                }

                if (obj is T casted)
                {
                    return casted;
                }

                LogUtility.LogDebug($"{context}: Object is not of type {typeof(T).Name}. Actual type: {obj.GetType().Name}");
                return null;
            }
            catch (InvalidCastException ex)
            {
                LogUtility.LogException(ex, $"{context}: Failed to cast object to {typeof(T).Name}. Object type: {obj?.GetType().Name ?? "null"}");
                return null;
            }
        }

        /// <summary>
        /// Sets up validation for ORACLE_GL_SEGMENT_VALUES cells to prevent direct editing
        /// </summary>
        private static void SetupGLSegmentValidation(Excel.Worksheet ctrlSheet, int lastRow)
        {
            try
            {

                if (ctrlSheet == null)
                {
                    LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: ctrlSheet is null");
                    return;
                }

                // Column J (10) is Value1 - only editable for ORACLE_GL_SEGMENT_VALUES
                Excel.Range value1Range = ctrlSheet.Range[ctrlSheet.Cells[4, 10], ctrlSheet.Cells[lastRow, 10]];
                if (value1Range == null)
                {
                    LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: value1Range is null");
                    return;
                }

                int processedCount = 0;
                int totalCells = 0;

                try
                {
                    foreach (Excel.Range cell in value1Range.Cells)
                    {
                        totalCells++;
                        try
                        {
                            if (cell == null)
                            {
                                continue;
                            }

                            int row = cell.Row;

                            // Check if this row is for ORACLE_GL_SEGMENT_VALUES
                            Excel.Range paramNameCell = SafeCast<Excel.Range>(ctrlSheet.Cells[row, 5], $"{nameof(SetupGLSegmentValidation)}: paramNameCell row {row}");
                            Excel.Range paramTypeCell = SafeCast<Excel.Range>(ctrlSheet.Cells[row, 4], $"{nameof(SetupGLSegmentValidation)}: paramTypeCell row {row}");

                            if (paramNameCell == null || paramTypeCell == null)
                            {
                                continue;
                            }

                            string paramName = paramNameCell.Value2 as string ?? string.Empty;
                            string paramType = paramTypeCell.Value2 as string ?? string.Empty;

                            // If this is an extraParameter with name ORACLE_GL_SEGMENT_VALUES
                            if (paramType.Equals("extraParameters", StringComparison.OrdinalIgnoreCase) &&
                                paramName.Equals("ORACLE_GL_SEGMENT_VALUES", StringComparison.OrdinalIgnoreCase))
                            {
                                LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: Found ORACLE_GL_SEGMENT_VALUES at row {row}");

                                // Clear any existing validation
                                try
                                {
                                    cell.Validation.Delete();
                                }
                                catch (Exception ex)
                                {
                                    LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: Failed to delete existing validation for row {row} - {ex.Message}");
                                }

                                // Add validation that prevents editing with a custom message
                                cell.Validation.Add(
                                    Excel.XlDVType.xlValidateCustom,
                                    Excel.XlDVAlertStyle.xlValidAlertStop,
                                    Excel.XlFormatConditionOperator.xlBetween,
                                    "=FALSE",
                                    Type.Missing);

                                cell.Validation.ErrorTitle = "GL Accounts";
                                cell.Validation.ErrorMessage =
                                    "ORACLE_GL_SEGMENT_VALUES cannot be edited directly.\n" +
                                    "Please DOUBLE-CLICK this cell to open the GL Accounts window.";
                                cell.Validation.IgnoreBlank = true;
                                cell.Validation.ShowError = true;

                                // Lock the cell as an extra safeguard
                                cell.Locked = true;

                                // Set a background color to indicate it's special
                                cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);

                                processedCount++;
                            }
                        }
                        catch (InvalidCastException ex)
                        {
                            LogUtility.LogException(ex, $"{nameof(SetupGLSegmentValidation)}: InvalidCastException at cell index {totalCells}");
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: Failed to set validation for cell at index {totalCells} - {ex.Message}");
                        }
                    }
                }
                finally
                {
                    if (value1Range != null)
                    {
                        Marshal.ReleaseComObject(value1Range);
                    }
                }

                LogUtility.LogDebug($"{nameof(SetupGLSegmentValidation)}: Completed. Processed {processedCount} of {totalCells} cells.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(SetupGLSegmentValidation));
            }
        }

        // Helper method to setup Value2 (Column K) validation
        private static void SetupValue2Validation(Excel.Worksheet ctrlSheet, int lastRow)
        {
            try
            {

                if (ctrlSheet == null)
                {
                    LogUtility.LogDebug($"{nameof(SetupValue2Validation)}: ctrlSheet is null");
                    return;
                }

                // Column K (11) is Value2 - only editable when Column I (9) contains "Between"
                Excel.Range value2Range = ctrlSheet.Range[ctrlSheet.Cells[4, 11], ctrlSheet.Cells[lastRow, 11]];
                if (value2Range == null)
                {
                    LogUtility.LogDebug($"{nameof(SetupValue2Validation)}: value2Range is null");
                    return;
                }

                int processedCount = 0;
                int totalCells = 0;

                try
                {
                    foreach (Excel.Range cell in value2Range.Cells)
                    {
                        totalCells++;
                        try
                        {
                            if (cell == null)
                            {
                                continue;
                            }

                            int row = cell.Row;
                            string cellFormula = $"=ISNUMBER(SEARCH(\"Between\", I{row}))";

                            // Clear existing validation
                            try
                            {
                                cell.Validation.Delete();
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(SetupValue2Validation)}: Failed to delete existing validation for row {row} - {ex.Message}");
                            }

                            // Add custom validation
                            cell.Validation.Add(
                                Excel.XlDVType.xlValidateCustom,
                                Excel.XlDVAlertStyle.xlValidAlertStop,
                                Excel.XlFormatConditionOperator.xlBetween,
                                cellFormula,
                                Type.Missing);

                            cell.Validation.ErrorTitle = "Value2 Validation";
                            cell.Validation.ErrorMessage = "Value2 can only be edited when Operator contains 'Between'.";
                            cell.Validation.IgnoreBlank = true;
                            cell.Validation.ShowError = true;

                            processedCount++;
                        }
                        catch (InvalidCastException ex)
                        {
                            LogUtility.LogException(ex, $"{nameof(SetupValue2Validation)}: InvalidCastException at cell index {totalCells}");
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(SetupValue2Validation)}: Failed to set validation for cell at index {totalCells} - {ex.Message}");
                        }
                    }
                }
                finally
                {
                    if (value2Range != null)
                    {
                        Marshal.ReleaseComObject(value2Range);
                    }
                }

                LogUtility.LogDebug($"{nameof(SetupValue2Validation)}: Completed. Processed {processedCount} of {totalCells} cells.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(SetupValue2Validation));
            }
        }

        public static string GetGLAccountDisplayValues(Excel.Worksheet ctrlSheet, int rowNumber)
        {
            try
            {

                if (ctrlSheet == null)
                {
                    LogUtility.LogDebug($"{nameof(GetGLAccountDisplayValues)}: ctrlSheet is null");
                    return string.Empty;
                }

                Excel.Range displayValueCell = SafeCast<Excel.Range>(ctrlSheet.Cells[rowNumber, 235], $"{nameof(GetGLAccountDisplayValues)}: displayValueCell");
                if (displayValueCell != null && displayValueCell.Value2 != null)
                {
                    string result = displayValueCell.Value2.ToString();
                    LogUtility.LogDebug($"{nameof(GetGLAccountDisplayValues)}: Found display values for row {rowNumber}: {result.Substring(0, Math.Min(100, result.Length))}...");
                    return result;
                }

                LogUtility.LogDebug($"{nameof(GetGLAccountDisplayValues)}: No display values found for row {rowNumber}");
                return string.Empty;
            }
            catch (InvalidCastException ex)
            {
                LogUtility.LogException(ex, $"{nameof(GetGLAccountDisplayValues)}: InvalidCastException for row {rowNumber}");
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(GetGLAccountDisplayValues)}: Failed to get display values for row {rowNumber}");
                return string.Empty;
            }
        }

        /// <summary>Locks a range to its current value via a self-referencing data-validation rule</summary>
        private static void LockToCurrentValue(Excel.Range range)
        {
            const int MaxValidationFormulaLength = 255;

            try
            {
                if (range == null)
                {
                    LogUtility.LogDebug($"{nameof(LockToCurrentValue)}: Range is null");
                    return;
                }

                LogUtility.LogDebug($"{nameof(LockToCurrentValue)}: Processing range {range.Address}");

                int processedCount = 0;
                int totalCells = 0;

                foreach (Excel.Range cell in range.Cells)
                {
                    totalCells++;
                    try
                    {
                        if (cell == null)
                        {
                            continue;
                        }

                        string currentValue = cell.Text as string ?? string.Empty;
                        string escapedValue = currentValue.Replace("\"", "\"\"");
                        string cellAddress = cell.Address[false, false];
                        string formula = "=" + cellAddress + "=\"" + escapedValue + "\"";

                        if (formula.Length > MaxValidationFormulaLength)
                        {
                            LogUtility.LogDebug($"{nameof(LockToCurrentValue)}: skipping lock for {cellAddress} - built formula is {formula.Length} chars, over Excel's {MaxValidationFormulaLength}-char Validation.Add limit (value: {currentValue}).");
                            Marshal.ReleaseComObject(cell);
                            continue;
                        }

                        try
                        {
                            cell.Validation.Delete();
                            cell.Validation.Add(
                                Excel.XlDVType.xlValidateCustom,
                                Excel.XlDVAlertStyle.xlValidAlertStop,
                                Type.Missing,
                                formula,
                                Type.Missing);
                            cell.Validation.ErrorTitle = "Parameter Control Sheet";
                            cell.Validation.ErrorMessage = "Editing this cell is not allowed.";
                            cell.Validation.IgnoreBlank = false;
                            cell.Validation.ShowError = true;
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, $"Failed applying data validation to cell {cellAddress} (value: {currentValue})");
                        }
                        finally
                        {
                            if (cell != null)
                            {
                                Marshal.ReleaseComObject(cell);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(LockToCurrentValue)}: Failed to process cell at index {totalCells} - {ex.Message}");
                    }
                }

                LogUtility.LogDebug($"{nameof(LockToCurrentValue)}: Completed. Processed {processedCount} of {totalCells} cells in range {range.Address}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(LockToCurrentValue)}: Failed for range {range?.Address ?? "null"}");
            }
        }

        private static void ProcessSheetParams(Excel.Worksheet sheet, List<object[]> data, string paramsJson, string reportId, string reportName)
        {
            try
            {

                JsonDocument doc;
                try
                {
                    doc = JsonDocument.Parse(paramsJson);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"{nameof(ProcessSheetParams)}: Failed to parse JSON for {reportName}");
                    return;
                }

                using (doc)
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Array)
                    {
                        LogUtility.LogDebug($"{nameof(ProcessSheetParams)}: Root element is not an array for {reportName}");
                        return;
                    }

                    int itemCount = 0;
                    int processedCount = 0;

                    foreach (JsonElement item in doc.RootElement.EnumerateArray())
                    {
                        itemCount++;
                        try
                        {
                            // Check if this is an extraParameters entry
                            if (JsonHelper.TryGetProperty(item, "extraParameters", out var extraParamsEl) &&
                                extraParamsEl.ValueKind == JsonValueKind.Object)
                            {
                                LogUtility.LogDebug($"{nameof(ProcessSheetParams)}: Found extraParameters at index {itemCount}");
                                ProcessExtraParameters(data, extraParamsEl, reportId, reportName, sheet.Name);
                                continue;
                            }

                            // Regular parameter processing
                            string displayName = GetString(item, "name") ?? GetString(item, "label");
                            string displayLabel = GetString(item, "label") ?? GetString(item, "name");
                            bool required = JsonHelper.TryGetProperty(item, "required", out var reqEl) && reqEl.ValueKind == JsonValueKind.True;
                            string type = GetString(item, "type");
                            string paramOperator = GetString(item, "operator");

                            if (string.IsNullOrEmpty(displayName) || string.IsNullOrEmpty(type) || string.IsNullOrEmpty(paramOperator))
                            {
                                LogUtility.LogDebug($"{nameof(ProcessSheetParams)}: Skipping item {itemCount} - missing required fields");
                                continue;
                            }

                            string operatorKey = XLEdgeOperatorMappings.Map
                                .FirstOrDefault(kvp => kvp.Value == paramOperator).Key ?? string.Empty;

                            string componentType = GetString(item, "componentType");
                            if (paramOperator == "IN" && (componentType == "single-selection-prompt" || componentType == "oracle-erp-resp-selection"))
                            {
                                operatorKey = "is equal to";
                            }
                            else if (paramOperator == "NOT IN" && (componentType == "single-selection-prompt" || componentType == "oracle-erp-resp-selection"))
                            {
                                operatorKey = "does not equal";
                            }

                            (string value1, string value2) = ExtractValues(item, paramOperator, type);
                            string requiredText = required ? "Yes" : "No";

                            data.Add(new object[]
                            {
                                XLEdgeValueFormatter.RemoveEquaSymbol(reportName),
                                reportId,
                                XLEdgeValueFormatter.RemoveEquaSymbol(sheet.Name),
                                "parameter",
                                XLEdgeValueFormatter.RemoveEquaSymbol(displayName),
                                XLEdgeValueFormatter.RemoveEquaSymbol(displayLabel),
                                requiredText,
                                type,
                                operatorKey,
                                XLEdgeValueFormatter.RemoveEquaSymbol(value1),
                                XLEdgeValueFormatter.RemoveEquaSymbol(value2)
                            });
                            processedCount++;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, $"{nameof(ProcessSheetParams)}: Failed to process item {itemCount}");
                        }
                    }

                    LogUtility.LogDebug($"{nameof(ProcessSheetParams)}: Processed {processedCount} of {itemCount} items for {reportName}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(ProcessSheetParams)}: Unexpected error for {reportName}");
            }
        }

        private static void ProcessExtraParameters(List<object[]> data, JsonElement extraParamsEl, string reportId, string reportName, string sheetName)
        {
            try
            {

                var extraParamMappings = new Dictionary<string, string>
                {
                    { "ORACLE_RESP_ID", "Responsibility" },
                    { "ORACLE_GL_SEGMENT_VALUES", "GL Accounts" }
                };

                foreach (var mapping in extraParamMappings)
                {
                    string paramName = mapping.Key;
                    string displayLabel = mapping.Value;

                    if (JsonHelper.TryGetProperty(extraParamsEl, paramName, out var valueEl) &&
                        valueEl.ValueKind != JsonValueKind.Null && valueEl.ValueKind != JsonValueKind.Undefined)
                    {
                        string value = valueEl.ToString();
                        string displayValue = string.Empty;

                        // Check for display values
                        string displayKey = paramName.Replace("_VALUES", "_DISPLAY_VALUE");
                        if (JsonHelper.TryGetProperty(extraParamsEl, displayKey, out var displayEl))
                        {
                            displayValue = displayEl.ToString();
                            if (paramName == "ORACLE_RESP_ID" && JsonHelper.TryGetProperty(extraParamsEl, "ORACLE_RESP_DISPLAY_VALUE", out var respDisplayEl))
                            {
                                _extraParamDisplayValues[paramName] = respDisplayEl.ToString();
                            }
                        }
                        else if (paramName == "ORACLE_GL_SEGMENT_VALUES")
                        {
                            // Store the nested display values for later use in cell IA
                            if (JsonHelper.TryGetProperty(extraParamsEl, "ORACLE_GL_SEGMENT_DISPLAY_VALUES", out var segmentDisplayEl))
                            {
                                // Store as string in the dictionary
                                string segmentDisplayStr = segmentDisplayEl.ToString();
                                _extraParamDisplayValues[paramName] = segmentDisplayStr;
                                displayValue = segmentDisplayStr;
                            }
                        }
                        else if (paramName == "ORACLE_RESP_ID")
                        {
                            if (JsonHelper.TryGetProperty(extraParamsEl, "ORACLE_RESP_DISPLAY_VALUE", out var respDisplayEl))
                            {
                                displayValue = respDisplayEl.ToString();
                                _extraParamDisplayValues[paramName] = displayValue;
                            }
                        }

                        // Determine type based on value
                        string type = "INTEGER";
                        if (value.Contains("\"") || value.Contains("{"))
                        {
                            type = "STRING";
                        }

                        // Add as extraParameter - set Is Required as "Yes"
                        data.Add(new object[]
                        {
                            XLEdgeValueFormatter.RemoveEquaSymbol(reportName),
                            reportId,
                            XLEdgeValueFormatter.RemoveEquaSymbol(sheetName),
                            "extraParameters",
                            XLEdgeValueFormatter.RemoveEquaSymbol(paramName),
                            XLEdgeValueFormatter.RemoveEquaSymbol(displayLabel),
                            "Yes",
                            type,
                            "is equal to",
                            XLEdgeValueFormatter.RemoveEquaSymbol(value),
                            string.Empty
                        });

                        LogUtility.LogDebug($"{nameof(ProcessExtraParameters)}: Added extra parameter: {paramName} = {value} with display: {displayValue}");
                    }
                    else
                    {
                        LogUtility.LogDebug($"{nameof(ProcessExtraParameters)}: {paramName} not found in extra parameters");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ProcessExtraParameters));
            }
        }

        /// <summary>
        /// Adds extra parameter display values to column IA and preserves them for GL Accounts
        /// </summary>
        private static void AddExtraParameterDisplayValues(Excel.Worksheet ctrlSheet, int lastRow)
        {
            try
            {

                if (_extraParamDisplayValues.Count == 0)
                {
                    LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: No extra parameter display values to store");
                    return;
                }

                if (ctrlSheet == null)
                {
                    LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: ctrlSheet is null");
                    return;
                }

                Excel.Range paramTypeCol = ctrlSheet.Range[ctrlSheet.Cells[4, 4], ctrlSheet.Cells[lastRow, 4]];
                Excel.Range paramNameCol = ctrlSheet.Range[ctrlSheet.Cells[4, 5], ctrlSheet.Cells[lastRow, 5]];
                Excel.Range value1Col = ctrlSheet.Range[ctrlSheet.Cells[4, 10], ctrlSheet.Cells[lastRow, 10]];

                if (paramTypeCol == null || paramNameCol == null || value1Col == null)
                {
                    LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: Required columns are null");
                    return;
                }

                int storedCount = 0;
                int totalRows = 0;

                try
                {
                    for (int i = 0; i < paramTypeCol.Rows.Count; i++)
                    {
                        totalRows++;
                        try
                        {
                            Excel.Range typeCell = paramTypeCol.Cells[i + 1, 1] as Excel.Range;
                            Excel.Range nameCell = paramNameCol.Cells[i + 1, 1] as Excel.Range;
                            Excel.Range valueCell = value1Col.Cells[i + 1, 1] as Excel.Range;

                            if (typeCell == null || nameCell == null)
                            {
                                continue;
                            }

                            string paramType = typeCell.Value2 as string ?? string.Empty;
                            string paramName = nameCell.Value2 as string ?? string.Empty;

                            if (paramType.Equals("extraParameters", StringComparison.OrdinalIgnoreCase) &&
                                _extraParamDisplayValues.TryGetValue(paramName, out string displayValuesStr))
                            {
                                int rowNumber = 4 + i;

                                // Store in column IA (column 235)
                                Excel.Range displayValueCell = ctrlSheet.Cells[rowNumber, 235] as Excel.Range;
                                if (displayValueCell != null)
                                {
                                    // For GL segment values, preserve the existing display label if it exists
                                    if (paramName == "ORACLE_GL_SEGMENT_VALUES" && valueCell != null && valueCell.Value2 != null)
                                    {
                                        // Check if we already have a display value in IA
                                        object existingDisplay = displayValueCell.Value2;
                                        if (existingDisplay != null && !string.IsNullOrWhiteSpace(existingDisplay.ToString()))
                                        {
                                            // Keep the existing display label - the user might have edited it
                                            LogUtility.LogDebug($"Preserved existing GL display label: {existingDisplay}");
                                        }
                                        else
                                        {
                                            // Store the new display values
                                            displayValueCell.Value2 = displayValuesStr;
                                        }
                                    }
                                    else
                                    {
                                        displayValueCell.Value2 = displayValuesStr;
                                    }

                                    // Lock the cell to prevent manual editing
                                    LockToCurrentValue(displayValueCell);
                                    storedCount++;
                                    LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: Stored display values for {paramName} in cell IA{rowNumber}");
                                }
                            }
                        }
                        catch (InvalidCastException ex)
                        {
                            LogUtility.LogException(ex, $"{nameof(AddExtraParameterDisplayValues)}: InvalidCastException at row index {i}");
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: Failed at row index {i} - {ex.Message}");
                        }
                    }
                }
                finally
                {
                    if (paramTypeCol != null) Marshal.ReleaseComObject(paramTypeCol);
                    if (paramNameCol != null) Marshal.ReleaseComObject(paramNameCol);
                    if (value1Col != null) Marshal.ReleaseComObject(value1Col);
                }

                LogUtility.LogDebug($"{nameof(AddExtraParameterDisplayValues)}: Completed. Stored {storedCount} of {totalRows} rows.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(AddExtraParameterDisplayValues));
            }
        }

        private static (string Value1, string Value2) ExtractValues(JsonElement item, string paramOperator, string type)
        {
            try
            {
                if (paramOperator.Contains("NULL"))
                {
                    return (string.Empty, string.Empty);
                }

                bool isDateTime = type.ToUpperInvariant().Contains("DATE") || type.ToUpperInvariant().Contains("TIME");

                if (isDateTime)
                {
                    if (JsonHelper.TryGetProperty(item, "displayValue", out var dv) && IsPresent(dv))
                    {
                        return (XLEdgeValueFormatter.FormatDateValue(dv.ToString()), string.Empty);
                    }

                    if (JsonHelper.TryGetProperty(item, "displayValues", out var dvs) && dvs.ValueKind == JsonValueKind.Array)
                    {
                        List<string> arr = dvs.EnumerateArray().Select(v => v.ToString()).ToList();
                        if (paramOperator.Contains("BETWEEN") && arr.Count >= 2)
                        {
                            return (XLEdgeValueFormatter.FormatDateValue(arr[0]), XLEdgeValueFormatter.FormatDateValue(arr[1]));
                        }

                        if (!paramOperator.Contains("BETWEEN"))
                        {
                            return (string.Join(",", arr.Select(XLEdgeValueFormatter.FormatDateValue)), string.Empty);
                        }
                    }

                    return (string.Empty, string.Empty);
                }

                if (JsonHelper.TryGetProperty(item, "value", out var v) && IsPresent(v))
                {
                    return (JoinValues(v, string.Empty), string.Empty);
                }

                if (JsonHelper.TryGetProperty(item, "values", out var vs) && vs.ValueKind == JsonValueKind.Array)
                {
                    List<string> arr = vs.EnumerateArray().Select(x => x.ToString()).ToList();
                    if (arr.Count == 0)
                    {
                        return (string.Empty, string.Empty);
                    }

                    if (paramOperator.Contains("BETWEEN"))
                    {
                        if (arr.Count >= 2)
                        {
                            return (arr[0], arr[1]);
                        }

                        if (JsonHelper.TryGetProperty(item, "displayValues", out var dvs2) && dvs2.ValueKind == JsonValueKind.Array)
                        {
                            List<string> dvArr = dvs2.EnumerateArray().Select(x => x.ToString()).ToList();
                            if (dvArr.Count >= 2)
                            {
                                return (dvArr[0], dvArr[1]);
                            }
                        }

                        return (string.Empty, string.Empty);
                    }

                    return (JoinList(arr), string.Empty);
                }

                if (JsonHelper.TryGetProperty(item, "displayValue", out var dv2) && IsPresent(dv2))
                {
                    return (JoinValues(dv2, string.Empty), string.Empty);
                }

                if (JsonHelper.TryGetProperty(item, "displayValues", out var dvs3) && dvs3.ValueKind == JsonValueKind.Array)
                {
                    List<string> arr = dvs3.EnumerateArray().Select(x => x.ToString()).ToList();
                    if (arr.Count == 0)
                    {
                        return (string.Empty, string.Empty);
                    }

                    if (paramOperator.Contains("BETWEEN"))
                    {
                        return arr.Count >= 2 ? (arr[0], arr[1]) : (string.Empty, string.Empty);
                    }

                    return (JoinList(arr), string.Empty);
                }

                return (string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(ExtractValues)}: Failed to extract values for type {type}");
                return (string.Empty, string.Empty);
            }
        }

        private static bool IsPresent(JsonElement element) =>
            element.ValueKind != JsonValueKind.Null && element.ValueKind != JsonValueKind.Undefined;

        private static string JoinValues(JsonElement element, string fallback)
        {
            try
            {
                if (element.ValueKind == JsonValueKind.Array)
                {
                    List<string> arr = element.EnumerateArray().Select(x => x.ToString()).ToList();
                    return arr.Count == 0 ? string.Empty : JoinList(arr);
                }

                return element.ToString() ?? fallback;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(JoinValues)}: Failed to join values");
                return fallback ?? string.Empty;
            }
        }

        private static string JoinList(List<string> values) =>
            string.Join(",", values.Select(v => v.Contains(",") ? $"\"{v}\"" : v));

        private static string GetString(JsonElement element, string propertyName)
        {
            try
            {
                return JsonHelper.TryGetProperty(element, propertyName, out var value) && IsPresent(value)
                    ? value.ToString()
                    : null;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(GetString)}: Failed to get property {propertyName}");
                return null;
            }
        }

        /// <summary>Matches VB's built-in RGB() function</summary>
        private static int Rgb(int r, int g, int b) => r + (g << 8) + (b << 16);
    }
}