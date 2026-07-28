using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Worksheet / named-range utility functions: checking whether a sheet exists, locating a
    /// parameter sheet by TableID, and checking/creating/deleting Excel named ranges. Each COM
    /// object obtained from an enumerator (Worksheets, Names) is released as it goes.
    /// </summary>
    public static class ExcelSheetHelper
    {
        /// <summary>
        /// Checks whether a worksheet with the given name exists in the supplied workbook,
        /// or the active workbook if none is supplied.
        /// </summary>
        public static bool SheetExists(string sheetName, Excel.Workbook workbook = null)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                return false;
            }

            Excel.Sheets sheets = null;

            try
            {
                Excel.Workbook targetWorkbook = workbook;
                if (targetWorkbook == null)
                {
                    Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                    targetWorkbook = excelApp.ActiveWorkbook;
                }

                if (targetWorkbook == null)
                {
                    return false;
                }

                // Workbook.Worksheets actually returns an Excel.Sheets COM object at runtime, not
                // Excel.Worksheets - casting/declaring it as Excel.Worksheets throws
                // InvalidCastException (E_NOINTERFACE) at runtime even though it compiles fine.
                sheets = targetWorkbook.Worksheets;

                foreach (Excel.Worksheet ws in sheets)
                {
                    try
                    {
                        if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
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
                LogUtility.LogException(ex, nameof(SheetExists));
                return false;
            }
            finally
            {
                if (sheets != null)
                {
                    Marshal.ReleaseComObject(sheets);
                }
            }
        }

        /// <summary>
        /// Finds the worksheet whose "IT2" cell identifies a given TableID, first checking the
        /// supplied SheetName (honoring the "_E"/"_P" suffix pairing rule) before falling back
        /// to a full workbook scan. Returns null if no match is found.
        /// </summary>
        public static Excel.Worksheet GetParameterSheet(string sheetName, string tableId)
        {
            Excel.Application excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
            if (excelApp?.ActiveWorkbook == null)
            {
                return null;
            }

            if (!string.IsNullOrWhiteSpace(sheetName) && SheetExists(sheetName, excelApp.ActiveWorkbook))
            {
                Excel.Worksheet sht = null;
                try
                {
                    sht = (Excel.Worksheet)excelApp.ActiveWorkbook.Worksheets[sheetName];

                    string tableBoundName = Convert.ToString(sht.Range["IT2"].Value) ?? string.Empty;
                    if (string.IsNullOrEmpty(tableBoundName))
                    {
                        return null;
                    }

                    bool matches;
                    if (tableId.EndsWith("_E", StringComparison.Ordinal) || tableId.EndsWith("_P", StringComparison.Ordinal))
                    {
                        string[] parts1 = tableId.Split('_');
                        string[] parts2 = tableBoundName.Split('_');
                        matches = parts1.Length > 1 && parts2.Length > 1 && parts1[1] == parts2[1];
                    }
                    else
                    {
                        matches = tableBoundName == tableId;
                    }

                    if (!matches)
                    {
                        return null;
                    }

                    Excel.Worksheet result = sht;
                    sht = null; // ownership transferred to the caller - do not release below
                    return result;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(GetParameterSheet));
                    return null;
                }
                finally
                {
                    if (sht != null)
                    {
                        Marshal.ReleaseComObject(sht);
                    }
                }
            }

            Excel.Sheets allSheets = null;
            try
            {
                allSheets = excelApp.ActiveWorkbook.Worksheets;
                foreach (Excel.Worksheet ws in allSheets)
                {
                    bool release = true;
                    try
                    {
                        object it2Value = ws.Range["IT2"].Value;
                        if (it2Value != null && Convert.ToString(it2Value) == tableId)
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
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(GetParameterSheet));
                return null;
            }
            finally
            {
                if (allSheets != null)
                {
                    Marshal.ReleaseComObject(allSheets);
                }
            }
        }

        public static bool NamedRangeExists(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            Excel.Names names = null;
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                names = excelApp.ActiveWorkbook.Names;

                foreach (Excel.Name xlName in names)
                {
                    try
                    {
                        if (xlName.Name == name)
                        {
                            return true;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(xlName);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(NamedRangeExists));
                return false;
            }
            finally
            {
                if (names != null)
                {
                    Marshal.ReleaseComObject(names);
                }
            }
        }

        public static void DeleteNamedRange(string rangeName)
        {
            if (string.IsNullOrWhiteSpace(rangeName))
            {
                return;
            }

            Excel.Names names = null;
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                names = excelApp.ActiveWorkbook.Names;

                foreach (Excel.Name xlName in names)
                {
                    try
                    {
                        if (xlName.Name == rangeName)
                        {
                            xlName.Delete();
                            return;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(xlName);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(DeleteNamedRange));
            }
            finally
            {
                if (names != null)
                {
                    Marshal.ReleaseComObject(names);
                }
            }
        }

        public static void CreateNamedRange(Excel.Range range, string name)
        {
            try
            {
                string cleanName = CleanUpName(name);
                if (NamedRangeExists(cleanName))
                {
                    DeleteNamedRange(cleanName);
                }

                range.Name = cleanName;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(CreateNamedRange));
            }
        }

        /// <summary>
        /// Finds the 1-based column index within a header range whose cell text matches
        /// headerValue (case-insensitive, trimmed). Returns 0 if not found.
        /// </summary>
        public static int HRMatch(Excel.Range headerRange, string headerValue)
        {
            if (headerRange == null || string.IsNullOrEmpty(headerValue))
            {
                return 0;
            }

            foreach (Excel.Range cell in headerRange.Cells)
            {
                try
                {
                    if (cell.Value2 != null &&
                        string.Equals(cell.Value2.ToString().Trim(), headerValue.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        return cell.Column;
                    }
                }
                catch (Exception ex)
                {
                    // Safe to ignore: unreadable/unusual cell in the header row - keep scanning the rest.
                    LogUtility.LogDebug($"{nameof(HRMatch)}: skipping unreadable header cell at column {cell.Column} - {ex.Message}");
                }
            }

            return 0;
        }

        private static readonly Regex InvalidNameChars = new Regex("[^a-zA-Z0-9_]", RegexOptions.Compiled);

        /// <summary>Strips anything that isn't alphanumeric/underscore, matching Excel named-range rules.</summary>
        public static string CleanUpName(string name)
        {
            try
            {
                string cleaned = InvalidNameChars.Replace(name ?? string.Empty, string.Empty);
                return cleaned;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(CleanUpName));
                return string.Empty;
            }
        }
    }
}
