using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge
{
    public static class XLApp
    {
        public static Excel.Application App
            => XLEdgeAppState.Instance.ExcelApp;

        public static IntPtr Handle
            => XLEdgeAppState.Instance.ExcelHandle;

        public static bool IsAvailable
            => App != null;

        public static void Initialize(Excel.Application app)
        {
            XLEdgeAppState.Instance.InitializeExcelApplication(app);
        }

        public static void Ensure()
        {
            XLEdgeAppState.Instance.EnsureExcelApplication();
        }
        public static DateTime? GetDateFromCell(Excel.Range cell)
        {
            return XLEdgeAppState.Instance.GetDateFromCell(cell);
        }

        public static void WriteDateToCell(Excel.Range cell, DateTime dateValue)
        {
            XLEdgeAppState.Instance.WriteDateToCell(cell, dateValue);
        }
    }
}
