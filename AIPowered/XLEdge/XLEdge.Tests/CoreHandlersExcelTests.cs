using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XLEdge.Tests.CoreWindowsTests;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 2d: XLEdge 11.1.2's handlers in the core, against a real Excel - an automation instance (no add-ins load in it), driven
    /// the way the shell drives the core: Excel events arrive as names, the core works on Excel through its own RCW from its own
    /// AppDomain. Each test ends with Excel gone: still running after Quit means a COM reference was left behind. The core's domain
    /// is unloaded before Quit (as when Excel closes), so this sees references held outside the core, not the core's own RCWs.
    /// </summary>
    [TestClass]
    [TestCategory("Excel")]
    public class CoreHandlersExcelTests
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        /// <summary>An automation Excel with one workbook; quits on Dispose and fails the test if the process outlives it.</summary>
        private sealed class AutomationExcel : IDisposable
        {
            public dynamic App { get; }
            public dynamic Book { get; private set; }
            private readonly Process _process;

            public AutomationExcel()
            {
                Type type = Type.GetTypeFromProgID("Excel.Application");
                if (type == null) Assert.Inconclusive("Excel is not installed");
                App = Activator.CreateInstance(type);
                App.Visible = false;
                App.DisplayAlerts = false;
                GetWindowThreadProcessId(new IntPtr((int)App.Hwnd), out uint pid);
                _process = Process.GetProcessById((int)pid);
                int typed = TypedInterfaceResult();
                if (typed != 0)
                {
                    Quit();
                    // Inside Excel (the real add-in) the core's calls are direct; from another process every typed interface is
                    // marshaled with Excel's registered type library, which this machine's Office cannot load.
                    Assert.Inconclusive($"Excel's typed interfaces cannot be marshaled to another process here (QueryInterface _Application: 0x{typed:X8}; " +
                                        "0x80029C4A = TYPE_E_CANTLOADLIBRARY, an Office registration problem). The core's Excel code is early-bound, so it cannot be driven from a test process.");
                }
                Book = App.Workbooks.Add();
            }

            /// <summary>QueryInterface for Excel's _Application from this process: 0 when typed (early-bound) calls can work.</summary>
            private int TypedInterfaceResult()
            {
                Guid application = new Guid("000208D5-0000-0000-C000-000000000046");
                IntPtr unknown = Marshal.GetIUnknownForObject((object)App);
                try
                {
                    int hr = Marshal.QueryInterface(unknown, ref application, out IntPtr typed);
                    if (typed != IntPtr.Zero) Marshal.Release(typed);
                    return hr;
                }
                finally
                {
                    Marshal.Release(unknown);
                }
            }

            public dynamic AddBook() => App.Workbooks.Add();

            /// <summary>A sheet with an XLEdge report table (ListObject) named <paramref name="table"/> from A1, header plus two rows.</summary>
            public static dynamic AddReportSheet(dynamic book, string name, string table)
            {
                dynamic sheet = book.Worksheets.Add();
                sheet.Name = name;
                sheet.Range["A1"].Value2 = "ACCOUNT";
                sheet.Range["B1"].Value2 = "AMOUNT";
                sheet.Range["A2"].Value2 = "1000";
                sheet.Range["B2"].Value2 = 12.5;
                sheet.Range["A3"].Value2 = "2000";
                sheet.Range["B3"].Value2 = 7;
                dynamic list = sheet.ListObjects.Add(1 /* xlSrcRange */, sheet.Range["A1:B3"], Type.Missing, 1 /* xlYes */);
                list.Name = table;
                return sheet;
            }

            /// <summary>A separate-sheet mode parameter sheet: IT2 names the report table it belongs to.</summary>
            public static dynamic AddParameterSheet(dynamic book, string name, string table)
            {
                dynamic sheet = book.Worksheets.Add();
                sheet.Name = name;
                sheet.Range["IT2"].Value2 = table;
                return sheet;
            }

            public string[] SheetNames()
            {
                var names = new System.Collections.Generic.List<string>();
                foreach (dynamic sheet in Book.Worksheets) names.Add((string)sheet.Name);
                return names.ToArray();
            }

            /// <summary>Quits Excel; true when its process ended on its own (it is killed otherwise, so no test leaves one behind).</summary>
            public bool Quit()
            {
                try
                {
                    foreach (dynamic book in App.Workbooks) book.Close(false);
                    App.Quit();
                }
                catch (COMException) { }
                Book = null;
                Marshal.FinalReleaseComObject(App);
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                // Pumping: the finalizer releases this STA thread's RCWs through its message loop.
                Win32Test.PumpUntil(() => _process.HasExited, TimeSpan.FromSeconds(15));
                bool exited = _process.HasExited;
                if (!exited) _process.Kill();
                return exited;
            }

            public void Dispose() { if (!_process.HasExited) Quit(); }
        }

        /// <summary>One core with <paramref name="excel"/> as its Excel, on an STA thread like Excel's, messages answered.</summary>
        private static void WithCoreOn(AutomationExcel excel, TempFolder t, Action<CoreProbe, FakeContext> use)
        {
            FakeContext context = NewContext(t);
            context.ExcelApp = excel.App;
            CoreOnThisThread(context, p =>
            {
                p.AnswerMessages();
                try { use(p, context); }
                finally { p.StopAnsweringMessages(); }
            });
        }

        private static void InExcel(Action<AutomationExcel, TempFolder> test) => Win32Test.OnSta(() =>
        {
            bool exited;
            using (var t = new TempFolder())
            {
                var excel = new AutomationExcel();
                // Quit in finally, checked after: a failing check is reported as itself, not as Excel left running.
                try { test(excel, t); }
                finally { exited = excel.Quit(); }
            }
            Assert.IsTrue(exited, "Excel was still running 15 s after Quit: a COM reference to it was left behind");
        });

        [TestMethod]
        public void DeletingAReportSheet_DeletesItsParameterSheet_ALittleLater()
        {
            InExcel((excel, t) =>
            {
                AutomationExcel.AddReportSheet(excel.Book, "Data", "ORB_1_E");
                AutomationExcel.AddParameterSheet(excel.Book, "P_Data", "ORB_1_E");
                WithCoreOn(excel, t, (p, context) =>
                {
                    p.EntryExcelEvent("SheetBeforeDelete", new object[] { "Data", (string)excel.Book.Name });
                    CollectionAssert.Contains(excel.SheetNames(), "P_Data", "deferred: not inside SheetBeforeDelete");
                    Win32Test.PumpUntil(() => !excel.SheetNames().Contains("P_Data"), TimeSpan.FromSeconds(5));
                    CollectionAssert.DoesNotContain(excel.SheetNames(), "P_Data", ((RecordingLogger)context.Logger).Dump());
                    CollectionAssert.Contains(excel.SheetNames(), "Data", "the report sheet itself is Excel's to delete");
                });
            });
        }

        [TestMethod]
        public void GotoReportData_ActivatesTheReportsSheet()
        {
            InExcel((excel, t) =>
            {
                AutomationExcel.AddReportSheet(excel.Book, "Data", "ORB_1_E");
                dynamic parameters = AutomationExcel.AddParameterSheet(excel.Book, "P_Data", "ORB_1_E");
                parameters.Hyperlinks.Add(parameters.Range["A1"], "", "'P_Data'!A1", Type.Missing, "Goto Report Data");
                parameters.Activate();
                WithCoreOn(excel, t, (p, context) =>
                {
                    p.EntryExcelEvent("SheetFollowHyperlink", new object[] { "P_Data", "A1", null, "'P_Data'!A1", "Goto Report Data", (string)excel.Book.Name });
                    Assert.AreEqual("Data", (string)excel.App.ActiveSheet.Name, ((RecordingLogger)context.Logger).Dump());
                });
            });
        }

        [TestMethod]
        public void TabLabel_DescribesTheActiveSheetsReport_AndOrphanedNamesGo()
        {
            InExcel((excel, t) =>
            {
                dynamic data = AutomationExcel.AddReportSheet(excel.Book, "Data", "ORB_1_E");
                data.Range["A2"].Hyperlinks.Add(data.Range["A2"], "", "A2", "DRILLDOWN|7|Child|1", "1000");
                AutomationExcel.AddParameterSheet(excel.Book, "P_Data", "ORB_1_E");
                excel.Book.Names.Add("ORB_Gone_ChildReport", "=1");
                excel.Book.Names.Add("ORB_Data_Instance", "=1");
                data.Activate();
                WithCoreOn(excel, t, (p, context) =>
                {
                    p.UpdateTabLabelOfTheActiveSheet();
                    var ribbon = (CoreRibbonFacadeTests.FakeRibbon)context.RibbonController;
                    Assert.AreEqual("This sheet has drilldown links on column(s): ACCOUNT", ribbon.GetControlLabel("RibSheetLabel"), ((RecordingLogger)context.Logger).Dump());
                    var names = new System.Collections.Generic.List<string>();
                    foreach (dynamic name in excel.Book.Names) names.Add((string)name.Name);
                    CollectionAssert.DoesNotContain(names, "ORB_Gone_ChildReport", "its sheet does not exist");
                    CollectionAssert.Contains(names, "ORB_Data_Instance", "its sheet exists");
                });
            });
        }

        [TestMethod]
        public void ControlSheet_IsBuilt_AndRebuiltAfterTheQuestion()
        {
            InExcel((excel, t) =>
            {
                WithCoreOn(excel, t, (p, context) =>
                {
                    p.EntryRibbonAction("RibControlSheet", null);
                    CollectionAssert.Contains(excel.SheetNames(), "Parameters Control Sheet", ((RecordingLogger)context.Logger).Dump());
                    p.EntryRibbonAction("RibControlSheet", null);
                    StringAssert.StartsWith(p.MessagesShown().LastOrDefault() ?? "", "Parameters Control Sheet already exists");
                    Assert.AreEqual(1, excel.SheetNames().Count(n => n == "Parameters Control Sheet"));
                    var log = (RecordingLogger)context.Logger;
                    Assert.IsFalse(log.Lines.Any(l => l.StartsWith("EXCEPTION")), log.Dump());
                });
            });
        }

        [TestMethod]
        public void AnEventNamingAnotherWorkbook_ResolvesThatWorkbooksSheet()
        {
            InExcel((excel, t) =>
            {
                AutomationExcel.AddReportSheet(excel.Book, "Data", "ORB_1_E");
                dynamic other = excel.AddBook();
                AutomationExcel.AddReportSheet(other, "Data", "ORB_2_E");
                other.Activate();
                WithCoreOn(excel, t, (p, context) =>
                {
                    Assert.AreEqual((string)excel.Book.Name + "!ORB_1_E", p.ResolveEventSheet("Data", (string)excel.Book.Name));
                    Assert.AreEqual((string)other.Name + "!ORB_2_E", p.ResolveEventSheet("Data", null), "no workbook sent: the active one");
                    Assert.AreEqual("null", p.ResolveEventSheet("Data", "NoSuchBook.xlsx"));
                    Assert.AreEqual("null", p.ResolveEventSheet("NoSuchSheet", (string)excel.Book.Name));
                });
            });
        }

        [TestMethod]
        public void TheShell_SendsTheSheetsWorkbook_WithEverySheetEvent()
        {
            InExcel((excel, t) =>
            {
                dynamic sheet = AutomationExcel.AddReportSheet(excel.Book, "Data", "ORB_1_E");
                var workbookOf = typeof(XLEdge.AddinModule).GetMethod("WorkbookName", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
                Assert.IsNotNull(workbookOf, "XLEdge.AddinModule.WorkbookName");
                Assert.AreEqual((string)excel.Book.Name, workbookOf.Invoke(null, new object[] { (object)sheet }));
                Assert.IsNull(workbookOf.Invoke(null, new object[] { null }));
            });
        }
    }
}
