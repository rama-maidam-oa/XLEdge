using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XLEdge.Tests.CoreWindowsTests;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 2d: the ribbon actions and Excel events the shell forwards reach XLEdge 11.1.2's handlers in the core, their async
    /// continuations stay on Excel's thread, a running handler or report keeps the core busy, and nothing of them outlives Shutdown.
    /// (Handlers working on a real workbook: CoreHandlersExcelTests.)
    /// </summary>
    [TestClass]
    public class CoreHandlersTests
    {
        private static void WithCore(Action<CoreProbe, FakeContext> use) => Win32Test.OnSta(() =>
        {
            using (var t = new TempFolder())
            {
                FakeContext context = NewContext(t);
                CoreOnThisThread(context, p => use(p, context));
            }
        });

        [TestMethod]
        public void AsyncHandler_ResumesOnExcelsThread_AndKeepsTheCoreBusyUntilItEnds()
        {
            WithCore((p, _) => Assert.AreEqual("resumedOnExcelThread=True busyWhilePending=True busyAfter=False", p.AsyncHandlerOnTheExcelThread()));
        }

        [TestMethod]
        public void ARunningReport_KeepsTheCoreBusy()
        {
            // ProcessRunning is what ReportGenerator's ExcelBulkOperationScope sets for a whole report run or refresh.
            WithCore((p, _) =>
            {
                p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "ProcessRunning", true);
                Assert.IsTrue(p.EntryIsBusy(), "a report is running");
                p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "ProcessRunning", false);
                Assert.IsFalse(p.EntryIsBusy(), "the report finished");
            });
        }

        [TestMethod]
        public void Shutdown_StopsTheDeferredDeleteTimer_AndForgetsTheModule()
        {
            WithCore((p, _) =>
            {
                p.QueueADeferredDelete();
                Assert.AreEqual("timerEnabled=True queued=1 current=True", p.DeferredDeleteState(), "precondition");
                p.ShutDownTheEntry();
                Assert.AreEqual("timerEnabled=False queued=0 current=False", p.DeferredDeleteState());
            });
        }

        [TestMethod]
        public void TempFilesLeftByThePreviousCore_AreDeletedWhenTheNextOneStarts()
        {
            // 11.1.2 cleared XLEdgeAppPaths.TempFolder when its ribbon loaded (once per Excel session); a reloaded core is a new start.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                {
                    FakeContext context = NewContext(t);
                    string leftOver = null;
                    CoreOnThisThread(context, p =>
                    {
                        string folder = (string)p.CallStatic("XLEdge.XLEdgeAppPaths", "TempFolder");
                        Directory.CreateDirectory(folder);
                        leftOver = Path.Combine(folder, "report_left_over.csv");
                        File.WriteAllText(leftOver, "a,b");
                    });
                    Assert.IsTrue(File.Exists(leftOver), "precondition: a temp report file is left");
                    CoreOnThisThread(context, p => { });
                    Assert.IsFalse(File.Exists(leftOver), "the next core cleared the temp folder");
                }
            });
        }

        [TestMethod]
        public void IncludeOutputData_SetsTheFlag_TheReportCodeReads()
        {
            WithCore((p, _) =>
            {
                p.EntryRibbonAction("RibEdgeIncludeOutputData", true);
                Assert.AreEqual(true, p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "DebugOutputData"));
                p.EntryRibbonAction("RibEdgeIncludeOutputData", false);
                Assert.AreEqual(false, p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "DebugOutputData"));
            });
        }

        [TestMethod]
        public void LoginGatedButtons_BeforeLogin_AskToLogIn_AsInXLEdge11_1_2()
        {
            WithCore((p, _) =>
            {
                p.AnswerMessages();
                foreach (string button in new[] { "RibEdgeOptions", "RibEdgeDialogBoxLauncher", "RibEdgeHelp" })
                    p.EntryRibbonAction(button, null);
                p.StopAnsweringMessages();
                CollectionAssert.AreEqual(Enumerable.Repeat("Please log in to the instance.", 3).ToArray(), p.MessagesShown());
            });
        }

        [TestMethod]
        public void RefreshButtons_BeforeLogin_DoNothing()
        {
            WithCore((p, context) =>
            {
                p.AnswerMessages();
                p.EntryRibbonAction("RibEdgeRefresh", null);
                p.EntryRibbonAction("RibEdgeRefreshAll", null);
                p.StopAnsweringMessages();
                CollectionAssert.AreEqual(new string[0], p.MessagesShown());
                Assert.IsFalse(p.EntryIsBusy());
            });
        }

        [TestMethod]
        public void UnknownAndNotYetPortedActions_AreLoggedAndIgnored()
        {
            WithCore((p, context) =>
            {
                p.EntryRibbonAction("NoSuchButton", null);
                p.EntryRibbonAction("RibEdgeAbout", null);
                Assert.IsTrue(p.EntryExcelEvent("NoSuchEvent", new object[] { "x" }));
                var log = (RecordingLogger)context.Logger;
                Assert.IsTrue(log.Contains("no handler for ribbon action 'NoSuchButton'"), log.Dump());
                Assert.IsTrue(log.Contains("ribbon action 'RibEdgeAbout' is not in the core yet"), log.Dump());
                Assert.IsTrue(log.Contains("no handler for Excel event 'NoSuchEvent'"), log.Dump());
                Assert.IsFalse(log.Lines.Any(l => l.StartsWith("EXCEPTION") || l.StartsWith("ERROR")), log.Dump());
            });
        }

        [TestMethod]
        public void ExcelEvents_WithoutExcelOrArguments_NeverThrowBackToTheShell()
        {
            // The FakeContext has no Excel: every event must end quietly (logged), whatever the shell sends.
            WithCore((p, context) =>
            {
                foreach (string name in new[] { "SheetSelectionChange", "SheetActivate", "WorkbookActivate", "SheetFollowHyperlink", "SheetBeforeDelete" })
                {
                    Assert.IsTrue(p.EntryExcelEvent(name, new object[] { "Sheet1", "A1", 1L, "screen", "sub", "text", "Book1" }), name);
                    Assert.IsTrue(p.EntryExcelEvent(name, new object[0]), name + " without arguments");
                    Assert.IsTrue(p.EntryExcelEvent(name, null), name + " with null arguments");
                }
                Assert.IsFalse(p.EntryIsBusy());
                var log = (RecordingLogger)context.Logger;
                Assert.IsFalse(log.Lines.Any(l => l.StartsWith("EXCEPTION") || l.StartsWith("ERROR")), "missing Excel or arguments are not errors: " + log.Dump());
            });
        }
    }
}
