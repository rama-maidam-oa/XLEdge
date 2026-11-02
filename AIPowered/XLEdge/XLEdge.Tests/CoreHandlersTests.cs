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
        public void TheEntry_RunsEveryRibbonHandler_UnderExcelsThreadContext()
        {
            // Through AddinEntry.OnRibbonAction itself: without its ExcelThreadContext.Run the handler would resume elsewhere and
            // the core would not know it is still running.
            WithCore((p, _) => Assert.AreEqual("resumedOnExcelThread=True busyWhilePending=True busyAfter=False", p.AwaitingHandlerThroughTheEntry()));
        }

        [TestMethod]
        public void TitleCommands_CountFromTheInstancesOwnPagesOnly()
        {
            WithCore((p, _) =>
            {
                foreach (var (source, login, home, expected) in new (string, string, string, bool)[]
                {
                    ("http://acme:8080/app?excel=Y", "http://acme:8080", null, true),
                    ("HTTP://ACME:8080/x", "http://acme:8080", null, true),
                    ("https://acme/app", "http://acme", null, true),            // the server upgraded the configured http address
                    ("https://acme:8443/app", "http://acme", null, false),
                    ("http://acme:9090/x", "http://acme:8080", null, false),
                    ("http://evil/x", "http://acme", null, false),
                    ("https://app.acme.cloud/x", "http://acme", "https://app.acme.cloud", true), // where the login really completed
                    ("https://evil/x", "http://acme", "https://app.acme.cloud", false),
                    ("about:blank", "http://acme", null, false),
                    ("data:text/html,x", "http://acme", null, false),
                    ("http://acme/x", "", null, false),
                    ("", "http://acme", null, false),
                })
                    Assert.AreEqual(expected, p.CallStatic("XLEdge.ADXExcelTaskPane1", "IsInstancePage", source, login, home), $"{source} | {login} | {home}");
            });
        }

        [TestMethod]
        public void ShellCalls_FromAPoolThread_RunOnExcelsThread()
        {
            // The shell's ribbon and panes are Add-in Express objects: 11.1.2's pane code can reach them after an await (pool thread).
            WithCore((p, context) =>
            {
                var ribbon = (CoreRibbonFacadeTests.FakeRibbon)context.RibbonController;
                using (var window = new System.Windows.Forms.Form())
                {
                    int before = ribbon.Calls.Length;
                    p.ShellCallsFromAPoolThread(window.Handle);
                    int excelThread = System.Threading.Thread.CurrentThread.ManagedThreadId;
                    var calls = ribbon.Calls.Zip(ribbon.CallThreads, (c, th) => (c, th)).Skip(before).ToList();
                    foreach (string expected in new[] { "Label:RibSheetLabel=from a pool thread", "Enabled:RibEdgeHelp=False",
                                                        $"PaneVisible:{window.Handle}=True", $"PaneCaption:{window.Handle}=caption", $"PaneWidth:{window.Handle}=600" })
                        Assert.IsTrue(calls.Any(x => x.c == expected), expected + " in " + string.Join(" | ", calls.Select(x => x.c)));
                    var off = calls.Where(x => x.th != excelThread).Select(x => x.c + "@" + x.th).ToList();
                    Assert.AreEqual(0, off.Count, "shell calls off Excel's thread: " + string.Join(", ", off));
                }
            });
        }

        [TestMethod]
        public void AHandlerStartedOffExcelsThread_StillResumesOnIt()
        {
            WithCore((p, _) => Assert.AreEqual("resumedOnExcelThread=True busyAfter=False", p.HandlerStartedOffExcelsThread()));
        }

        [TestMethod]
        public void AnExceptionEscapingAnAsyncHandler_IsLogged_NotThrownIntoExcel()
        {
            // An IOException is the one WpfAppManager's dispatcher handler leaves unhandled: rethrown into Excel's message loop it ends Excel.
            WithCore((p, context) =>
            {
                Assert.AreEqual("reached=True busyAfter=False", p.HandlerThatThrowsAfterAnAwait());
                var log = (RecordingLogger)context.Logger;
                Assert.IsTrue(log.Lines.Any(l => l.StartsWith("EXCEPTION") && l.Contains("Unhandled exception in an XLEdge handler") && l.Contains("disk gone")), log.Dump());
            });
        }

        [TestMethod]
        public void EventArguments_AreReadAtTheShellsPositions()
        {
            WithCore((p, _) =>
            {
                Assert.AreEqual("sheet=Data workbook=Book2.xlsx address=J5 shape=null",
                    p.DescribeEventTarget("SheetSelectionChange", new object[] { "Data", "J5", 1L, "Book2.xlsx" }));
                Assert.AreEqual("sheet=Data workbook=Book2.xlsx address=null shape=null", p.DescribeEventTarget("SheetActivate", new object[] { "Data", "Book2.xlsx" }));
                Assert.AreEqual("sheet=Data workbook=Book2.xlsx address=null shape=null", p.DescribeEventTarget("SheetBeforeDelete", new object[] { "Data", "Book2.xlsx" }));
                Assert.AreEqual("sheet=P_Data workbook=Book2.xlsx address=A1 shape=null",
                    p.DescribeEventTarget("SheetFollowHyperlink", new object[] { "P_Data", "A1", "DRILLDOWN|7", "'P_Data'!A1", "1000", "Book2.xlsx", null }));
                Assert.AreEqual("sheet=Data workbook=Book2.xlsx address=null shape=Picture 3",
                    p.DescribeEventTarget("SheetFollowHyperlink", new object[] { "Data", null, "ATTACHMENT|x", null, null, "Book2.xlsx", "Picture 3" }));
                Assert.AreEqual("sheet=null workbook=Book2.xlsx address=null shape=null", p.DescribeEventTarget("WorkbookActivate", new object[] { "Book2.xlsx" }));
                // A shell that sends fewer arguments (no workbook): the active workbook is used.
                Assert.AreEqual("sheet=Data workbook=null address=J5 shape=null", p.DescribeEventTarget("SheetSelectionChange", new object[] { "Data", "J5", 1L }));
                Assert.AreEqual("sheet=null workbook=null address=null shape=null", p.DescribeEventTarget("SheetActivate", null));
                Assert.AreEqual("null", p.DescribeEventTarget("NoSuchEvent", new object[] { "x" }));
            });
        }

        [TestMethod]
        public void EveryForwardedButton_ReachesAHandler()
        {
            WithCore((p, context) =>
            {
                p.AnswerMessages();
                foreach (string button in new[] { "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeOptions", "RibEdgeDialogBoxLauncher", "RibEdgeHelp",
                                                  "RibControlSheet", "RibEdgeIncludeOutputData", "RibEdgeDebug",
                                                  // Stage 2e
                                                  "RibEdgeLogin", "RibEdgeShowHide", "RibEdgeParamRefresh", "RibEdgeParamRefreshBook", "RibEdgeAbout",
                                                  // OISR-22777
                                                  "RibEdgeBroadcast" })
                    p.EntryRibbonAction(button, true);
                p.EntryRibbonAction("RibEdgeLogout_PropertyChanging", "Visible");
                p.StopAnsweringMessages();
                var log = (RecordingLogger)context.Logger;
                Assert.IsFalse(log.Contains("no handler for ribbon action"), log.Dump());
                Assert.IsTrue(context.DebugMode, "RibEdgeDebug: debug logging follows the button");
                string[] shown = p.MessagesShown();
                CollectionAssert.Contains(shown, "window:XLEdgeServerConfiguration", "Login opens 11.1.2's Server Configuration window");
                CollectionAssert.Contains(shown, "window:XLEdgeAbout", "About opens 11.1.2's About window");
                CollectionAssert.Contains(log.Events, "flush:Debug session ended", "RibEdgeDebug flushes the open buffers, as 11.1.2");
            });
        }

        [TestMethod]
        public void EachCore_StartsInXLEdge11_1_2sRibbonLoadState()
        {
            // 11.1.2's OnRibbonLoaded: LoggedOut until a login; IncludeOutputData's flag follows its (now un-pressed) button.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                {
                    FakeContext context = NewContext(t);
                    var ribbon = (CoreRibbonFacadeTests.FakeRibbon)context.RibbonController;
                    ribbon.SetControlPressed("RibEdgeIncludeOutputData", true);
                    CoreOnThisThread(context, p =>
                    {
                        Assert.IsTrue(ribbon.Calls.Contains("Enabled:RibEdgeRefresh=False"), string.Join(" | ", ribbon.Calls));
                        Assert.IsTrue(ribbon.Calls.Contains("Visible:RibEdgeLogin=True"), string.Join(" | ", ribbon.Calls));
                        Assert.IsFalse(ribbon.GetControlPressed("RibEdgeIncludeOutputData"));
                        Assert.AreEqual(false, p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "DebugOutputData"));
                    });
                }
            });
        }

        [TestMethod]
        public void ANewCore_KeepsTheShellsDebugToggle()
        {
            // The Debug button and debug logging are the shell's and outlive a reload: a reload is not a logout.
            Win32Test.OnSta(() =>
            {
                using (var t = new TempFolder())
                {
                    FakeContext context = NewContext(t);
                    context.DebugMode = true;
                    var ribbon = (CoreRibbonFacadeTests.FakeRibbon)context.RibbonController;
                    ribbon.SetControlPressed("RibEdgeDebug", true);
                    CoreOnThisThread(context, p =>
                    {
                        Assert.IsTrue(context.DebugMode, "debug logging stays on");
                        Assert.IsTrue(ribbon.GetControlPressed("RibEdgeDebug"), "the Debug button stays pressed");
                    });
                }
            });
        }

        [TestMethod]
        public void ARunningReport_KeepsTheCoreBusy()
        {
            // ProcessRunning is what ReportGenerator's ExcelBulkOperationScope sets for a whole report run or refresh.
            WithCore((p, context) =>
            {
                p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "ProcessRunning", true);
                Assert.IsTrue(p.EntryIsBusy(), "a report is running");
                Assert.IsTrue(((RecordingLogger)context.Logger).Contains("AddinEntry.IsBusy: busy (dialog open=False, handlers running=0, background=0, report running=True)."),
                    "the reason a reload is refused is logged");
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
        public void RunWorkbookScope_WithoutExcel_SetsNothingAndNeverThrows()
        {
            // Stage 4: with no Excel there is no workbook to keep; a run scope must still open and close cleanly.
            WithCore((p, _) =>
            {
                Assert.AreEqual("override=null busyInside=True busyAfter=False overrideAfter=null", p.BulkScopeWithoutExcel());
            });
        }

        [TestMethod]
        public void TrackedBackgroundWork_KeepsTheCoreBusyUntilItEnds()
        {
            // Stage 4: the pane's background work (title-command reports, XLSX downloads) runs in Task.Run - a reload must wait.
            WithCore((p, _) =>
            {
                foreach (bool fault in new[] { false, true })
                {
                    p.StartTrackedWork();
                    Assert.IsTrue(p.EntryIsBusy(), "busy while it runs");
                    p.EndTrackedWork(fault);
                    Assert.IsTrue(Win32Test.PumpUntil(() => !p.EntryIsBusy(), TimeSpan.FromSeconds(5)), "not busy once it ended (fault=" + fault + ")");
                }
            });
        }

        [TestMethod]
        public void CountedHandler_KeepsTheCoreBusyWhileTheAsyncVoidHandlerAwaits()
        {
            // Stage 4: the pane's WebView2 handlers (login cookies, broadcast, title commands) are async void.
            WithCore((p, _) =>
            {
                p.RaiseCountedHandler();
                Assert.IsTrue(p.EntryIsBusy(), "busy while the handler awaits");
                p.EndCountedHandler();
                Assert.IsTrue(Win32Test.PumpUntil(() => !p.EntryIsBusy(), TimeSpan.FromSeconds(5)), "not busy once it ended");
            });
        }

        [TestMethod]
        public void PaneBackgroundWork_IsTracked()
        {
            string pane = File.ReadAllText(Path.Combine(Path.GetDirectoryName(typeof(CoreHandlersTests).Assembly.Location),
                @"..\..\..\XLEdge.Addin.Core\ADXExcelTaskPane1.cs")).Replace("\r\n", "\n");
            StringAssert.Contains(CorePortParityTests.MethodText(pane, "void SafeFireAndForget("), "ExcelThreadContext.Track(Task.Run(");
            string hook = CorePortParityTests.MethodText(pane, "void HookWebViewEvents()");
            StringAssert.Contains(hook, "SourceChanged += _countedSourceChanged;");
            StringAssert.Contains(hook, "DocumentTitleChanged += _countedTitleChanged;");
            StringAssert.Contains(pane, "ExcelThreadContext.Counted<CoreWebView2SourceChangedEventArgs>(WebCtrl_SourceChanged)");
            StringAssert.Contains(pane, "ExcelThreadContext.Counted<object>(InstanceDocumentTitleChanged)");
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
        public void Broadcast_ShowsTheInstancesMessages_AsInXLEdge11_1_2()
        {
            // OISR-22777: the ribbon's Broadcast button asks the instance for its messages and shows them, or says there are none.
            WithCore((p, _) =>
            {
                using (var instance = new InstanceServer())
                {
                    p.AnswerMessages();
                    p.EntryRibbonAction("RibEdgeBroadcast", null);
                    Assert.IsFalse(Win32Test.PumpUntil(() => p.MessagesShown().Length > 0, TimeSpan.FromSeconds(1)), "logged out: nothing shown");

                    p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "LoginUrl", instance.Url + " ");
                    p.OnSingleton("XLEdge.XLEdgeAppState", "Instance", "IsLoginCompleted", true);
                    p.EntryRibbonAction("RibEdgeBroadcast", null);
                    Assert.IsTrue(Win32Test.PumpUntil(() => p.MessagesShown().Length == 1, TimeSpan.FromSeconds(15)), "no messages");

                    instance.Broadcast = "[{\"msgType\":\"Alert\",\"message\":\"Maintenance at 6 PM\"},{\"message\":\"Welcome\"}]";
                    p.EntryRibbonAction("RibEdgeBroadcast", null);
                    Assert.IsTrue(Win32Test.PumpUntil(() => p.MessagesShown().Length == 2, TimeSpan.FromSeconds(15)), "two messages");
                    p.StopAnsweringMessages();

                    CollectionAssert.AreEqual(new[] { "No broadcast messages available.", "1.) Alert : Maintenance at 6 PM" + Environment.NewLine + "2.) Info : Welcome" },
                                              p.MessagesShown());
                    Assert.AreEqual(2, instance.Requests.Count(r => r.StartsWith("POST /web/secure/get-broadcast-msg ", StringComparison.Ordinal)),
                                    "one request per click once logged in, none before");
                    Assert.IsFalse(p.EntryIsBusy());
                }
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
        public void UnknownActionsAndEvents_AreLoggedAndIgnored()
        {
            WithCore((p, context) =>
            {
                p.EntryRibbonAction("NoSuchButton", null);
                Assert.IsTrue(p.EntryExcelEvent("NoSuchEvent", new object[] { "x" }));
                var log = (RecordingLogger)context.Logger;
                Assert.IsTrue(log.Contains("no handler for ribbon action 'NoSuchButton'"), log.Dump());
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
