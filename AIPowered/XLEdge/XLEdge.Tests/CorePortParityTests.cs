using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// XLEdge 11.1.2 code moves into the core unchanged unless a ledgered ruling says otherwise: every file ported verbatim must
    /// stay byte-identical to its original (BOM and line endings aside), so a later edit to one copy cannot drift silently.
    /// </summary>
    [TestClass]
    public class CorePortParityTests
    {
        private static string TestBin => Path.GetDirectoryName(typeof(CorePortParityTests).Assembly.Location);
        private static string Today => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\..\..\XLEdge"));
        private static string Core => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\XLEdge.Addin.Core"));

        /// <summary>Ported unchanged (relative to the project folder in both).</summary>
        internal static readonly string[] Verbatim =
        {
            // Stage 2a
            @"Helpers\ExcelSheetHelper.cs", @"Utilities\DpiAwarenessHelper.cs", @"Utilities\WpfAppManager.cs", @"Utilities\UiDispatcher.cs", "XLApp.cs",
            // Stage 2b
            @"Helpers\ApiErrorMessageExtractor.cs", @"Helpers\XLEdgeValueFormatter.cs", @"Helpers\ApiResponseHelper.cs",
            @"Helpers\ImageDownloadHelper.cs", @"Helpers\XLEdgePreferencesManager.cs", @"Helpers\NumericJsonConverter.cs", @"Helpers\PerformanceHelper.cs",
            @"Helpers\JsonGlobals.cs", @"Helpers\CancellationHelper.cs", @"Helpers\AttachmentLinkHelper.cs",
            @"Helpers\XLEdgeTempFileCleaner.cs", @"Helpers\CancellationTokenHelper.cs", @"Helpers\ExceptionHelper.cs", @"Helpers\ApiExceptions.cs",
            @"Helpers\XLEdgeOperatorMappings.cs", @"Helpers\EdgeRequestParser.cs", @"Helpers\SerializationHelper.cs",
            @"Models\ReportMetaModels.cs", @"Models\AllModels.cs", @"Models\ApiResult.cs", @"Utilities\StrictCertificateValidator.cs", "XLEdgeAppConstants.cs",
            // Stage 2c
            @"Helpers\ApiOperationHelper.cs", @"Helpers\APIHelper.cs", @"Helpers\ExcelWindowHelper.cs", @"Helpers\EnhancedDragDropHelper.cs",
            @"Utilities\DpiAwareWindow.cs", @"Utilities\MessageFunctions.cs", @"Utilities\WpfUiBootstrapper.cs", @"Converters\WidthPercentageConverter.cs",
            @"Themes\GlobalStyles.xaml",
            @"Views\AppOverlay.xaml", @"Views\AppOverlay.xaml.cs",
            @"Views\XLEdgeCalendar.xaml", @"Views\XLEdgeCalendar.xaml.cs", @"Views\XLEdgeDrilldownReports.xaml", @"Views\XLEdgeDrilldownReports.xaml.cs",
            @"Views\XLEdgeGLAccountsWindow.xaml", @"Views\XLEdgeGLAccountsWindow.xaml.cs", @"Views\XLEdgeLoginDetails.xaml", @"Views\XLEdgeLoginDetails.xaml.cs",
            @"Views\XLEdgeMessageWindow.xaml", @"Views\XLEdgeMessageWindow.xaml.cs", @"Views\XLEdgeOptions.xaml", @"Views\XLEdgeOptions.xaml.cs",
            @"Views\XLEdgeWaitWindow.xaml", @"Views\XLEdgeWaitWindow.xaml.cs",
            // Stage 2d
            @"Helpers\ReportGenerator.cs", @"Helpers\ParamsControlSheetBuilder.cs", @"Helpers\XLEdgeParamsBuilder.cs", @"Helpers\DrilldownRequestBuilder.cs",
            @"Helpers\ReportParameterRequestSerializer.cs", @"Helpers\ProgressCoordinator.cs",
            // Stage 2e
            @"Views\XLEdgeServerConfiguration.xaml", @"Views\XLEdgeServerConfiguration.xaml.cs", @"Views\XLEdgeAbout.xaml", @"Views\XLEdgeAbout.xaml.cs",
        };

        /// <summary>Images, compared byte for byte (WPF resources).</summary>
        internal static readonly string[] Binary = { @"Images\orbit_logo.png", @"Images\OrbitGLSense.ico", @"Images\orbit_bitmap.bmp" };

        /// <summary>
        /// The one change a verbatim port may carry: pack URIs name the assembly, and the core's is XLEdge.Addin.Core, not XLEdge.
        /// </summary>
        private static string AsInTheCore(string original) => original.Replace("/XLEdge;component/", "/XLEdge.Addin.Core;component/");

        /// <summary>The csproj item type each kind of ported file must have.</summary>
        private static string ItemType(string file) =>
            file.EndsWith(".xaml", StringComparison.OrdinalIgnoreCase) ? "Page" : Binary.Contains(file) ? "Resource" : "Compile";

        /// <summary>Ported with ledgered changes forced by the AppDomain boundary (Stage 2a rulings): must exist, may differ.</summary>
        internal static readonly string[] Adapted =
        {
            @"Utilities\LogUtility.cs", "XLEdgeAppPaths.cs", "XLEdgeAppState.cs", @"Helpers\XLEdgeRibbonHelper.cs", @"Helpers\ExcelApplicationHelper.cs",
            "AddinModule.cs", "ADXExcelTaskPane1.cs",
        };

        /// <summary>
        /// Stage 2d: AddinModule's handlers live in the core's plain XLEdge.AddinModule class (no ADX there). Each one listed here keeps
        /// 11.1.2's text exactly - signature line to closing brace - except the ADX ribbon-control type in handler signatures.
        /// </summary>
        internal static readonly string[] AddinModuleMethods =
        {
            "void SafeInvokeWpf(System.Action action)", "void SafeLogException(Exception ex, string context = null)",
            "void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)",
            "void TryShowCalendarControl(Excel.Worksheet selectedSheet, Excel.Range selectedRange)",
            "void TryShowSegmentSelectionWindow(Excel.Worksheet selectedSheet, Excel.Range selectedRange)",
            "void RibEdgeOptions_OnClick(", "void RibEdgeDialogBoxLauncher_OnAction(", "void RibEdgeHelp_OnClick(",
            "void UpdateTabLabel(Excel.Worksheet workSheet)", "void RibEdgeIncludeOutputData_OnClick(", "void RibEdgeDebug_OnClick(",
            "void AdxExcelAppEvents1_SheetBeforeDelete(object sender, object sheet)", "void DeleteTimer_Tick(object sender, EventArgs e)",
            "Excel.Worksheet FindSheetWithTableName(Excel.Workbook workbook, string tableName)", "void DeleteHashNamedRanges()",
            "void DeleteNamedCache()", "string SafeRefersTo(Excel.Name nm)", "bool NmSheetExists(string sheetName, Excel.Workbook workbook)",
            "void adxExcelAppEvents1_SheetActivate(object sender, object hostObj)",
            "void ApplyRibbonState(string stateName)", "void ApplyRibbonState(Excel.Workbook workbook)",
            "void RibEdgeRefresh_OnClick(", "void RibEdgeRefreshAll_OnClick(",
            "bool TryResolveInstanceAndChildFlag(Excel.Worksheet sheet, Excel.ListObject tableObj, string tableName, out bool isChild, bool showMismatchMessage = false)",
            "string BuildRefreshParamsPayload(Excel.Workbook workbook, Excel.Worksheet sheet, Excel.ListObject tableObj)",
            "void RibControlSheet_OnClick(", "void adxExcelAppEvents1_SheetFollowHyperlink(object sender, object sheet, object hyperlink)",
            "Task HandleAttachmentDownloadAsync(string screenTip)", "void NavigateToReportDataSheet(Excel.Worksheet dataSheet)",
            // Stage 2e: login and the task pane
            "ADXExcelTaskPane1 GetPaneInstance()", "void LaunchExcelPane()", "void XLEdgeLogin(bool togglePane)",
            "bool NavigateReportsToAddress(string name, string address)", "string NormalizeAddress(string address)",
            "void RibbonInitialize()", "bool loginButtonVisibility()", "void RibLogin_OnClick(", "void RibEdgeAbout_OnClick(",
            "Task LogOffAllTaskPanesAsync(CancellationToken token)", "Task LogoffFromXLEdgeAddin()", "void LogoffFromAddin()",
            "Task LogoffFromAddinAsync()", "void RibEdgeLogout_OnClick(", "void RibEdgeShowHide_OnClick(",
            "void RibEdgeParamRefreshBook_OnClick(", "Task RefreshBookParametersAsync(string runId, ADXExcelTaskPane1 taskPane)",
            "void RibEdgeParamRefresh_OnClick(", "Task RefreshParametersAsync(string runId, ADXExcelTaskPane1 taskPane)",
            "void adxExcelAppEvents1_WorkbookActivate(object sender, object hostObj)",
            // OISR-22777 (11.1.2, 02-Nov-2026): the Broadcast button
            "void RibEdgeBroadcast_OnClick(",
            // Stage 2f: the GLSense sync
            "void InvokedFromGLSense(string eeName", "void NotifyGLSenseOfLogin(string authToken", "void NotifyGLSenseOfLogout()",
        };

        /// <summary>
        /// Stage 2e: 11.1.2's task pane (ADXExcelTaskPane1) is the core's pane content; these of its methods keep 11.1.2's text
        /// (its header replaces the ADX pane members; HookWebViewEvents adds the instance-origin check and is not listed).
        /// </summary>
        internal static readonly string[] PaneMethods =
        {
            "void ADXExcelTaskPane1_HandleCreated(object sender, EventArgs e)", "void ADXExcelTaskPane1_ADXAfterTaskPaneShow(",
            "void HidePaneSafe()", "void RefreshWebViewHeight()", "Task RunOnUIAsync(Action action)", "Task RunOnUIAsync(Func<Task> func)",
            "Task<T> RunOnUIAsync<T>(Func<T> func)", "Task<T> RunOnUIAsync<T>(Func<Task<T>> func)",
            "void SafeFireAndForget(Func<Task> taskFactory, string context)", "Task EnsureWebViewInitializedAsync()",
            "Task<CoreWebView2Environment> GetOrCreateSharedEnvironmentAsync()", "Task<CoreWebView2Environment> CreateSharedEnvironmentAsync()",
            "Task InitializeWebViewInternalAsync()", "Task NavigateToLoginUrlAsync()", "Task<bool> LogoutAsync(string loginUrl, CancellationToken token)",
            "Task ExecuteScriptAsync(string script)", "void CoreWebView2_PermissionRequested(", "void CoreWebView2_ProcessFailed(",
            "void WebCtrl_SourceChanged(", "void HandleExcelHomeSource()", "Task ProcessCookiesAsync()", "void UpdateDialogLauncherState()",
            "Task ProcessBroadcastMessagesAsync()", "Task<string> BroadcastMessageFromApi(", "Task<string> FetchApiResponseAsync(",
            "string FormatBroadcastMessages(string rawResponse)", "string BuildMessageString(List<BroadcastMessage> messages)",
            "void SyncLoginToGLSense()", "void UpdateExcelTabLabel()", "void WebView_CoreWebView2InitializationCompleted(",
            "void WebView_WebResourceRequested(", "void WebView_DocumentTitleChanged(object sender, object e)", "bool IsCellInEditModeAndWarn()",
            "Task<string?> FetchWorkbookRerunIdsAsync()", "Task NavigateToLoginUrlSafeAsync()", "Task NavigateBlankAsync()",
            "Task RefreshLoginNavigationAsync()", "void SetPaneCaption(string caption)", "void ReleaseFocusToExcel()",
        };

        /// <summary>The ADX types a ported method may name differently in the core (it has no Add-in Express).</summary>
        private static string WithoutAdx(string text) => text
            .Replace("IRibbonControl control", "object control")
            .Replace("ADXAfterTaskPaneShowEventArgs e", "EventArgs e")
            .Replace("AddinExpress.XL.ADXExcelTaskPanePosition.Right", "TaskPanePosition.Right");

        /// <summary>The method's text, from its declaration line (8-space indent, an access modifier) to the brace that closes it.</summary>
        internal static string MethodText(string source, string signature)
        {
            string[] lines = source.Split('\n');
            int start = Array.FindIndex(lines, l => l.StartsWith("        ") && !l.StartsWith("         ") && l.Contains(signature) &&
                                                    System.Text.RegularExpressions.Regex.IsMatch(l, @"^\s{8}(public|private|internal|protected)\b"));
            if (start < 0) return null;
            int depth = 0;
            bool opened = false;
            var text = new System.Text.StringBuilder();
            for (int i = start; i < lines.Length; i++)
            {
                text.Append(lines[i]).Append('\n');
                foreach (char ch in lines[i])
                {
                    if (ch == '{') { depth++; opened = true; }
                    else if (ch == '}') depth--;
                }
                if (opened && depth == 0) return text.ToString();
            }
            return null;
        }

        private static void AssertMethodsMatch(string file, string[] signatures)
        {
            string original = Normalized(Path.Combine(Today, file)), core = Normalized(Path.Combine(Core, file));
            var problems = new List<string>();
            foreach (string signature in signatures)
            {
                string before = MethodText(original, signature), after = MethodText(core, WithoutAdx(signature));
                if (before == null) problems.Add("not in 11.1.2: " + signature);
                else if (after == null) problems.Add("not in the core: " + signature);
                else if (WithoutAdx(before) != after) problems.Add("differs from 11.1.2: " + signature);
            }
            Assert.AreEqual(0, problems.Count, file + ": " + string.Join(Environment.NewLine, problems));
        }

        [TestMethod]
        public void AddinModuleMethods_MatchXLEdge11_1_2() => AssertMethodsMatch("AddinModule.cs", AddinModuleMethods);

        [TestMethod]
        public void PaneMethods_MatchXLEdge11_1_2() => AssertMethodsMatch("ADXExcelTaskPane1.cs", PaneMethods);

        private static string Normalized(string path) => File.ReadAllText(path).TrimStart('﻿').Replace("\r\n", "\n");

        [TestMethod]
        public void PortedFiles_AreByteIdenticalToXLEdge11_1_2()
        {
            Assert.IsTrue(Directory.Exists(Today), Today);
            var problems = new List<string>();
            foreach (string file in Verbatim)
            {
                string original = Path.Combine(Today, file), copy = Path.Combine(Core, file);
                if (!File.Exists(original)) problems.Add("no original: " + file);
                else if (!File.Exists(copy)) problems.Add("not ported: " + file);
                else if (AsInTheCore(Normalized(original)) != Normalized(copy)) problems.Add("differs from 11.1.2: " + file);
            }
            foreach (string file in Binary)
            {
                string original = Path.Combine(Today, file), copy = Path.Combine(Core, file);
                if (!File.Exists(copy)) problems.Add("not ported: " + file);
                else if (!File.ReadAllBytes(original).SequenceEqual(File.ReadAllBytes(copy))) problems.Add("differs from 11.1.2: " + file);
            }
            foreach (string file in Adapted.Where(f => !File.Exists(Path.Combine(Core, f))))
                problems.Add("adapted file missing: " + file);
            Assert.AreEqual(0, problems.Count, string.Join(Environment.NewLine, problems));
        }

        [TestMethod]
        public void EveryPortedFile_IsCompiledIntoTheCore()
        {
            string project = File.ReadAllText(Path.Combine(Core, "XLEdge.Addin.Core.csproj"));
            var missing = Verbatim.Concat(Adapted).Concat(Binary).Where(f => !project.Contains($"<{ItemType(f)} Include=\"{f}\"")).ToList();
            Assert.AreEqual(0, missing.Count, "not in XLEdge.Addin.Core.csproj: " + string.Join(", ", missing));
        }
    }
}
