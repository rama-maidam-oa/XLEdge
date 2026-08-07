using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Xml.Linq;
using XLEdge.Helpers;
using XLEdge.Models;
using XLEdge.Utilities;
using XLEdge.Views;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Reusable report creation helper. Provides methods to create reports by parsing a document title
    /// or by using an existing Excel ListObject name. Uses XLApp for Excel access and existing
    /// utilities for API and UI interactions.
    /// </summary>
    public static class ReportGenerator
    {
        private static EdgeRequest _edgeRequest;
        private static bool _showWaitWindow;
        private static CancellationHelper _ctsHelper;
        private static XLEdgeWaitWindow _waitWindow;
        private static AppOverlay _appOverlay;
        private static readonly string _eeLoginUrl = XLEdgeAppState.Instance.LoginUrl;

        private const string ExtraParametersKey = "extraParameters";
        private const string RequestTimedOutMessage = "The request timed out. Please try again.";
        private const string ColumnElementName = "Column";

        private static void GetEdgeRequestFromTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                throw new ArgumentException("Title cannot be null or empty.", nameof(title));
            var parts = title.Split('|');
            if (parts.Length < 3)
                throw new ArgumentException("Title must contain at least three parts separated by '|'.", nameof(title));
            _edgeRequest = EdgeRequestParser.Parse(title);
        }

        private static async Task DisplayErrorAsync(string message)
        {
            if (_showWaitWindow && _waitWindow != null)
            {
                // Matches FormProcessBar.vb's error handling: Edge_CloseProgress() followed by a
                // separate XLEdgeMsgDisplay MessageBox call. SetProcessMessage+RequestClose alone
                // (the previous behavior) closed the wait window immediately after setting its
                // label text, so the message was never actually visible - MessageFunctions.XLEdgeMessage
                // is the C# port of XLEdgeMsgDisplay and shows a real, independent dialog instead.
                await UiDispatcher.RunAsync(() =>
                {
                    _waitWindow.RequestClose();
                    MessageFunctions.XLEdgeMessage(message, System.Windows.Forms.MessageBoxIcon.Error);
                });
            }
            else
            {
                await UiDispatcher.RunAsync(() => _appOverlay?.ShowError(message));
            }
        }

        private static async Task CreateAndShowWaitWindow()
        {
            await UiDispatcher.RunAsync(() =>
            {
                _waitWindow = new XLEdgeWaitWindow(_ctsHelper);
                _waitWindow.SetProcessTitle("Generating report", MahApps.Metro.IconPacks.PackIconFontAwesomeKind.FileExcelSolid);
                _waitWindow.StartMonitoring();
                _waitWindow.Show();
            });
        }

        private static async Task SetMessage(string message)
        {
            if (_showWaitWindow && _waitWindow != null)
            {
                await UiDispatcher.RunAsync(() => _waitWindow.SetProcessMessage(message));
            }
            else
            {
                await ShowBusyOverlayAsync(message);
            }
        }

        private static async Task ShowBusyOverlayAsync(string message)
        {
            await UiDispatcher.RunAsync(() =>
            {
                _appOverlay?.ShowBusyasyn(
                    message: message + " (Click cancel to stop)",
                    cancelAction: async () =>
                    {
                        if (!_ctsHelper.IsCancellationRequested)
                        {
                            _ctsHelper.Cancel();
                            LogUtility.LogWarn($"Operation cancelled by user: {message}");
                        }
                        await Task.Delay(80);
                    }
                );
            });
        }

        public static async Task XLEdgeDataGeneration(List<string> reportDetails, AppOverlay appOverlay = null, bool useWaitWindow = false)
        {
            _showWaitWindow = useWaitWindow;
            _appOverlay = appOverlay;

            if (reportDetails == null || reportDetails.Count == 0)
            {
                await DisplayErrorAsync("Invalid report details provided. Cannot generate report.");
                return;
            }

            if (string.IsNullOrWhiteSpace(_eeLoginUrl))
            {
                await DisplayErrorAsync("Login URL is not set. Please configure the login URL in the settings.");
                return;
            }

            _ctsHelper = new CancellationHelper();

            if (_showWaitWindow)
            {
                await CreateAndShowWaitWindow();
            }
        }

        /// <summary>
        /// Strips extra parameter display values from the payload for CSV endpoint.
        /// Keeps raw values (ORACLE_RESP_ID, ORACLE_GL_SEGMENT_VALUES).
        /// </summary>
        private static string StripExtraParameterDisplayValues(string paramsJson)
        {
            if (string.IsNullOrWhiteSpace(paramsJson))
            {
                return paramsJson;
            }

            try
            {
                using var doc = JsonDocument.Parse(paramsJson);

                // Find extraParameters (case insensitive)
                if (!TryFindExtraParametersElement(doc.RootElement, out JsonElement extraParamsEl, out string extraParamsKey))
                {
                    return paramsJson;
                }

                // Build cleaned extra parameters - ONLY strip display values, keep raw values
                Dictionary<string, object> cleanedExtraParams = BuildCleanedExtraParameters(extraParamsEl);

                // Build cleaned JSON
                return WriteCleanedParamsJson(doc.RootElement, extraParamsKey, cleanedExtraParams);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(StripExtraParameterDisplayValues));
                return paramsJson;
            }
        }

        /// <summary>
        /// Finds the case-insensitive "extraParameters" property on the given JSON object element.
        /// Returns true only when the property was found AND its value is a JSON object.
        /// </summary>
        private static bool TryFindExtraParametersElement(JsonElement root, out JsonElement extraParamsEl, out string extraParamsKey)
        {
            JsonProperty match = root.EnumerateObject()
                .FirstOrDefault(prop => prop.Name.Equals(ExtraParametersKey, StringComparison.OrdinalIgnoreCase));

            extraParamsEl = match.Value;
            extraParamsKey = match.Name;

            if (extraParamsEl.ValueKind == JsonValueKind.Undefined || extraParamsEl.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Builds the cleaned extraParameters dictionary - strips display-only values, keeps raw values.
        /// </summary>
        private static Dictionary<string, object> BuildCleanedExtraParameters(JsonElement extraParamsEl)
        {
            var cleanedExtraParams = new Dictionary<string, object>();
            var displayKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ORACLE_RESP_DISPLAY_VALUE",
                "ORACLE_GL_SEGMENT_DISPLAY_VALUES"
            };

            foreach (JsonProperty prop in extraParamsEl.EnumerateObject().Where(p => !displayKeys.Contains(p.Name)))
            {
                // Keep raw values
                if (prop.Name.Equals("ORACLE_RESP_ID", StringComparison.OrdinalIgnoreCase))
                {
                    // Store as string
                    cleanedExtraParams[prop.Name] = prop.Value.ToString();
                }
                else if (prop.Name.Equals("ORACLE_GL_SEGMENT_VALUES", StringComparison.OrdinalIgnoreCase))
                {
                    // Keep GL segment values as string to preserve format
                    cleanedExtraParams[prop.Name] = prop.Value.ToString();
                }
                else if (prop.Value.ValueKind == JsonValueKind.String)
                {
                    cleanedExtraParams[prop.Name] = prop.Value.GetString();
                }
                else
                {
                    cleanedExtraParams[prop.Name] = prop.Value.ToString();
                }
            }

            return cleanedExtraParams;
        }

        /// <summary>
        /// Re-serializes the root JSON object, substituting the extraParameters property with the cleaned version.
        /// </summary>
        private static string WriteCleanedParamsJson(JsonElement root, string extraParamsKey, Dictionary<string, object> cleanedExtraParams)
        {
            using var stream = new MemoryStream();
            using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false });

            writer.WriteStartObject();

            foreach (JsonProperty prop in root.EnumerateObject())
            {
                if (prop.Name.Equals(extraParamsKey, StringComparison.OrdinalIgnoreCase))
                {
                    writer.WritePropertyName(ExtraParametersKey);
                    writer.WriteStartObject();
                    foreach (var kvp in cleanedExtraParams)
                    {
                        writer.WritePropertyName(kvp.Key);
                        writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                    }
                    writer.WriteEndObject();
                }
                else
                {
                    writer.WritePropertyName(prop.Name);
                    prop.Value.WriteTo(writer);
                }
            }

            writer.WriteEndObject();
            writer.Flush();

            return Encoding.UTF8.GetString(stream.ToArray());
        }

        /// <summary>
        /// Entry point for live ("Edge") report generation from a document title. Wraps the real work
        /// in <see cref="CreateReportFromTitleAsyncCore"/> with a top-level safety net: this method is
        /// invoked fire-and-forget (see XLEdgeCTP.xaml.cs's SafeFireAndForget), which only logs an
        /// exception that escapes it - it never hides the busy overlay or shows an error. Every step
        /// inside the core method that talks to the API or Excel already has its own try/catch that
        /// calls DisplayErrorAsync+CleanupAsync, but a handful of things in between (progress-message/
        /// busy-overlay updates via SetMessage, request-parsing helpers, etc.) were not covered - if
        /// any of those ever threw, the busy spinner that was already showing would be left stuck
        /// spinning forever with no error surfaced, since the fire-and-forget wrapper silently swallows
        /// it. This catch-all guarantees that no matter where an exception originates, the user always
        /// gets an immediate error toast and the busy overlay/wait window is always dismissed.
        /// </summary>
        public static async Task CreateReportFromTitleAsync(string title, AppOverlay appOverlay = null, bool useWaitWindow = false, string paramsJsonPayload = null)
        {
            try
            {
                await CreateReportFromTitleAsyncCore(title, appOverlay, useWaitWindow, paramsJsonPayload);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "CreateReportFromTitleAsync: unhandled exception escaped report generation");
                await DisplayErrorAsync($"An unexpected error occurred during report generation.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
            }
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 33): each network fetch-and-validate
        // step is pulled into its own method returning (ShouldContinue, Value); the core method
        // becomes a linear sequence of "if a step failed, stop" checks. Every step's exact URL
        // construction, message text, and catch-clause ordering (OperationCanceledException before
        // ApiTimeoutException before the generic Exception fallback) is preserved unchanged - only
        // the packaging into methods changed, not the behavior.
        private static async Task CreateReportFromTitleAsyncCore(string title, AppOverlay appOverlay, bool useWaitWindow, string paramsJsonPayload)
        {
            using var excelBulkScope = new ExcelBulkOperationScope();

            _appOverlay = appOverlay;
            _ctsHelper = new CancellationHelper();

            bool isDrilldownRequest = !string.IsNullOrWhiteSpace(paramsJsonPayload);

            // Mirrors VB.NET's AddinModule.vb (FollowDrilldown = True right before Edge_ThreadProgress()
            // in the drilldown hyperlink handler, reset to False for a plain, non-drilldown run). Set
            // fresh on every call rather than relying on a separate "operation completed" reset, so it
            // always reflects THIS invocation regardless of how the previous one finished. This is what
            // RewriteParameterSectionRows reads to write IT1 = "Child Report" on the parameter sheet,
            // which the ribbon Refresh/Refresh All/Run handlers check to block re-running child reports -
            // without this, IT1 stayed empty for every drilldown-generated report and refresh wasn't
            // actually blocked for them.
            XLEdgeAppState.Instance.FollowDrilldown = isDrilldownRequest;

            if (string.IsNullOrWhiteSpace(title))
            {
                await DisplayErrorAsync("Title is empty. Cannot generate report.");
                return;
            }

            _showWaitWindow = useWaitWindow;
            if (_showWaitWindow)
            {
                await CreateAndShowWaitWindow();
            }

            if (!await ParseEdgeRequestOrShowErrorAsync(title))
            {
                return;
            }

            // "Process" titles are submitted/scheduled report runs, not live ad-hoc ("Edge") ones -
            // matches VB.NET's FormProcessBar.vb ReturnHTTP/MetaInfo, which both branch on this same
            // type segment (var(0)/strIDs(0)) to hit an entirely different set of endpoints
            // (/rest/secure/process/... with a processId, GET+Form, no post body) instead of the
            // live-run endpoints (/rest/secure/report/..., runId, POST+JSON). A drilldown always
            // takes the live-run shape regardless of the parent report's type, matching VB's
            // FollowDrilldown check taking priority over the type check in both functions.
            bool isProcessReport = !isDrilldownRequest &&
                string.Equals(_edgeRequest.ReportType, "Process", StringComparison.OrdinalIgnoreCase);

            // Download report data first, matching FormProcessBar.vb's original order (StartTaskHere/
            // ReturnHTTP runs before Edge_GenerateData_Multisheet's MetaInfo/ParamInfo calls) - do not
            // reorder this.
            (bool csvOk, string csvResponse) = await FetchAndPersistCsvResponseAsync(isDrilldownRequest, isProcessReport, paramsJsonPayload);
            if (!csvOk)
            {
                return;
            }

            (bool metaOk, string metaResponse) = await FetchReportMetaResponseAsync(isProcessReport, isDrilldownRequest);
            if (!metaOk)
            {
                return;
            }

            (bool paramsOk, string paramsResponse) = await FetchReportParamsResponseAsync(isDrilldownRequest, paramsJsonPayload);
            if (!paramsOk)
            {
                return;
            }

            (bool metaParsedOk, ReportMeta reportMeta) = await TryDeserializeReportMetaAsync(metaResponse);
            if (!metaParsedOk)
            {
                return;
            }

            if (!await TryBuildReportTableAsync(reportMeta, csvResponse, metaResponse, paramsResponse, title))
            {
                return;
            }

            await CleanupAsync();
        }

        // Extracted from CreateReportFromTitleAsyncCore.
        private static async Task<bool> ParseEdgeRequestOrShowErrorAsync(string title)
        {
            try
            {
                GetEdgeRequestFromTitle(title);
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Failed to parse title for report generation: {title}");
                await DisplayErrorAsync($"Invalid title format for report generation. Title Format {title}");
                return false;
            }
        }

        // Extracted from CreateReportFromTitleAsyncCore - downloads the report's CSV data and writes
        // it to the temporary CSV file used by BuildReportTable.
        private static async Task<(bool ShouldContinue, string CsvResponse)> FetchAndPersistCsvResponseAsync(bool isDrilldownRequest, bool isProcessReport, string paramsJsonPayload)
        {
            string csvResponse;
            try
            {
                await SetMessage("Downloading report data...");
                string csvUrl;
                if (isDrilldownRequest)
                {
                    csvUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/runner?type=csv";
                }
                else if (isProcessReport)
                {
                    csvUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/process/excel-data?processId={_edgeRequest.ReportId}&type=csv";
                }
                else
                {
                    csvUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/runner?runId={_edgeRequest.ReportRunId}&type=csv";
                }

                // Strip display values from the payload before sending it to the CSV endpoint.
                string csvPayload = isDrilldownRequest ? StripExtraParameterDisplayValues(paramsJsonPayload) : null;

                // Process reports fetch via a plain GET with no body (matching VB's ReturnHTTP: Form
                // content type, empty PostData, GET, whenever it's not FollowDrilldown/"Edge"); Edge
                // (live) reports and drilldowns both POST as JSON, same as before.
                csvResponse = isProcessReport
                    ? await ApiHelper.ServerAPI(csvUrl, "Form", string.Empty, "GET", _ctsHelper.GetToken())
                    : await ApiHelper.ServerAPI(csvUrl, "JSON", csvPayload ?? string.Empty, "POST", _ctsHelper.GetToken());
                await Task.Delay(100);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Report generation cancelled by user.");
                await ApiHelper.NotifyCancelRunAsync(XLEdgeAppState.Instance.LoginUrl, _edgeRequest?.ReportRunId);
                await DisplayErrorAsync("Report generation was cancelled by the user.");
                await CleanupAsync();
                return (false, null);
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Report generation request timed out");
                await DisplayErrorAsync(RequestTimedOutMessage);
                await CleanupAsync();
                return (false, null);
            }
            catch (Exception ex)
            {
                // Clean up the wait window/overlay and restore Excel focus on any unhandled error.
                LogUtility.LogException(ex, "Unhandled error in report generation");
                await DisplayErrorAsync($"An unexpected error occurred during report generation.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return (false, null);
            }

            if (string.IsNullOrWhiteSpace(csvResponse))
            {
                LogUtility.LogWarn("CSV response is empty. Cannot generate report.");
                await DisplayErrorAsync("Failed to download report data. The response was empty.");
                await CleanupAsync();
                return (false, null);
            }

            try
            {
                await SetMessage("Writing temporary CSV file...");
                await Task.Run(() => WriteTempCsv(csvResponse, _edgeRequest.ReportRunId));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write temporary CSV file");
                await DisplayErrorAsync($"Failed to write temporary CSV file for report generation.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return (false, null);
            }

            return (true, csvResponse);
        }

        // Extracted from CreateReportFromTitleAsyncCore - fetches the report definition (Meta).
        private static async Task<(bool ShouldContinue, string MetaResponse)> FetchReportMetaResponseAsync(bool isProcessReport, bool isDrilldownRequest)
        {
            // Fetch report definition (Meta) - always need this from API. Process reports use a
            // different endpoint/id shape than live Edge reports or drilldowns - matches VB.NET's
            // MetaInfo (FollowDrilldown always wins and uses the reportId+runId shape; otherwise
            // "PROCESS" uses /rest/secure/process/report-definition?processId=..., everything else
            // keeps the /rest/secure/report/report-definition?reportId=...&runId=... shape).
            await SetMessage("Fetching report definition...");
            string metaUrl;
            if (isProcessReport)
            {
                metaUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/process/report-definition?processId={_edgeRequest.ReportId}&isDrillDown=false";
            }
            else
            {
                string isDrillDownFlag = isDrilldownRequest ? "true" : "false";
                metaUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/report-definition?reportId={_edgeRequest.ReportId}&runId={_edgeRequest.ReportRunId}&isDrillDown={isDrillDownFlag}";
            }
            string metaResponse = string.Empty;

            try
            {
                metaResponse = await ApiHelper.ServerAPI(metaUrl, "Form", "", "GET", _ctsHelper.GetToken());
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Report definition fetch cancelled by user.");
                await ApiHelper.NotifyCancelRunAsync(XLEdgeAppState.Instance.LoginUrl, _edgeRequest?.ReportRunId);
                await DisplayErrorAsync("Report definition fetch was cancelled by the user.");
                await CleanupAsync();
                return (false, null);
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Report definition fetch timed out");
                await DisplayErrorAsync(RequestTimedOutMessage);
                await CleanupAsync();
                return (false, null);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unhandled error fetching report definition");
                await DisplayErrorAsync($"An unexpected error occurred while fetching report definition.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return (false, null);
            }

            if (string.IsNullOrWhiteSpace(metaResponse))
            {
                LogUtility.LogWarn("Report definition response is empty. Cannot generate report.");
                await DisplayErrorAsync("Failed to fetch report definition. The response was empty.");
                await CleanupAsync();
                return (false, null);
            }

            LogResponsePayload("Report definition response (metaResponse)", metaResponse);

            return (true, metaResponse);
        }

        // Extracted from CreateReportFromTitleAsyncCore - fetches the report's parameter display
        // payload (including the round trip required for a drilldown request).
        private static async Task<(bool ShouldContinue, string ParamsResponse)> FetchReportParamsResponseAsync(bool isDrilldownRequest, string paramsJsonPayload)
        {
            // Fetch report parameters. For a drilldown, paramsJsonPayload is the request body built
            // by DrilldownRequestBuilder (reportId/parameters/extraParameters scoped to the clicked
            // row) - it has to be POSTed to this endpoint and the actual response captured, matching
            // FormProcessBar.vb's ParamInfo (which always posts DrillPostData and returns the real
            // server response for FollowDrilldown). It is not itself the params list and must not be
            // reused as-is - that was the bug: drilldowns were skipping this round trip entirely.
            string paramsResponse;
            await SetMessage("Fetching report parameters...");
            string paramsUrl = isDrilldownRequest
                ? $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/params?runId=&type={_edgeRequest.ReportType}"
                : $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/params?runId={_edgeRequest.ReportRunId}&type={_edgeRequest.ReportType}";

            try
            {
                paramsResponse = isDrilldownRequest
                    ? await ApiHelper.ServerAPI(paramsUrl, "JSON", paramsJsonPayload, "POST", _ctsHelper.GetToken())
                    : await ApiHelper.ServerAPI(paramsUrl, "Form", "", "GET", _ctsHelper.GetToken());
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Report parameters fetch cancelled by user.");
                await ApiHelper.NotifyCancelRunAsync(XLEdgeAppState.Instance.LoginUrl, _edgeRequest?.ReportRunId);
                await DisplayErrorAsync("Report parameters fetch was cancelled by the user.");
                await CleanupAsync();
                return (false, null);
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Report parameters fetch timed out");
                await DisplayErrorAsync(RequestTimedOutMessage);
                await CleanupAsync();
                return (false, null);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unhandled error fetching report parameters");
                await DisplayErrorAsync($"An unexpected error occurred while fetching report parameters.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return (false, null);
            }

            if (string.IsNullOrWhiteSpace(paramsResponse))
            {
                LogUtility.LogWarn("Report parameters response is empty. Cannot generate report.");
                await DisplayErrorAsync("Failed to fetch report parameters. The response was empty.");
                await CleanupAsync();
                return (false, null);
            }

            LogResponsePayload("Report parameters response (paramsResponse)", paramsResponse);

            return (true, paramsResponse);
        }

        // Extracted from CreateReportFromTitleAsyncCore.
        private static async Task<(bool ShouldContinue, ReportMeta ReportMeta)> TryDeserializeReportMetaAsync(string metaResponse)
        {
            try
            {
                ReportMeta reportMeta = JsonSerializer.Deserialize<ReportMeta>(metaResponse, JsonGlobals.Options);
                if (reportMeta == null)
                {
                    await DisplayErrorAsync("Report definition could not be parsed.");
                    await CleanupAsync();
                    return (false, null);
                }

                return (true, reportMeta);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to parse report definition JSON");
                await DisplayErrorAsync("Report definition is not in the expected format.");
                await CleanupAsync();
                return (false, null);
            }
        }

        // Extracted from CreateReportFromTitleAsyncCore.
        private static async Task<bool> TryBuildReportTableAsync(ReportMeta reportMeta, string csvResponse, string metaResponse, string paramsResponse, string title)
        {
            try
            {
                await SetMessage("Building report in Excel...");
                BuildReportTable(_edgeRequest, reportMeta, csvResponse, metaResponse, paramsResponse, title);
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to build report table in Excel");
                await DisplayErrorAsync($"Failed to write the report into Excel.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return false;
            }
        }

        /// <summary>
        /// Ported from FormProcessBar.vb's BGWorker_DoWork "MultiData" branch
        /// </summary>
        public static async Task CreateMultiDataReportsAsync(string runIdsRaw, AppOverlay appOverlay = null, bool useWaitWindow = false)
        {
            if (string.IsNullOrWhiteSpace(runIdsRaw))
            {
                LogUtility.LogWarn("CreateMultiDataReportsAsync called with no run ids.");
                return;
            }

            List<string> reportTitles = ProcessRunIds(runIdsRaw);
            if (reportTitles.Count == 0)
            {
                LogUtility.LogWarn($"CreateMultiDataReportsAsync: no reports parsed from '{runIdsRaw}'.");
                return;
            }

            LogUtility.LogDebug($"CreateMultiDataReportsAsync: running {reportTitles.Count} report(s) from '{runIdsRaw}'.");

            foreach (string reportTitle in reportTitles)
            {
                try
                {
                    await CreateReportFromTitleAsync(reportTitle, appOverlay, useWaitWindow);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"CreateMultiDataReportsAsync: failed processing '{reportTitle}'");
                }

                if (_ctsHelper?.IsCancellationRequested == true)
                {
                    LogUtility.LogWarn("CreateMultiDataReportsAsync: batch cancelled by user; stopping remaining reports.");
                    break;
                }
            }
        }

        private static List<string> ProcessRunIds(string input)
        {
            var result = new List<string>();

            if (string.IsNullOrWhiteSpace(input))
            {
                return result;
            }

            string[] parts = input.Split('^');
            foreach (string part in parts)
            {
                string[] subParts = part.Split('|');
                string reportId = subParts.Length > 0 ? subParts[0] : string.Empty;
                string runId = subParts.Length > 1 ? subParts[1] : string.Empty;
                result.Add($"Edge|{reportId}|{runId}|");
            }

            return result;
        }

        /// <summary>
        /// Ported from FormProcessBar.vb's Edge_GenerateLogs/Edge_FillLogs
        /// </summary>
        public static async Task CreateLogsReportAsync(string logsRequestStr, AppOverlay appOverlay = null, bool useWaitWindow = false)
        {
            using var excelBulkScope = new ExcelBulkOperationScope();

            _appOverlay = appOverlay;
            _showWaitWindow = useWaitWindow;
            _ctsHelper = new CancellationHelper();

            if (string.IsNullOrWhiteSpace(logsRequestStr))
            {
                await DisplayErrorAsync("Logs request is empty. Cannot fetch process logs.");
                return;
            }

            string[] parts = logsRequestStr.Split('|');
            string processId = parts.Length > 1 ? parts[1] : string.Empty;

            if (string.IsNullOrWhiteSpace(processId))
            {
                await DisplayErrorAsync($"Invalid logs request format. Title Format {logsRequestStr}");
                return;
            }

            if (string.IsNullOrWhiteSpace(XLEdgeAppState.Instance.LoginUrl))
            {
                await DisplayErrorAsync("Login URL is not set. Please configure the login URL in the settings.");
                return;
            }

            if (_showWaitWindow)
            {
                await CreateAndShowWaitWindow();
            }

            string logText;
            try
            {
                await SetMessage("Downloading process logs...");
                string logsUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/process/excel-log?processId={processId}";
                logText = await ApiHelper.ServerAPI(logsUrl, "Form", "", "GET", _ctsHelper.GetToken());
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Process logs fetch cancelled by user.");
                await DisplayErrorAsync("Fetching process logs was cancelled by the user.");
                await CleanupAsync();
                return;
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Process logs fetch timed out");
                await DisplayErrorAsync(RequestTimedOutMessage);
                await CleanupAsync();
                return;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unhandled error fetching process logs");
                await DisplayErrorAsync($"An unexpected error occurred while fetching process logs.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return;
            }

            if (string.IsNullOrWhiteSpace(logText))
            {
                LogUtility.LogWarn("Process logs response is empty.");
                await DisplayErrorAsync("Failed to fetch process logs. The response was empty.");
                await CleanupAsync();
                return;
            }

            try
            {
                await SetMessage("Writing logs into Excel...");
                BuildLogsSheet(processId, logText);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write process logs into Excel");
                await DisplayErrorAsync($"Failed to write process logs into Excel.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return;
            }

            await CleanupAsync();
        }

        private static void BuildLogsSheet(string processId, string logText)
        {
            Excel.Application excelApp = XLApp.App;
            if (excelApp == null)
            {
                throw new InvalidOperationException("Excel application instance is not available.");
            }

            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("No active workbook.");
            }

            string sheetName = $"Logs_{processId}";

            Excel.Worksheet logsSheet;
            if (ExcelSheetHelper.SheetExists(sheetName, workbook))
            {
                logsSheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                logsSheet.Cells.Clear();
            }
            else
            {
                try
                {
                    logsSheet = (Excel.Worksheet)workbook.Worksheets.Add();
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(BuildLogsSheet)}: default Worksheets.Add() failed, falling back to append-at-end - {ex.Message}");
                    logsSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
                }
                logsSheet.Name = sheetName;
            }

            logsSheet.Activate();

            string tempFile = Path.Combine(XLEdgeAppPaths.TempFolder, $"{processId}_Logs.txt");
            try
            {
                Directory.CreateDirectory(XLEdgeAppPaths.TempFolder);
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
                File.WriteAllText(tempFile, logText);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error writing logs temp file: {tempFile}");
                throw;
            }

            Excel.QueryTable queryTable = null;
            try
            {
                queryTable = logsSheet.QueryTables.Add($"TEXT;{tempFile}", (Excel.Range)logsSheet.Cells[1, 1]);
                queryTable.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                queryTable.TextFilePlatform = 65001;
                queryTable.TextFilePromptOnRefresh = false;
                queryTable.Refresh(false);
            }
            finally
            {
                if (queryTable != null)
                {
                    try
                    {
                        queryTable.SaveData = false;
                        queryTable.Delete();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "Failed to delete QueryTable definition after importing process logs");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(queryTable);
                    }
                }
            }

            try
            {
                excelApp.ActiveWindow.DisplayGridlines = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to set DisplayGridlines=false on logs sheet");
            }
        }

        public static async Task DownloadFile1Async(string downloadUrl)
        {
            if (string.IsNullOrWhiteSpace(downloadUrl))
            {
                LogUtility.LogWarn("DownloadFile1Async called with an empty URL.");
                return;
            }

            try
            {
                string savedPath = await ApiHelper.DownloadFileAsync(downloadUrl);
                if (string.IsNullOrWhiteSpace(savedPath))
                {
                    MessageFunctions.XLEdgeMessage("Failed to download the file.", System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxButtons.OK);
                    return;
                }

                string fileName = Path.GetFileName(savedPath);
                MessageFunctions.XLEdgeMessage(
                    $"Attachment has been saved to the downloads folder and the file name is \"{fileName}\"",
                    System.Windows.Forms.MessageBoxIcon.Information, System.Windows.Forms.MessageBoxButtons.OK);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "DownloadFile1Async failed");
                MessageFunctions.XLEdgeMessage($"Failed to download the file.{Environment.NewLine}{ex.Message}", System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxButtons.OK);
            }
        }

        private static void LogResponsePayload(string context, string response)
        {
            if (!LogUtility.DebugMode)
            {
                return;
            }

            if (XLEdgeAppState.Instance.DebugOutputData)
            {
                LogUtility.LogDebug($"{context}: {response}");
            }
            else
            {
                LogUtility.LogDebug($"{context}: <{response?.Length ?? 0} character(s) - enable 'Include Output Data' to log the full payload>");
            }
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 85): the original single method is
        // decomposed below into small, single-purpose private helpers. Every line of logic is
        // unchanged - this only changes how the logic is packaged into methods, not what it does or
        // the order in which it runs, to avoid introducing any behavioral regression.
        private static void BuildReportTable(EdgeRequest request, ReportMeta reportMeta, string csvResponse, string metaJson, string paramsJson, string title)
        {
            Excel.Application excelApp = XLApp.App;
            if (excelApp == null)
            {
                throw new InvalidOperationException("Excel application instance is not available.");
            }

            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            if (workbook == null)
            {
                throw new InvalidOperationException("No active workbook.");
            }

            (string tableId, List<List<string>> rows, int dataRowCount, List<(string Original, string Modified, int RawIndex)> mappings) =
                BuildColumnMappings(request, reportMeta, csvResponse);

            bool sameSheet = XLEdgeAppState.Instance.ParamDataSameSheet;
            int headerRow = sameSheet ? 8 : 1;
            int dataStartRow = headerRow + 1;

            Excel.Worksheet sheet = ResolveOrCreateReportSheet(workbook, tableId, reportMeta, sameSheet, headerRow, out string companionSheetToDelete);

            ActivateAndUnfreezeSheet(excelApp, sheet);

            Excel.ListObject listObject = WriteReportDataAndCreateTable(sheet, tableId, headerRow, dataStartRow, mappings, reportMeta, rows);

            HideFlaggedColumns(listObject, reportMeta, mappings);

            string reportTitleText = ComputeReportTitleText(reportMeta, request);

            WriteReportParameterSection(workbook, sheet, sameSheet, reportTitleText, paramsJson, dataRowCount, tableId);

            AddDrilldownHyperlinks(sheet, listObject, reportMeta);
            AddAttachmentAndImageColumns(sheet, listObject, reportMeta);

            PersistReportMetadata(workbook, title, tableId, metaJson, paramsJson, mappings);

            ApplyReportTableStyling(listObject);

            ApplyColumnFreeze(excelApp, reportMeta, mappings.Count);

            DeleteOrphanedCompanionSheet(workbook, companionSheetToDelete);
        }

        // Extracted from BuildReportTable - builds the Excel table identifier, parses the raw CSV
        // response, and produces the ordered list of (original, sanitized, raw-column-index)
        // mappings used to write the header/data.
        private static (string TableId, List<List<string>> Rows, int DataRowCount, List<(string Original, string Modified, int RawIndex)> Mappings) BuildColumnMappings(
            EdgeRequest request, ReportMeta reportMeta, string csvResponse)
        {
            // Matches VB.NET's FormProcessBar.vb EETableID assignment: a submitted/scheduled
            // ("Process") report's table is suffixed "_P" instead of "_E". AddinModule.cs's
            // UpdateTabLabel/XLEdgeRibbonHelper.ProcessActiveWorkbook already recognize "_P" tables
            // (showing "This sheet has a scheduled output." and disabling Refresh/Param Refresh for
            // them) - that logic was already ported and correct, it just never actually fired because
            // this method always produced "_E" regardless of report type.
            bool isProcessTable = string.Equals(request.ReportType, "Process", StringComparison.OrdinalIgnoreCase);
            string tableId = $"ORB_{request.ReportId}_{request.ReportRunId}_{(isProcessTable ? "P" : "E")}";

            List<List<string>> rows = ParseCsv(csvResponse).ToList();
            List<string> rawHeader = rows.Count > 0 ? rows[0] : new List<string>();
            int rawCols = rawHeader.Count;
            int dataRowCount = Math.Max(0, rows.Count - 1);

            var mappings = new List<(string Original, string Modified, int RawIndex)>();
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (reportMeta.Columns != null)
            {
                foreach (RptColumn col in reportMeta.Columns)
                {
                    int rawIndex = rawHeader.FindIndex(h => string.Equals(h?.Trim(), col.Name?.Trim(), StringComparison.OrdinalIgnoreCase)) + 1;
                    string safeName = MakeUniqueName(col.Name, usedNames);
                    mappings.Add((col.Name, safeName, rawIndex));
                }
            }

            for (int i = 1; i <= rawCols; i++)
            {
                if (mappings.Any(m => m.RawIndex == i))
                {
                    continue;
                }

                string orig = rawHeader[i - 1] ?? string.Empty;
                string safeName = MakeUniqueName(string.IsNullOrWhiteSpace(orig) ? $"Column{i}" : orig, usedNames);
                mappings.Add((orig, safeName, i));
            }

            if (mappings.Count == 0)
            {
                throw new InvalidOperationException("Report has no columns to write.");
            }

            return (tableId, rows, dataRowCount, mappings);
        }

        // Extracted from BuildReportTable - finds/prepares the worksheet to write this report's
        // table into, and tracks the name of any now-orphaned companion parameter sheet that should
        // be deleted once the new table/banner has been written.
        private static Excel.Worksheet ResolveOrCreateReportSheet(Excel.Workbook workbook, string tableId, ReportMeta reportMeta, bool sameSheet, int headerRow, out string companionSheetToDelete)
        {
            Excel.Worksheet sheet = FindSheetWithTable(workbook, tableId);

            if (sheet != null)
            {
                PrepareExistingReportSheet(sheet, tableId, headerRow, out companionSheetToDelete);
            }
            else
            {
                sheet = CreateOrReuseReportSheet(workbook, reportMeta);
                companionSheetToDelete = null;
            }

            if (sameSheet && string.IsNullOrEmpty(companionSheetToDelete))
            {
                companionSheetToDelete = FindOrphanedCompanionSheetName(sheet.Name, tableId);
            }

            return sheet;
        }

        // Extracted from BuildReportTable - when a table with this tableId already exists on a
        // sheet, removes the old table and, if the header-row layout is changing (same-sheet banner
        // added/removed), reconciles the banner/companion-sheet state to match the new layout.
        private static void PrepareExistingReportSheet(Excel.Worksheet sheet, string tableId, int headerRow, out string companionSheetToDelete)
        {
            companionSheetToDelete = null;
            int? oldHeaderRow = null;
            try
            {
                Excel.ListObject existing = sheet.ListObjects[tableId];
                oldHeaderRow = existing.HeaderRowRange.Row;
                existing.Delete();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to remove existing table before rebuilding");
            }

            if (oldHeaderRow == 8 && headerRow == 1)
            {
                RemoveSameSheetBanner(sheet);
            }
            else if (oldHeaderRow == 1 && headerRow == 8)
            {
                companionSheetToDelete = TransitionSheetToSameSheetMode(sheet, tableId);
            }
        }

        // Extracted from BuildReportTable - handles the "sheet is switching into same-sheet mode"
        // case: resolves the now-orphaned companion parameter sheet's name (if any) and makes room
        // for the in-sheet banner.
        private static string TransitionSheetToSameSheetMode(Excel.Worksheet sheet, string tableId)
        {
            string companionSheetToDelete = null;
            try
            {
                Excel.Worksheet oldParamSheet = ExcelSheetHelper.GetParameterSheet($"P_{sheet.Name}", tableId);
                if (oldParamSheet != null)
                {
                    companionSheetToDelete = oldParamSheet.Name;
                    Marshal.ReleaseComObject(oldParamSheet);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to resolve old companion parameter sheet before switching to same-sheet mode");
            }

            InsertRoomForSameSheetBanner(sheet);
            return companionSheetToDelete;
        }

        // Extracted from BuildReportTable - resolves the worksheet to use when no existing table
        // with this tableId was found: reuse the report's named sheet if it already exists (clearing
        // it first), otherwise create a brand-new sheet.
        private static Excel.Worksheet CreateOrReuseReportSheet(Excel.Workbook workbook, ReportMeta reportMeta)
        {
            string sheetName = BuildSheetName(reportMeta);
            if (ExcelSheetHelper.SheetExists(sheetName, workbook))
            {
                Excel.Worksheet existingSheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                existingSheet.Cells.Clear();
                ResetLeftoverRowArtifacts(existingSheet);
                return existingSheet;
            }

            Excel.Worksheet newSheet;
            try
            {
                newSheet = (Excel.Worksheet)workbook.Worksheets.Add();
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(CreateOrReuseReportSheet)}: default Worksheets.Add() failed, falling back to append-at-end - {ex.Message}");
                newSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
            }

            newSheet.Name = sheetName;
            return newSheet;
        }

        // Extracted from BuildReportTable - when writing into same-sheet mode, checks for a
        // leftover companion parameter sheet (from a prior non-same-sheet run of this report) that
        // is now orphaned and should be cleaned up.
        private static string FindOrphanedCompanionSheetName(string sheetName, string tableId)
        {
            try
            {
                string paramSheetName = $"P_{sheetName}";
                if (paramSheetName.Length >= 29)
                {
                    paramSheetName = paramSheetName.Substring(0, 28);
                }

                Excel.Worksheet oldParamSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableId);
                if (oldParamSheet != null)
                {
                    string orphanedName = oldParamSheet.Name;
                    Marshal.ReleaseComObject(oldParamSheet);
                    return orphanedName;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to resolve orphaned companion parameter sheet before same-sheet write");
            }

            return null;
        }

        // Extracted from BuildReportTable - activates the target sheet and unfreezes panes so the
        // header/data write always lands starting at row/column 1 regardless of the previous
        // report's frozen-pane state.
        private static void ActivateAndUnfreezeSheet(Excel.Application excelApp, Excel.Worksheet sheet)
        {
            sheet.Activate();

            //unfreezing the columns and rows if they are frozen
            try
            {
                excelApp.ActiveWindow.FreezePanes = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to unfreeze panes on report sheet");
            }
        }

        // Extracted from BuildReportTable - writes the header row and (if any) the data rows as bulk
        // Value2 array writes, then wraps the written range in a new ListObject named tableId.
        private static Excel.ListObject WriteReportDataAndCreateTable(Excel.Worksheet sheet, string tableId, int headerRow, int dataStartRow, List<(string Original, string Modified, int RawIndex)> mappings, ReportMeta reportMeta, List<List<string>> rows)
        {
            int dataRowCount = Math.Max(0, rows.Count - 1);

            object[,] headerArr = new object[1, mappings.Count];
            for (int c = 0; c < mappings.Count; c++)
            {
                headerArr[0, c] = mappings[c].Modified;
            }

            Excel.Range headerWriteRange = ((Excel.Range)sheet.Cells[headerRow, 1]).Resize[1, mappings.Count];
            headerWriteRange.Value2 = headerArr;

            int rowsToReserve = Math.Max(1, dataRowCount);

            if (dataRowCount > 0)
            {
                object[,] writeArr = BuildDataWriteArray(rows, mappings, reportMeta, dataRowCount);

                Excel.Range startCell = (Excel.Range)sheet.Cells[dataStartRow, 1];
                Excel.Range writeRange = startCell.Resize[dataRowCount, mappings.Count];
                writeRange.Value2 = writeArr;
            }

            Excel.Range tableRange = sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow + rowsToReserve, mappings.Count]];
            Excel.ListObject listObject = sheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
            listObject.Name = tableId;
            listObject.TableStyle = "TableStyleLight9";

            return listObject;
        }

        // Extracted from BuildReportTable - builds the 2D data array (row-major) to bulk-write into
        // the sheet, applying each column's configured DataType formatting via
        // XLEdgeValueFormatter.FormatValue.
        private static object[,] BuildDataWriteArray(List<List<string>> rows, List<(string Original, string Modified, int RawIndex)> mappings, ReportMeta reportMeta, int dataRowCount)
        {
            object[,] writeArr = new object[dataRowCount, mappings.Count];

            for (int c = 0; c < mappings.Count; c++)
            {
                int rawIndex = mappings[c].RawIndex;
                string colType = reportMeta.Columns?
                    .FirstOrDefault(rc => string.Equals(rc.Name, mappings[c].Original, StringComparison.OrdinalIgnoreCase))?
                    .DataType;

                for (int r = 0; r < dataRowCount; r++)
                {
                    List<string> rowVals = rows[r + 1];
                    object raw = (rawIndex >= 1 && rawIndex <= rowVals.Count) ? rowVals[rawIndex - 1] : string.Empty;
                    writeArr[r, c] = string.IsNullOrEmpty(colType) ? raw : (XLEdgeValueFormatter.FormatValue(raw, colType) ?? string.Empty);
                }
            }

            return writeArr;
        }

        // Extracted from BuildReportTable - hides any table column whose report-metadata column is
        // flagged Properties.Hidden.
        private static void HideFlaggedColumns(Excel.ListObject listObject, ReportMeta reportMeta, List<(string Original, string Modified, int RawIndex)> mappings)
        {
            foreach (RptColumn col in reportMeta.Columns ?? Array.Empty<RptColumn>())
            {
                if (col.Properties?.Hidden != true)
                {
                    continue;
                }

                var mapping = mappings.FirstOrDefault(m => string.Equals(m.Original, col.Name, StringComparison.OrdinalIgnoreCase));
                if (mapping.Modified == null)
                {
                    continue;
                }

                try
                {
                    Excel.ListColumn listColumn = listObject.ListColumns[mapping.Modified];
                    (listColumn.Range).EntireColumn.Hidden = true;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Failed to hide column '{mapping.Modified}'");
                }
            }
        }

        // Extracted from BuildReportTable - resolves the display title used for the same-sheet
        // banner / companion parameter sheet header (drilldown child label takes priority when set).
        private static string ComputeReportTitleText(ReportMeta reportMeta, EdgeRequest request)
        {
            return XLEdgeValueFormatter.RemoveEquaSymbol(
                (XLEdgeAppState.Instance.FollowDrilldown && !string.IsNullOrWhiteSpace(XLEdgeAppState.Instance.ChildRptLabel))
                    ? XLEdgeAppState.Instance.ChildRptLabel
                    : (reportMeta.Name ?? request.ReportName));
        }

        // Extracted from BuildReportTable - writes the report's parameter display, either as the
        // in-sheet banner (same-sheet mode) or the separate companion parameter sheet.
        private static void WriteReportParameterSection(Excel.Workbook workbook, Excel.Worksheet sheet, bool sameSheet, string reportTitleText, string paramsJson, int dataRowCount, string tableId)
        {
            if (sameSheet)
            {
                try
                {
                    WriteSameSheetBanner(sheet, reportTitleText, paramsJson, dataRowCount, tableId);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to write same-sheet parameter banner");
                }
            }
            else
            {
                try
                {
                    BuildCompanionParameterSheet(workbook, sheet, reportTitleText, paramsJson, tableId, dataRowCount);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to build companion parameter sheet");
                }
            }
        }

        // Extracted from BuildReportTable - persists the report's custom XML metadata part.
        private static void PersistReportMetadata(Excel.Workbook workbook, string title, string tableId, string metaJson, string paramsJson, List<(string Original, string Modified, int RawIndex)> mappings)
        {
            try
            {
                string xml = BuildCustomXml(title, tableId, metaJson, paramsJson, mappings);
                SaveCustomXmlPart(workbook, xml, tableId);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to persist report metadata");
            }
        }

        // Extracted from BuildReportTable - cosmetic-only column autofit/font sizing; safe to ignore
        // on failure.
        private static void ApplyReportTableStyling(Excel.ListObject listObject)
        {
            try
            {
                Excel.Range styleRange = listObject.Range;
                styleRange.Columns.AutoFit();
                styleRange.Font.Size = 9;
            }
            catch (Exception)
            {
                // Cosmetic-only (column width/font size); safe to ignore if it fails.
            }
        }

        // Extracted from BuildReportTable - freezes panes at the metadata-configured locked-column
        // boundary, if any.
        private static void ApplyColumnFreeze(Excel.Application excelApp, ReportMeta reportMeta, int mappingCount)
        {
            //Attempting to freeae the columns based on metadata settings
            try
            {
                int columnLockCount = reportMeta.LockedColumnsCount;
                if (columnLockCount > 0 && columnLockCount < mappingCount)
                {
                    excelApp.ActiveWindow.SplitColumn = columnLockCount;
                    excelApp.ActiveWindow.SplitRow = 0;
                    excelApp.ActiveWindow.FreezePanes = true;
                }

            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to freeze panes on report sheet");
            }
        }

        // Extracted from BuildReportTable - deletes the now-orphaned companion parameter sheet left
        // behind when a report table switched into/out of same-sheet mode.
        private static void DeleteOrphanedCompanionSheet(Excel.Workbook workbook, string companionSheetToDelete)
        {
            if (string.IsNullOrEmpty(companionSheetToDelete))
            {
                return;
            }

            try
            {
                if (ExcelSheetHelper.SheetExists(companionSheetToDelete, workbook))
                {
                    ((Excel.Worksheet)workbook.Worksheets[companionSheetToDelete]).Delete();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to delete orphaned companion parameter sheet after switching to same-sheet mode");
            }
        }

        private static void RemoveSameSheetBanner(Excel.Worksheet sheet)
        {
            try
            {
                Excel.Range entireRow = sheet.Range["A3"].EntireRow;
                int outlineLevel = Convert.ToInt32(entireRow.OutlineLevel);
                if (outlineLevel > 1)
                {
                    sheet.Range["A3:A6"].Rows.Ungroup();
                }

                sheet.Range["A1:A7"].RowHeight = 15;
            }
            catch (Exception)
            {
                // Cosmetic-only (row grouping/height); safe to ignore if it fails.
            }

            try
            {
                sheet.Range["1:7"].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to remove the old same-sheet banner rows before switching to separate-sheet mode");
            }
        }

        private static void ResetLeftoverRowArtifacts(Excel.Worksheet sheet)
        {
            try
            {
                sheet.Cells.ClearOutline();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to clear leftover row outline/grouping before rewriting a data-only sheet");
            }

            try
            {
                sheet.Rows.RowHeight = sheet.StandardHeight;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to reset leftover row heights before rewriting a data-only sheet");
            }
        }

        private static void InsertRoomForSameSheetBanner(Excel.Worksheet sheet)
        {
            try
            {
                sheet.Range["1:7"].Insert(Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to insert banner rows before switching to same-sheet mode");
            }
        }

        private static void WriteSameSheetBanner(Excel.Worksheet sheet, string reportTitle, string paramsJson, int dataRowCount, string tableId)
        {
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                excelApp.ActiveWindow.DisplayGridlines = false;
                sheet.Outline.SummaryRow = Excel.XlSummaryRow.xlSummaryAbove;
                sheet.Outline.AutomaticStyles = true;
                excelApp.ActiveWindow.DisplayOutline = true;
            }
            catch (Exception)
            {
                // Cosmetic-only (gridlines/outline display); safe to ignore if it fails.
            }

            try
            {
                Excel.Range entireRow = sheet.Range["A3"].EntireRow;
                int outlineLevel = Convert.ToInt32(entireRow.OutlineLevel);
                if (outlineLevel > 1)
                {
                    sheet.Range["A3:A6"].Rows.Ungroup();
                }

                sheet.Range["A3:A6"].Rows.Group();
            }
            catch (Exception)
            {
                // Cosmetic-only (row grouping); safe to ignore if it fails.
            }

            try
            {
                sheet.Range["A3"].RowHeight = 5;
                sheet.Range["A7"].RowHeight = 5;
            }
            catch (Exception)
            {
                // Cosmetic-only (spacer row heights); safe to ignore if it fails.
            }

            Excel.Range titleRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 5]];
            titleRange.Merge();
            titleRange.Value2 = reportTitle;
            titleRange.Font.Bold = true;
            titleRange.Font.Italic = true;
            titleRange.Font.Size = 10;
            titleRange.Font.ColorIndex = 2;
            titleRange.Interior.Color = Rgb(21, 96, 130);

            WriteRunInfoStrip(sheet, dataRowCount);

            Excel.Range sectionRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, 5]];
            sectionRange.Merge();
            sectionRange.Value2 = "Parameters Section:";
            sectionRange.Font.Bold = true;
            sectionRange.Font.Italic = true;
            sectionRange.Font.Size = 10;
            sectionRange.Font.ColorIndex = 2;
            sectionRange.Interior.Color = Rgb(241, 169, 131);

            // Writes the parameter label/value rows and bookkeeping cells; shared with the refresh path.
            RewriteParameterSectionRows(sheet, paramsJson, tableId, sameSheetMode: true);

            try
            {
                ExcelApplicationHelper.RequireActiveExcelApplication().Goto(sheet.Range["A1"], true);
                sheet.Outline.ShowLevels(1, 1);
            }
            catch (Exception)
            {
                // Cosmetic-only (navigation/outline levels); safe to ignore if it fails.
            }
        }

        /// <summary>
        /// Writes (or rewrites) the "Parameters Section:" label/value rows plus the IT1/IT2/IT4/IU4/
        /// IV4/IW4/IT5 bookkeeping cells onto <paramref name="paramSheet"/> - used by both report creation
        /// and report refresh. Always clears the target rows/columns first, so a refresh with fewer
        /// parameters than the prior run doesn't leave stale label/value pairs behind.
        /// </summary>
        // Cognitive-complexity refactor (SonarQube S3776, was 29): decomposed into single-purpose
        // helpers, one per parameter-sheet cell group. Every clear/write, comment, and error message
        // is unchanged.
        private static void RewriteParameterSectionRows(Excel.Worksheet paramSheet, string paramsJson, string tableId, bool sameSheetMode)
        {
            List<(string Label, string ValueText)> paramRows = ParseParamDisplayRows(
                paramsJson, out string oracleRespId, out string oracleRespValue, out string segmentValues, out string segmentDisplayValues);

            if (sameSheetMode)
            {
                WriteSameSheetParamRows(paramSheet, paramRows, tableId);
            }
            else
            {
                WriteCompanionParamRows(paramSheet, paramRows, tableId);
            }

            WriteParameterBookkeepingCells(paramSheet, tableId, sameSheetMode);
            WriteOracleResponsibilityCells(paramSheet, oracleRespId, oracleRespValue);
            WriteSegmentValueCell(paramSheet, segmentValues);
            WriteSegmentDisplayValueCell(paramSheet, segmentDisplayValues);
        }

        // Extracted from RewriteParameterSectionRows - writes the multi-column, row-wrapping
        // (rows 4-6, then next column pair) same-sheet banner parameter grid.
        private static void WriteSameSheetParamRows(Excel.Worksheet paramSheet, List<(string Label, string ValueText)> paramRows, string tableId)
        {
            try
            {
                // Generously wide/tall clear (rows 4-6, columns A through BZ) so a refresh that now
                // has fewer parameters than the previous run doesn't leave old label/value pairs
                // behind in columns/rows the new, shorter list no longer reaches.
                paramSheet.Range[paramSheet.Cells[4, 1], paramSheet.Cells[6, 78]].Clear();
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(WriteSameSheetParamRows)}: failed to clear stale same-sheet parameter rows before rewrite - {ex.Message}");
            }

            int irow = 4;
            int icol = 1;
            foreach ((string Label, string ValueText) param in paramRows)
            {
                if (irow > 6)
                {
                    irow = 4;
                    icol += 2;
                }

                try
                {
                    WriteParamLabelCell((Excel.Range)paramSheet.Cells[irow, icol], param.Label);
                    WriteParamValueCell((Excel.Range)paramSheet.Cells[irow, icol + 1], param.ValueText, tableId);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Failed to write same-sheet parameter row for '{param.Label}'");
                }

                irow++;
            }
        }

        // Extracted from RewriteParameterSectionRows - writes the single-column, downward-growing
        // companion-sheet parameter list.
        private static void WriteCompanionParamRows(Excel.Worksheet paramSheet, List<(string Label, string ValueText)> paramRows, string tableId)
        {
            try
            {
                // Companion parameter sheet rows grow downward with no fixed cap (row 3, 4, 5, ...) -
                // clear a generously tall range so a refresh with fewer parameters doesn't leave
                // stale rows below the new, shorter list.
                paramSheet.Range[paramSheet.Cells[3, 1], paramSheet.Cells[300, 2]].Clear();
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(WriteCompanionParamRows)}: failed to clear stale companion-sheet parameter rows before rewrite - {ex.Message}");
            }

            int row = 3;
            foreach ((string Label, string ValueText) param in paramRows)
            {
                try
                {
                    WriteParamLabelCell((Excel.Range)paramSheet.Cells[row, 1], param.Label);
                    WriteParamValueCell((Excel.Range)paramSheet.Cells[row, 2], param.ValueText, tableId);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"Failed to write companion parameter row for '{param.Label}'");
                }

                row++;
            }
        }

        // Extracted from RewriteParameterSectionRows - writes the IT1/IT2/IT5 bookkeeping cells.
        private static void WriteParameterBookkeepingCells(Excel.Worksheet paramSheet, string tableId, bool sameSheetMode)
        {
            try
            {
                paramSheet.Range["IT1"].Value2 = XLEdgeAppState.Instance.FollowDrilldown ? "Child Report" : string.Empty;
                paramSheet.Range["IT5"].Value2 = XLEdgeValueFormatter.RemoveEquaSymbol(XLEdgeAppState.Instance.LoginUrl);

                if (!sameSheetMode)
                {
                    // Same-sheet mode never writes IT2 (see UpdateTabLabel's own comment on why - only
                    // separate-sheet mode's companion sheet needs this marker, to let
                    // ExcelSheetHelper.GetParameterSheet validate a found sheet actually belongs to this
                    // table).
                    paramSheet.Range["IT2"].Value2 = tableId;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write parameter sheet IT1/IT2/IT5 bookkeeping cells");
            }
        }

        // Extracted from RewriteParameterSectionRows - IT4/IU4 are cleared first, then only
        // re-populated if there's an actual value - ensures a blank value this round leaves an
        // actually-blank cell rather than stale content.
        private static void WriteOracleResponsibilityCells(Excel.Worksheet paramSheet, string oracleRespId, string oracleRespValue)
        {
            try
            {
                paramSheet.Range["IT4"].Clear();
                paramSheet.Range["IU4"].Clear();

                if (!string.IsNullOrWhiteSpace(oracleRespId) && !string.IsNullOrWhiteSpace(oracleRespValue))
                {
                    paramSheet.Range["IT4"].Value2 = XLEdgeValueFormatter.RemoveEquaSymbol(oracleRespId);
                    paramSheet.Range["IT4"].WrapText = false;

                    paramSheet.Range["IU4"].Value2 = XLEdgeValueFormatter.RemoveEquaSymbol(oracleRespValue);
                    paramSheet.Range["IU4"].WrapText = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write parameter sheet IT4/IU4 responsibility cells");
            }
        }

        // Extracted from RewriteParameterSectionRows.
        private static void WriteSegmentValueCell(Excel.Worksheet paramSheet, string segmentValues)
        {
            try
            {
                paramSheet.Range["IV4"].Clear();

                if (!string.IsNullOrWhiteSpace(segmentValues))
                {
                    paramSheet.Range["IV4"].Value2 = XLEdgeValueFormatter.RemoveEquaSymbol(segmentValues);
                    paramSheet.Range["IV4"].WrapText = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write parameter sheet IV4 segment values cell");
            }
        }

        // Extracted from RewriteParameterSectionRows - IW4 holds the segment display value,
        // alongside IV4's raw segment value.
        private static void WriteSegmentDisplayValueCell(Excel.Worksheet paramSheet, string segmentDisplayValues)
        {
            try
            {
                paramSheet.Range["IW4"].Clear();

                if (!string.IsNullOrWhiteSpace(segmentDisplayValues))
                {
                    paramSheet.Range["IW4"].Value2 = XLEdgeValueFormatter.RemoveEquaSymbol(segmentDisplayValues);
                    paramSheet.Range["IW4"].WrapText = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to write parameter sheet IW4 segment display values cell");
            }
        }

        private static void WriteRunInfoStrip(Excel.Worksheet sheet, int dataRowCount)
        {
            void Label(Excel.Range cell, string text)
            {
                cell.Value2 = text;
                cell.Font.Bold = true;
                cell.Font.Italic = true;
                cell.Font.Size = 10;
                cell.Font.ColorIndex = 2;
                cell.Interior.Color = Rgb(241, 169, 131);
                try { cell.EntireColumn.AutoFit(); }
                catch (Exception)
                {
                    // Cosmetic-only (column width); safe to ignore if it fails.
                }
            }

            void Value(Excel.Range cell, object val, string numberFormat = null)
            {
                if (numberFormat != null)
                {
                    cell.NumberFormat = numberFormat;
                }

                cell.Value2 = val;
                cell.Font.Bold = false;
                cell.Font.Italic = true;
                cell.Font.Size = 10;
                cell.Font.Color = Rgb(21, 96, 130);
                try { cell.EntireColumn.AutoFit(); }
                catch (Exception)
                {
                    // Cosmetic-only (column width); safe to ignore if it fails.
                }
            }

            Label((Excel.Range)sheet.Cells[1, 7], "Run Date : ");
            Value((Excel.Range)sheet.Cells[1, 8], DateTime.Now, "dd-mmm-yyyy hh:mm:ss");

            Label((Excel.Range)sheet.Cells[1, 9], "Time Zone : ");
            Value((Excel.Range)sheet.Cells[1, 10], XLEdgeValueFormatter.PrintTimeZone(DateTime.Now));

            Label((Excel.Range)sheet.Cells[1, 11], "Executed in : ");
            Value((Excel.Range)sheet.Cells[1, 12], XLEdgeValueFormatter.RemoveEquaSymbol(XLEdgeAppState.Instance.LoginUrl));

            Label((Excel.Range)sheet.Cells[2, 11], "Record Count : ");
            Excel.Range recordCountCell = (Excel.Range)sheet.Cells[2, 12];
            Value(recordCountCell, dataRowCount > 0 ? (object)dataRowCount : null);
            recordCountCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }

        private static void WriteParamLabelCell(Excel.Range cell, string label)
        {
            string text = label ?? string.Empty;
            if (text.Length > 28)
            {
                text = text.Substring(0, 28);
            }

            cell.Value2 = text;
            cell.Font.Bold = true;
            cell.Font.Italic = true;
            cell.Font.Size = 9;
            cell.Font.ColorIndex = 14;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;
        }

        private static void WriteParamValueCell(Excel.Range cell, string value, string tableId)
        {
            cell.Value2 = value ?? string.Empty;
            cell.Font.Bold = false;
            cell.Font.Italic = true;
            cell.Font.Size = 9;
            cell.Font.ColorIndex = 16;
            cell.NumberFormat = "@";
            cell.WrapText = false;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;

            try
            {
                string cellAddress = cell.Address[false, false, Excel.XlReferenceStyle.xlA1];
                string rngName = $"{tableId}_{cellAddress}";
                if (rngName.Length > 30)
                {
                    rngName = rngName.Substring(0, 30);
                }

                rngName = ExcelSheetHelper.CleanUpName(rngName);

                if (ExcelSheetHelper.NamedRangeExists(rngName))
                {
                    ExcelSheetHelper.DeleteNamedRange(rngName);
                }

                cell.Name = rngName;

                if (!string.IsNullOrEmpty(value))
                {
                    Excel.Validation validation = cell.Validation;
                    validation.Delete();
                    validation.Add(
                        Excel.XlDVType.xlValidateCustom,
                        Excel.XlDVAlertStyle.xlValidAlertStop,
                        Excel.XlFormatConditionOperator.xlEqual,
                        XLEdgeValueFormatter.RemoveEquaSymbol(rngName),
                        Type.Missing);
                    validation.IgnoreBlank = true;
                    validation.InCellDropdown = false;
                    validation.ErrorTitle = "Orbit";
                    validation.ErrorMessage = "To change parameters, use the Run button on the ribbon or use param control sheet";
                    validation.ShowError = true;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WriteParamValueCell: failed to create named-range/data-validation lock");
            }
        }

        private static int Rgb(int r, int g, int b) => r + (g << 8) + (b << 16);

        private static void BuildCompanionParameterSheet(Excel.Workbook workbook, Excel.Worksheet dataSheet, string reportTitle, string paramsJson, string tableId, int dataRowCount)
        {
            string paramSheetName = $"P_{dataSheet.Name}";
            if (paramSheetName.Length >= 29)
            {
                paramSheetName = paramSheetName.Substring(0, 28);
            }

            Excel.Worksheet paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableId);
            if (paramSheet != null)
            {
                paramSheet.Cells.Clear();
            }
            else
            {
                try
                {
                    paramSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, dataSheet, Type.Missing, Type.Missing);
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(BuildCompanionParameterSheet)}: failed to add param sheet after data sheet, falling back to append-at-end - {ex.Message}");
                    paramSheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count], Type.Missing, Type.Missing);
                }

                paramSheet.Name = paramSheetName;
            }

            try
            {
                ExcelApplicationHelper.RequireActiveExcelApplication().ActiveWindow.DisplayGridlines = false;
            }
            catch (Exception)
            {
                // Cosmetic-only (gridlines display); safe to ignore if it fails.
            }

            Excel.Range titleRange = paramSheet.Range[paramSheet.Cells[1, 1], paramSheet.Cells[1, 5]];
            titleRange.Merge();
            titleRange.Value2 = reportTitle;
            titleRange.Font.Bold = true;
            titleRange.Font.Italic = true;
            titleRange.Font.Size = 10;
            titleRange.Font.ColorIndex = 2;
            titleRange.Interior.Color = Rgb(21, 96, 130);

            WriteRunInfoStrip(paramSheet, dataRowCount);

            Excel.Range sectionRange = paramSheet.Range[paramSheet.Cells[2, 1], paramSheet.Cells[2, 2]];
            sectionRange.Merge();
            sectionRange.Value2 = "Parameters Section:";
            sectionRange.Font.Bold = true;
            sectionRange.Font.Italic = true;
            sectionRange.Font.Size = 10;
            sectionRange.Font.ColorIndex = 2;
            sectionRange.Interior.Color = Rgb(241, 169, 131);

            // Writes the parameter label/value rows and bookkeeping cells; shared with the refresh path.
            RewriteParameterSectionRows(paramSheet, paramsJson, tableId, sameSheetMode: false);

            try
            {
                paramSheet.Hyperlinks.Add(paramSheet.Cells[3, 7], "", "", "Goto Report Data", "Goto Report Data");
                ((Excel.Range)paramSheet.Cells[3, 7]).Font.Size = 10;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to add 'Goto Report Data' hyperlink on companion parameter sheet");
            }

            try
            {
                ((Excel.Range)paramSheet.Cells[1, 1]).EntireColumn.AutoFit();
                ExcelApplicationHelper.RequireActiveExcelApplication().Goto(paramSheet.Range["A1"], true);
            }
            catch (Exception)
            {
                // Cosmetic-only (column width/navigation); safe to ignore if it fails.
            }
        }

        private static List<(string Label, string ValueText)> ParseParamDisplayRows(string paramsJson)
        {
            return ParseParamDisplayRows(paramsJson, out _, out _, out _, out _);
        }

        // Parses the parameter rows for display, also returning the responsibility id/value and the
        // raw/display GL segment values via out parameters so callers can persist them separately.
        // Cognitive-complexity refactor (SonarQube S3776, was 33): each foreach loop's per-item body
        // is pulled into its own helper. The out-parameters are threaded through the first helper by
        // ref (an out-parameter is a normal assignable variable once definitely assigned, which
        // oracleRespId/oracleRespValue/segmentValues/segmentDisplayValues already are by that point).
        // Every condition, comment, and error message is unchanged.
        private static List<(string Label, string ValueText)> ParseParamDisplayRows(string paramsJson, out string oracleRespId, out string oracleRespValue, out string segmentValues, out string segmentDisplayValues)
        {
            oracleRespId = null;
            oracleRespValue = null;
            segmentValues = null;
            segmentDisplayValues = null;

            var result = new List<(string Label, string ValueText)>();

            if (string.IsNullOrWhiteSpace(paramsJson))
            {
                return result;
            }

            try
            {
                using JsonDocument doc = JsonDocument.Parse(paramsJson);
                if (doc.RootElement.ValueKind != JsonValueKind.Array)
                {
                    return result;
                }

                foreach (JsonElement item in doc.RootElement.EnumerateArray())
                {
                    ProcessExtraParamsForRow(item, result, ref oracleRespId, ref oracleRespValue, ref segmentValues, ref segmentDisplayValues);
                }

                foreach (JsonElement item in doc.RootElement.EnumerateArray())
                {
                    ProcessLabelValueRow(item, result);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ParseParamDisplayRows));
            }

            return result;
        }

        // Extracted from ParseParamDisplayRows - handles one array entry's "extraParameters" block:
        // Oracle responsibility id/value and GL segment raw/display values.
        private static void ProcessExtraParamsForRow(JsonElement item, List<(string Label, string ValueText)> result, ref string oracleRespId, ref string oracleRespValue, ref string segmentValues, ref string segmentDisplayValues)
        {
            try
            {
                if (!JsonHelper.TryGetProperty(item, ExtraParametersKey, out JsonElement extraEl) ||
                    extraEl.ValueKind != JsonValueKind.Object)
                {
                    return;
                }

                (string RespId, string RespValue, string GlSegments, string GLSegmentValues) extra = ExtractExtraParams(extraEl);

                if (!string.IsNullOrWhiteSpace(extra.RespId) && !string.IsNullOrWhiteSpace(extra.RespValue))
                {
                    oracleRespId = extra.RespId;
                    oracleRespValue = extra.RespValue;
                    result.Add(("Responsibility", "'" + extra.RespValue));
                }

                // The raw segment value (IV4) and display segment value (IW4) are surfaced
                // independently, each based on its own non-blank check.
                if (!string.IsNullOrWhiteSpace(extra.GLSegmentValues))
                {
                    segmentValues = extra.GLSegmentValues;
                }

                if (!string.IsNullOrWhiteSpace(extra.GlSegments))
                {
                    segmentDisplayValues = extra.GlSegments;
                    result.Add(("GL Accounts", extra.GlSegments));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ParseParamDisplayRows: failed to extract extraParameters for one entry");
            }
        }

        // Extracted from ParseParamDisplayRows - handles one array entry's label/operator/type/value
        // display row.
        private static void ProcessLabelValueRow(JsonElement item, List<(string Label, string ValueText)> result)
        {
            try
            {
                string label;

                if (JsonHelper.TryGetProperty(item, "label", out JsonElement labelEl) && labelEl.ValueKind != JsonValueKind.Null)
                {
                    label = labelEl.ToString();
                }
                else if (JsonHelper.TryGetProperty(item, "name", out JsonElement nameEl))
                {
                    label = nameEl.ToString();
                }
                else
                {
                    label = null;
                }

                string paramOperator = JsonHelper.TryGetProperty(item, "operator", out JsonElement opEl) ? opEl.ToString() : null;
                string paramType = JsonHelper.TryGetProperty(item, "type", out JsonElement typeEl) ? typeEl.ToString() : null;

                if (string.IsNullOrWhiteSpace(label) || paramOperator == null || paramType == null)
                {
                    return;
                }

                string componentType = JsonHelper.TryGetProperty(item, "componentType", out JsonElement ctEl) ? ctEl.ToString() : null;
                string operatorKey = XLEdgeOperatorMappings.Map.FirstOrDefault(kvp => kvp.Value == paramOperator).Key ?? paramOperator;

                string valueText = BuildReportParamValue(item, componentType, paramOperator, paramType, operatorKey);

                result.Add((XLEdgeValueFormatter.RemoveEquaSymbol(label), XLEdgeValueFormatter.RemoveEquaSymbol(valueText)));
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ParseParamDisplayRows: failed to parse one parameter entry");
            }
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 55): the "ORACLE_GL_SEGMENT_DISPLAY_VALUES"
        // case (by far the deepest-nested branch here - object-vs-string-vs-parsed-string resolution,
        // then a nested foreach building the formatted string) is pulled into its own helper chain.
        // Every condition, comment, and log message is unchanged.
        private static (string RespId, string RespValue, string GlSegments, string GLSegmentValues) ExtractExtraParams(JsonElement extraParamsEl)
        {
            string respId = null;
            string respValue = null;
            string glSegments = null;
            string glSegmentValues = null;
            try
            {
                if (extraParamsEl.ValueKind != JsonValueKind.Object || !extraParamsEl.EnumerateObject().Any())
                {
                    return (respId, respValue, glSegments, glSegmentValues);
                }

                foreach (JsonProperty prop in extraParamsEl.EnumerateObject())
                {
                    switch (prop.Name)
                    {
                        case "ORACLE_RESP_ID":
                            respId = prop.Value.ValueKind == JsonValueKind.Null ? null : prop.Value.ToString();
                            break;

                        case "ORACLE_RESP_DISPLAY_VALUE":
                            respValue = prop.Value.ValueKind == JsonValueKind.Null ? null : prop.Value.ToString();
                            break;

                        case "ORACLE_GL_SEGMENT_VALUES":
                            glSegmentValues = prop.Value.ValueKind == JsonValueKind.Null ? null : prop.Value.ToString();
                            break;

                        case "ORACLE_GL_SEGMENT_DISPLAY_VALUES":
                            glSegments = ExtractGlSegmentsDisplayString(prop.Value);
                            break;
                    }
                }

                return (respId, respValue, glSegments, glSegmentValues);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ExtractExtraParams));
                return (string.Empty, string.Empty, string.Empty, string.Empty);
            }
        }

        // Extracted from ExtractExtraParams - resolves the ORACLE_GL_SEGMENT_DISPLAY_VALUES property
        // (either a real JSON object, or a string that itself parses as one) into the formatted
        // "Key=Value, Key=Value" display string.
        private static string ExtractGlSegmentsDisplayString(JsonElement propValue)
        {
            JsonElement? segmentObjectEl = ResolveSegmentObjectElement(propValue);
            return segmentObjectEl.HasValue ? FormatSegmentObjectAsString(segmentObjectEl.Value) : null;
        }

        // Extracted from ExtractExtraParams - accepts either a real JSON object, or a string whose
        // content itself parses as a JSON object (e.g. {"Company":"1000-5000","Department":"-",...}).
        private static JsonElement? ResolveSegmentObjectElement(JsonElement propValue)
        {
            if (propValue.ValueKind == JsonValueKind.Object)
            {
                return propValue;
            }

            if (propValue.ValueKind == JsonValueKind.String)
            {
                string rawText = propValue.GetString();
                if (!string.IsNullOrWhiteSpace(rawText))
                {
                    try
                    {
                        using var innerDoc = JsonDocument.Parse(rawText);
                        if (innerDoc.RootElement.ValueKind == JsonValueKind.Object)
                        {
                            return innerDoc.RootElement.Clone();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(ResolveSegmentObjectElement)}: ORACLE_GL_SEGMENT_DISPLAY_VALUES string value did not parse as a JSON object - {ex.Message}");
                    }
                }
            }

            return null;
        }

        // Extracted from ExtractExtraParams - formats a resolved segment object as "Key=Value, ..."
        // (blank/"-" values are rendered as an explicit empty-quoted string).
        private static string FormatSegmentObjectAsString(JsonElement segmentObjectEl)
        {
            var segmentString = new StringBuilder();
            foreach (JsonProperty innerProp in segmentObjectEl.EnumerateObject())
            {
                string val = innerProp.Value.ValueKind == JsonValueKind.Null ? null : innerProp.Value.ToString()?.Trim();
                if (string.IsNullOrEmpty(val) || val == "-")
                {
                    val = "\"\"";
                }

                segmentString.AppendFormat("{0}={1}, ", innerProp.Name, val);
            }

            if (segmentString.Length > 2)
            {
                segmentString.Length -= 2;
            }

            return segmentString.ToString();
        }

        private static string BuildReportParamValue(JsonElement item, string componentType, string paramOperator, string paramType, string operatorKey)
        {
            string paramValue = string.Empty;

            try
            {
                paramValue = ComputeRawParamDisplayValue(item, componentType, paramOperator, paramType);
                paramValue = ApplyParamOperatorFormatting(paramValue, paramOperator, componentType, operatorKey);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Exception in returning the parameter value.");
            }

            return paramValue.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
        }

        /// <summary>
        /// Computes the raw (pre-operator-formatting) display value for a single report parameter item.
        /// </summary>
        // Cognitive-complexity refactor (SonarQube S3776, was 28): the three "just return empty"
        // conditions are merged into a single guard (identical short-circuit order and null-safety -
        // paramOperator is confirmed non-null before .Contains("NULL") runs, exactly as before), and
        // the two deeply-nested displayValue/displayValues branches are pulled into their own
        // helpers. Every condition, comment, and log message is unchanged.
        private static string ComputeRawParamDisplayValue(JsonElement item, string componentType, string paramOperator, string paramType)
        {
            bool hasAnyProperty = item.ValueKind == JsonValueKind.Object && item.EnumerateObject().Any();

            if (!hasAnyProperty || paramOperator == null || paramType == null || paramOperator.Contains("NULL"))
            {
                return string.Empty;
            }

            if (JsonHelper.TryGetProperty(item, "displayValue", out JsonElement dvEl) && dvEl.ValueKind != JsonValueKind.Null && dvEl.ValueKind != JsonValueKind.Undefined)
            {
                return ComputeFromDisplayValue(dvEl, paramType);
            }

            if (JsonHelper.TryGetProperty(item, "displayValues", out JsonElement dvsEl) && dvsEl.ValueKind == JsonValueKind.Array)
            {
                return ComputeFromDisplayValues(dvsEl.EnumerateArray().ToList(), componentType, paramOperator, paramType);
            }

            return string.Empty;
        }

        // Extracted from ComputeRawParamDisplayValue - handles the single-item "displayValue"
        // property (array / object / scalar).
        private static string ComputeFromDisplayValue(JsonElement dvEl, string paramType)
        {
            if (dvEl.ValueKind == JsonValueKind.Array)
            {
                List<string> items = dvEl.EnumerateArray().Select(v => v.ToString()).ToList();
                return items.Count > 0
                    ? string.Join(",", items.Select(v => JoinFormatted(v, paramType)))
                    : string.Empty;
            }

            if (dvEl.ValueKind == JsonValueKind.Object)
            {
                LogUtility.LogWarn($"Type of jToken as object is not handled yet. {dvEl}");
                return string.Empty;
            }

            return Convert.ToString(XLEdgeValueFormatter.FormatValue(dvEl.ToString(), paramType));
        }

        // Extracted from ComputeRawParamDisplayValue - handles the multi-item "displayValues" array,
        // including the range/BETWEEN "X and Y" formatting.
        private static string ComputeFromDisplayValues(List<JsonElement> values, string componentType, string paramOperator, string paramType)
        {
            if (values.Count == 0)
            {
                return string.Empty;
            }

            bool isRangeStyle = (componentType != null && componentType.Contains("range")) ||
                paramOperator == "BETWEEN" || paramOperator == "NOT BETWEEN";

            if (!isRangeStyle)
            {
                return string.Join(",", values.Select(v => JoinFormatted(v.ToString(), paramType)));
            }

            if (values.Count == 2)
            {
                return $"{XLEdgeValueFormatter.FormatValue(values[0].ToString(), paramType)} and {XLEdgeValueFormatter.FormatValue(values[1].ToString(), paramType)}";
            }

            if (values.Count == 1)
            {
                return Convert.ToString(XLEdgeValueFormatter.FormatValue(values[0].ToString(), paramType));
            }

            return string.Empty;
        }

        /// <summary>
        /// Wraps the raw parameter value with operator-specific wording (e.g. "is in list ...", "is equal to ...").
        /// </summary>
        private static string ApplyParamOperatorFormatting(string paramValue, string paramOperator, string componentType, string operatorKey)
        {
            switch (paramOperator)
            {
                case "IN":
                case "NOT IN":
                    bool isSingleSelection = componentType != null &&
                        (componentType == "single-selection-prompt" || componentType == "oracle-erp-resp-selection");

                    if (paramOperator == "IN")
                    {
                        paramValue = isSingleSelection ? $"is equal to {paramValue}" : $"is in list {paramValue}";
                    }
                    else
                    {
                        paramValue = isSingleSelection ? $"does not equal {paramValue}" : $"is not in list {paramValue}";
                    }
                    break;

                default:
                    paramValue = $"{operatorKey} {paramValue}";
                    break;
            }

            return paramValue;
        }

        private static string JoinFormatted(string rawValue, string paramType)
        {
            if (!XLEdgeValueFormatter.IsNumeric(rawValue) && rawValue != null && rawValue.Contains(","))
            {
                return $"\"{rawValue}\"";
            }

            return Convert.ToString(XLEdgeValueFormatter.FormatValue(rawValue, paramType));
        }

        public static bool TryGetStoredReportXml(Excel.Workbook workbook, string listObjectName, out string title, out string metaJson, out string paramsJson)
        {
            title = null;
            metaJson = null;
            paramsJson = null;

            if (workbook == null || string.IsNullOrWhiteSpace(listObjectName))
            {
                return false;
            }

            Excel.ListObject listObject = null;
            Excel.Sheets allSheets = workbook.Worksheets;
            try
            {
                foreach (Excel.Worksheet ws in allSheets)
                {
                    try
                    {
                        if (ws.ListObjects.Count > 0 && string.Equals(ws.ListObjects[1].Name, listObjectName, StringComparison.OrdinalIgnoreCase))
                        {
                            listObject = ws.ListObjects[1];
                            break;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(ws);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "TryGetStoredReportXml: failed to locate the ListObject's parent worksheet");
            }
            finally
            {
                Marshal.ReleaseComObject(allSheets);
            }

            bool found = TryResolveReportXmlForRefresh(workbook, listObjectName, listObject, out ReportXmlRefreshResult result);
            title = result.Title;
            metaJson = result.MetaJson;
            paramsJson = result.ParamsJson;
            return found;
        }

        /// <summary>
        /// Bundles TryResolveReportXmlForRefresh's several out values - keeping this as one object
        /// instead of a fistful of out parameters is what gets that method under the 7-parameter limit.
        /// </summary>
        private sealed class ReportXmlRefreshResult
        {
            public string Title;
            public string ReportId;
            public string RunId;
            public string MetaJson;
            public string ParamsJson;
            public List<(string Original, string Modified, int RawIndex)> Mappings = new();
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 35): the per-part "try current format,
        // else try legacy format" logic is pulled out of the loop into TryResolveXmlPartForRefresh
        // and its two format-specific helpers. Every "continue" in the original per-part logic maps
        // to a "return false" here (the outer loop's finally-release still runs either way, exactly
        // as it did for "continue" in a try/finally), and the current-format XDocument.Parse call is
        // still uncaught here so a malformed part still surfaces through the same outer per-part
        // catch as before. Every condition, comment, and log message is unchanged.
        private static bool TryResolveReportXmlForRefresh(
            Excel.Workbook workbook,
            string listObjectName,
            Excel.ListObject listObject,
            out ReportXmlRefreshResult result)
        {
            result = new ReportXmlRefreshResult();

            if (workbook == null || string.IsNullOrWhiteSpace(listObjectName))
            {
                return false;
            }

            Microsoft.Office.Core.CustomXMLParts parts = workbook.CustomXMLParts;
            try
            {
                for (int i = 1; i <= parts.Count; i++)
                {
                    Microsoft.Office.Core.CustomXMLPart part = parts[i];
                    try
                    {
                        if (TryResolveXmlPartForRefresh(part.XML, listObjectName, listObject, result))
                        {
                            return true;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "TryResolveReportXmlForRefresh: error inspecting a CustomXMLPart");
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(part);
                    }
                }

                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(parts);
            }
        }

        // Extracted from TryResolveReportXmlForRefresh - tries the current XML format first (by
        // <ListObjectName> tag match); a part matching that tag never falls through to the legacy
        // format check, exactly as in the original inline logic.
        private static bool TryResolveXmlPartForRefresh(string xml, string listObjectName, Excel.ListObject listObject, ReportXmlRefreshResult result)
        {
            if (string.IsNullOrWhiteSpace(xml))
            {
                return false;
            }

            if (xml.Contains($"<ListObjectName>{listObjectName}</ListObjectName>"))
            {
                return TryResolveCurrentFormatXmlPart(xml, result);
            }

            return TryResolveLegacyFormatXmlPart(xml, listObjectName, listObject, result);
        }

        // Extracted from TryResolveReportXmlForRefresh - current XML format (ListObjectName-tagged
        // CustomXMLPart with Title/Meta/Params/Columns elements).
        private static bool TryResolveCurrentFormatXmlPart(string xml, ReportXmlRefreshResult result)
        {
            XDocument xdoc = XDocument.Parse(xml);
            result.Title = xdoc.Root?.Element("Title")?.Value ?? string.Empty;
            result.MetaJson = xdoc.Root?.Element("Meta")?.Value ?? string.Empty;
            result.ParamsJson = xdoc.Root?.Element("Params")?.Value ?? string.Empty;

            string[] titleParts = result.Title.Split('|');
            if (titleParts.Length < 3)
            {
                return false;
            }

            result.ReportId = titleParts[1];
            result.RunId = titleParts[2];

            XElement colsElem = xdoc.Root?.Element("Columns");
            if (colsElem != null)
            {
                foreach (XElement ce in colsElem.Elements(ColumnElementName))
                {
                    string orig = ce.Attribute("original")?.Value ?? string.Empty;
                    string mod = ce.Attribute("modified")?.Value ?? string.Empty;
                    int.TryParse(ce.Attribute("rawIndex")?.Value ?? "0", out int idx);
                    result.Mappings.Add((orig, mod, idx));
                }
            }

            return true;
        }

        // Extracted from TryResolveReportXmlForRefresh - legacy XML format (a "Data" element with an
        // InfoID matching listObjectName, and the report/run id derived from the table-name pattern).
        private static bool TryResolveLegacyFormatXmlPart(string xml, string listObjectName, Excel.ListObject listObject, ReportXmlRefreshResult result)
        {
            if (xml.IndexOf("<DataMeta>", StringComparison.OrdinalIgnoreCase) < 0 ||
                xml.IndexOf(listObjectName, StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            XDocument legacyDoc;
            try
            {
                legacyDoc = XDocument.Parse(xml);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "TryResolveReportXmlForRefresh: failed to parse a legacy CustomXMLPart");
                return false;
            }

            XElement dataElem = legacyDoc.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "Data");
            if (dataElem == null)
            {
                return false;
            }

            string infoId = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "InfoID")?.Value ?? string.Empty;
            if (!string.Equals(infoId, listObjectName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            Match tableNameMatch = Regex.Match(listObjectName, @"^ORB_(?<reportId>[^_]+)_(?<runId>[^_]+)_[EP]$", RegexOptions.IgnoreCase);
            if (!tableNameMatch.Success)
            {
                LogUtility.LogWarn($"TryResolveReportXmlForRefresh|Legacy metadata found for '{listObjectName}' but its name doesn't match the expected ORB_<reportId>_<runId>_E/P pattern - cannot derive report/run id.");
                return false;
            }

            result.MetaJson = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "DataMeta")?.Value ?? string.Empty;
            result.ParamsJson = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "DataParam")?.Value ?? string.Empty;
            result.ReportId = tableNameMatch.Groups["reportId"].Value;
            result.RunId = tableNameMatch.Groups["runId"].Value;
            result.Title = $"Edge|{result.ReportId}|{result.RunId}|{listObjectName}";

            if (listObject?.HeaderRowRange != null)
            {
                int col = 1;
                foreach (Excel.Range headerCell in listObject.HeaderRowRange.Cells)
                {
                    string headerText = Convert.ToString(headerCell.Value) ?? string.Empty;
                    result.Mappings.Add((headerText, headerText, col));
                    col++;
                }
            }

            return true;
        }

        private static Excel.Worksheet FindSheetWithTable(Excel.Workbook workbook, string tableId)
        {
            Excel.Sheets sheets = workbook.Worksheets;
            try
            {
                foreach (Excel.Worksheet ws in sheets)
                {
                    bool release = true;
                    try
                    {
                        foreach (Excel.ListObject lo in ws.ListObjects)
                        {
                            if (string.Equals(lo.Name, tableId, StringComparison.OrdinalIgnoreCase))
                            {
                                release = false;
                                return ws;
                            }
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
            finally
            {
                Marshal.ReleaseComObject(sheets);
            }
        }

        private static string BuildSheetName(ReportMeta reportMeta)
        {
            string name;

            if (XLEdgeAppState.Instance.FollowDrilldown && !string.IsNullOrWhiteSpace(XLEdgeAppState.Instance.ChildShtName))
            {
                string childShtName = XLEdgeAppState.Instance.ChildShtName;
                name = childShtName.Length >= 23 ? childShtName.Substring(childShtName.Length - 22, 22) : childShtName;
            }
            else if (string.IsNullOrEmpty(reportMeta.Name))
            {
                name = "No-Name";
            }
            else if (string.Equals((reportMeta.Name ?? string.Empty).Trim(), (reportMeta.BaseReportName ?? string.Empty).Trim(), StringComparison.OrdinalIgnoreCase))
            {
                string reportNameInCell = reportMeta.Name;
                name = reportNameInCell.Length >= 23 ? reportNameInCell.Substring(0, 22) : reportNameInCell;
            }
            else
            {
                string insightRptName = reportMeta.Name;
                name = insightRptName.Length >= 23 ? insightRptName.Substring(0, 22) : insightRptName;
            }

            return SanitizeSheetName(name);
        }

        private static readonly Regex SheetNameSanitizePattern = new Regex("[^a-zA-Z0-9_\\-\" \"]");

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                name = "Report";
            }

            name = SheetNameSanitizePattern.Replace(name, string.Empty);

            if (string.IsNullOrEmpty(name))
            {
                name = "Report";
            }

            if (name.Length > 22)
            {
                name = name.Substring(0, 22);
            }

            return name;
        }

        private static string MakeUniqueName(string baseName, HashSet<string> used)
        {
            if (string.IsNullOrWhiteSpace(baseName))
            {
                baseName = ColumnElementName;
            }

            string candidate = baseName;
            int suffix = 1;
            while (used.Contains(candidate))
            {
                candidate = $"{baseName}{suffix}";
                suffix++;
            }

            used.Add(candidate);
            return candidate;
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 24): the column-grouping step and the
        // per-column hyperlink-writing loop are pulled into their own helpers. The original
        // "return" (used once the hyperlink cap is hit, to exit the whole method from inside the
        // innermost loop) becomes a "reachedLimit" bool that the caller checks and turns back into a
        // "return" of its own - same overall stopping behavior. Every condition, comment, and log
        // message is unchanged.
        private static void AddDrilldownHyperlinks(Excel.Worksheet sheet, Excel.ListObject listObject, ReportMeta reportMeta)
        {
            if (reportMeta.Drilldowns == null || reportMeta.Drilldowns.Length == 0 || listObject.DataBodyRange == null)
            {
                return;
            }

            const int maxHyperlinks = 65530;
            int hyperlinkCount = 0;

            Dictionary<string, List<string>> byColumn = BuildDrilldownColumnMap(reportMeta.Drilldowns, reportMeta.ReportId);

            foreach (KeyValuePair<string, List<string>> kvp in byColumn)
            {
                if (AddHyperlinksForColumn(sheet, listObject, kvp.Key, kvp.Value, ref hyperlinkCount, maxHyperlinks))
                {
                    return;
                }
            }
        }

        // Extracted from AddDrilldownHyperlinks - groups drilldown definitions by their target
        // column name, joining every drilldown's tooltip text for columns shared by more than one.
        private static Dictionary<string, List<string>> BuildDrilldownColumnMap(RptDrilldown[] drilldowns, int reportId)
        {
            var byColumn = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (RptDrilldown dd in drilldowns)
            {
                string col = dd.DrillColumnName?.Trim();
                if (string.IsNullOrEmpty(col))
                {
                    continue;
                }

                if (!byColumn.TryGetValue(col, out List<string> list))
                {
                    list = new List<string>();
                    byColumn[col] = list;
                }

                list.Add($"DRILLDOWN|{dd.DrillReportId}|{dd.DrillReportName}|{reportId}");
            }

            return byColumn;
        }

        // Extracted from AddDrilldownHyperlinks - writes the hyperlink for every data-row cell in one
        // matched column. Returns true if the hyperlink cap was reached (caller should stop entirely).
        private static bool AddHyperlinksForColumn(Excel.Worksheet sheet, Excel.ListObject listObject, string columnName, List<string> tooltipParts, ref int hyperlinkCount, int maxHyperlinks)
        {
            int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, columnName);
            if (matchCol <= 0)
            {
                return false;
            }

            string tooltip = string.Join(",", tooltipParts);
            if (tooltip.Length > 255)
            {
                tooltip = tooltip.Substring(0, 250) + "...";
            }

            Excel.Range dataRange = listObject.DataBodyRange;
            try
            {
                for (int r = 1; r <= dataRange.Rows.Count; r++)
                {
                    if (hyperlinkCount >= maxHyperlinks)
                    {
                        LogUtility.LogWarn($"Reached maximum hyperlink limit of {maxHyperlinks}; stopping further drilldown hyperlinks.");
                        return true;
                    }

                    Excel.Range cell = (Excel.Range)dataRange.Cells[r, matchCol];
                    try
                    {
                        if (cell.Value2 != null && cell.Value2.ToString().Length > 0)
                        {
                            sheet.Hyperlinks.Add(cell, "", cell.Address, tooltip, cell.Value2.ToString());
                            hyperlinkCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"Failed to add drilldown hyperlink at {cell.Address}");
                    }
                    finally
                    {
                        // Part of the COM-leak fix (see AddImageColumn) - every Range obtained
                        // here is a live COM reference that must be explicitly released.
                        Marshal.ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dataRange);
            }

            return false;
        }

        private static void DeleteReportShapes(Excel.Worksheet sheet)
        {
            if (sheet == null || sheet.Shapes == null || sheet.Shapes.Count == 0)
            {
                return;
            }

            try
            {
                for (int i = sheet.Shapes.Count; i >= 1; i--)
                {
                    Excel.Shape shape = null;
                    try
                    {
                        shape = sheet.Shapes.Item(i);
                        if (shape.Name.ToUpperInvariant().Contains("ORB_"))
                        {
                            shape.Delete();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "DeleteReportShapes: failed to delete one shape");
                    }
                    finally
                    {
                        if (shape != null)
                        {
                            Marshal.ReleaseComObject(shape);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(DeleteReportShapes));
            }
        }

        private static void AddAttachmentAndImageColumns(Excel.Worksheet sheet, Excel.ListObject listObject, ReportMeta reportMeta)
        {
            if (reportMeta.Columns == null || listObject.DataBodyRange == null)
            {
                return;
            }

            const int maxHyperlinks = 65530;
            int hyperlinkCount = 0;

            foreach (RptColumn col in reportMeta.Columns)
            {
                try
                {
                    if (col.IsFileAttached)
                    {
                        hyperlinkCount = AddAttachmentColumn(sheet, listObject, col, hyperlinkCount, maxHyperlinks);
                        continue;
                    }

                    string outputType = col.Properties?.OutputProp?.Type;
                    if (string.IsNullOrWhiteSpace(outputType))
                    {
                        continue;
                    }

                    if (string.Equals(outputType, "HYPERLINK", StringComparison.OrdinalIgnoreCase))
                    {
                        hyperlinkCount = AddHyperlinkColumn(listObject, col, hyperlinkCount, maxHyperlinks);
                    }
                    else if (string.Equals(outputType, "IMAGE", StringComparison.OrdinalIgnoreCase))
                    {
                        AddImageColumn(sheet, listObject, col);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"AddAttachmentAndImageColumns: failed for column '{col?.Name}'");
                }
            }
        }

        private static int AddAttachmentColumn(Excel.Worksheet sheet, Excel.ListObject listObject, RptColumn col, int hyperlinkCount, int maxHyperlinks)
        {
            int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, col.Name?.Trim() ?? string.Empty);
            if (matchCol <= 0)
            {
                return hyperlinkCount;
            }

            Excel.Range dataRange = listObject.DataBodyRange;
            try
            {
                for (int r = 1; r <= dataRange.Rows.Count; r++)
                {
                    if (hyperlinkCount >= maxHyperlinks)
                    {
                        LogUtility.LogWarn($"Reached maximum hyperlink limit of {maxHyperlinks}; stopping further attachment hyperlinks.");
                        return hyperlinkCount;
                    }

                    Excel.Range cell = (Excel.Range)dataRange.Cells[r, matchCol];
                    try
                    {
                        object rawValue = cell.Value;
                        string rawText = rawValue != null ? Convert.ToString(rawValue) : string.Empty;
                        if (string.IsNullOrWhiteSpace(rawText))
                        {
                            continue;
                        }

                        if (!AttachmentLinkHelper.TryParseAttachmentLink(rawText, out string displayValue, out string linkValue))
                        {
                            continue;
                        }

                        sheet.Hyperlinks.Add(cell, "", "", linkValue, displayValue);
                        hyperlinkCount++;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"Failed to add attachment hyperlink at {cell.Address}");
                    }
                    finally
                    {
                        // Part of the COM-leak fix (see AddImageColumn) - every Range obtained here
                        // is a live COM reference that must be explicitly released.
                        Marshal.ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dataRange);
            }

            return hyperlinkCount;
        }

        private static int AddHyperlinkColumn(Excel.ListObject listObject, RptColumn col, int hyperlinkCount, int maxHyperlinks)
        {
            int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, col.Name?.Trim() ?? string.Empty);
            if (matchCol <= 0)
            {
                return hyperlinkCount;
            }

            Excel.Range dataRange = listObject.DataBodyRange;
            try
            {
                for (int r = 1; r <= dataRange.Rows.Count; r++)
                {
                    if (hyperlinkCount >= maxHyperlinks)
                    {
                        LogUtility.LogWarn($"Reached maximum hyperlink limit of {maxHyperlinks}; stopping further hyperlink columns.");
                        return hyperlinkCount;
                    }

                    Excel.Range cell = (Excel.Range)dataRange.Cells[r, matchCol];
                    try
                    {
                        object rawValue = cell.Value2;
                        if (rawValue == null)
                        {
                            continue;
                        }

                        string linkRef = Convert.ToString(rawValue);
                        string displayText = !string.IsNullOrWhiteSpace(col.Properties?.OutputProp?.HlinkDisplayValue)
                            ? col.Properties.OutputProp.HlinkDisplayValue
                            : linkRef;

                        cell.Hyperlinks.Add(cell, linkRef, cell.Address, string.Empty, displayText);
                        hyperlinkCount++;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, $"Failed to add hyperlink-type column value at {cell.Address}");
                    }
                    finally
                    {
                        // Part of the COM-leak fix (see AddImageColumn) - every Range obtained here
                        // is a live COM reference that must be explicitly released.
                        Marshal.ReleaseComObject(cell);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dataRange);
            }

            return hyperlinkCount;
        }

        // Cognitive-complexity refactor (SonarQube S3776, was 34): the per-row body is pulled into
        // EmbedImageForRow, with the destination-path-building and post-AddPicture sizing logic
        // further split out. This is deliberately conservative about the existing COM-leak fix: every
        // COM release still happens at exactly the same point, in the same finally block, as before -
        // entireRow/entireColumn are threaded out of ApplyImagePlacementSizing via out parameters
        // rather than being released inside it, so if an exception is thrown before one of them would
        // have been assigned, the caller's variable simply keeps its pre-call null value (exactly
        // like today) and the existing null-check-before-release guards in the caller's finally still
        // behave identically. Every original "continue" becomes a "return" from the per-row method
        // (the per-row finally still runs either way, exactly as it did for "continue" in a
        // try/finally). Every condition, comment, and log message is unchanged.
        private static void AddImageColumn(Excel.Worksheet sheet, Excel.ListObject listObject, RptColumn col)
        {
            int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, col.Name?.Trim() ?? string.Empty);
            if (matchCol <= 0)
            {
                return;
            }

            Excel.Range dataRange = listObject.DataBodyRange;
            var rowMaxHeights = new Dictionary<int, double>();
            var colMaxWidths = new Dictionary<int, double>();

            double imgHeight = (col.Properties?.OutputProp?.ImgHeight ?? 0) / 72.0 * 96.0;
            double imgWidth = (col.Properties?.OutputProp?.ImgWidth ?? 0) / 72.0 * 96.0;
            if (imgHeight <= 0) imgHeight = 20;
            if (imgWidth <= 0) imgWidth = 20;

            try
            {
                for (int r = 1; r <= dataRange.Rows.Count; r++)
                {
                    EmbedImageForRow(sheet, dataRange, r, matchCol, (imgHeight, imgWidth), rowMaxHeights, colMaxWidths);
                }
            }
            finally
            {
                Marshal.ReleaseComObject(dataRange);
            }
        }

        // Extracted from AddImageColumn - downloads and embeds the image for one data row's cell (if
        // any), tracking per-row/per-column max size for row-height/column-width autosizing.
        private static void EmbedImageForRow(Excel.Worksheet sheet, Excel.Range dataRange, int r, int matchCol, (double Height, double Width) imageSize, Dictionary<int, double> rowMaxHeights, Dictionary<int, double> colMaxWidths)
        {
            Excel.Range cell = (Excel.Range)dataRange.Cells[r, matchCol];
            // entireRow/entireColumn/imgShape below are only ever assigned when actually needed -
            // released in the finally block alongside cell as part of the COM-leak fix: every
            // Range/Shape obtained from Excel here is a live COM reference (RCW) that has to be
            // explicitly released, or it lingers until the next GC pass finalizes it - across a
            // report with many image rows, that's a lot of outstanding references piling up, which
            // is what was keeping excel.exe running in the background after closing the workbook
            // following a report with images.
            Excel.Range entireRow = null;
            Excel.Range entireColumn = null;
            Excel.Shape imgShape = null;
            string destinationPath = null;
            try
            {
                object rawValue = cell.Value;
                if (rawValue == null)
                {
                    return;
                }

                string url = Convert.ToString(rawValue);
                cell.Clear();

                if (string.IsNullOrWhiteSpace(url))
                {
                    return;
                }

                destinationPath = BuildImageDestinationPath(url);

                bool downloaded = ImageDownloadHelper.TryDownloadImage(url, destinationPath);
                if (!downloaded || !File.Exists(destinationPath))
                {
                    return;
                }

                // cell.Left/cell.Top are declared as `object` in the Excel Interop PIA (boxed
                // double at runtime) - a direct (float) cast is a strict CLR unboxing conversion
                // that only succeeds if the boxed type is exactly float, so it always threw
                // InvalidCastException here. Ported from FormProcessBar.vb, which passed
                // CR.Left/CR.Top with no cast at all - VB's Option-Strict-Off runtime conversion
                // helpers handle boxed-double-to-Single conversions the C# unboxing cast can't.
                // Convert.ToDouble first (matches the same pattern already used for this exact
                // property elsewhere - see ExcelWindowHelper.cs, XLEdgeDrilldownReports.xaml.cs).
                imgShape = sheet.Shapes.AddPicture(
                    destinationPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue,
                    (float)Convert.ToDouble(cell.Left), (float)Convert.ToDouble(cell.Top), (float)imageSize.Height, (float)imageSize.Width);

                ApplyImagePlacementSizing(cell, imgShape, rowMaxHeights, colMaxWidths, out entireRow, out entireColumn);

                string address = cell.Address[false, false, Excel.XlReferenceStyle.xlA1];

                imgShape.Name = $"ORB_{sheet.Name}_{address}";
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Failed to embed image at {cell.Address}");
            }
            finally
            {
                if (destinationPath != null)
                {
                    try { File.Delete(destinationPath); }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(AddAttachmentAndImageColumns)}: failed to delete temp image '{destinationPath}' - {ex.Message}");
                    }
                }

                if (entireRow != null) Marshal.ReleaseComObject(entireRow);
                if (entireColumn != null) Marshal.ReleaseComObject(entireColumn);
                if (imgShape != null) Marshal.ReleaseComObject(imgShape);
                Marshal.ReleaseComObject(cell);
            }
        }

        // Extracted from AddImageColumn (EmbedImageForRow) - sanitizes the image URL's file name and
        // builds the temporary download destination path under the user's Downloads folder.
        private static string BuildImageDestinationPath(string url)
        {
            string fileName = url.Contains("/") ? url.Substring(url.LastIndexOf('/') + 1) : url;
            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(invalidChar, '_');
            }

            string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            return Path.Combine(downloadsFolder, fileName);
        }

        // Extracted from AddImageColumn (EmbedImageForRow) - after a picture is embedded, grows the
        // row height / column width to fit it if it's the tallest/widest seen so far for that
        // row/column. entireRow/entireColumn are only assigned (via out) when actually touched, so
        // the caller's existing null-check-before-release in its finally block behaves exactly as it
        // did when this logic was inline.
        private static void ApplyImagePlacementSizing(Excel.Range cell, Excel.Shape imgShape, Dictionary<int, double> rowMaxHeights, Dictionary<int, double> colMaxWidths, out Excel.Range entireRow, out Excel.Range entireColumn)
        {
            entireRow = null;
            entireColumn = null;

            int rowIndex = cell.Row;
            int colIndex = cell.Column;

            double actualRowHeight = Math.Min(imgShape.Height, 409);
            if (!rowMaxHeights.TryGetValue(rowIndex, out double existingRowHeight) || actualRowHeight > existingRowHeight)
            {
                rowMaxHeights[rowIndex] = actualRowHeight;
                entireRow = cell.EntireRow;
                entireRow.RowHeight = actualRowHeight;
            }

            double colWidthEstimate = imgShape.Width / 10.0;
            double adjustedColWidth = colWidthEstimate + (colWidthEstimate - 1);
            if (!colMaxWidths.TryGetValue(colIndex, out double existingColWidth) || adjustedColWidth > existingColWidth)
            {
                colMaxWidths[colIndex] = adjustedColWidth;
                entireColumn = cell.EntireColumn;
                entireColumn.ColumnWidth = adjustedColWidth;
            }
        }

        private static void WriteTempCsv(string recsStr, string reportRunId)
        {
            string tempFile = Path.Combine(XLEdgeAppPaths.TempFolder, $"{reportRunId}.csv");
            try
            {
                Directory.CreateDirectory(XLEdgeAppPaths.TempFolder);
                File.WriteAllText(tempFile, recsStr);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error writing to file: {tempFile}");
                throw;
            }
        }

        public static async Task CreateReportFromListObjectAsync(string listObjectName, AppOverlay appOverlay = null, bool useWaitWindow = false)
        {
            using var excelBulkScope = new ExcelBulkOperationScope();

            if (string.IsNullOrWhiteSpace(listObjectName))
                return;

            try
            {
                var excelApp = XLApp.App;
                if (excelApp == null)
                    throw new InvalidOperationException("Excel instance not available.");

                var sheet = excelApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                if (sheet == null)
                    throw new InvalidOperationException("No active worksheet.");

                Microsoft.Office.Interop.Excel.ListObject lo = null;
                try
                {
                    lo = sheet.ListObjects[listObjectName];
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(CreateReportFromListObjectAsync)}: ListObject '{listObjectName}' not found on active sheet - {ex.Message}");
                }

                if (lo == null)
                {
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => appOverlay?.ShowError($"Table '{listObjectName}' not found."));
                    return;
                }

                string possibleName = lo.Name;
                if (!string.IsNullOrWhiteSpace(possibleName))
                {
                    if (possibleName.Contains("|"))
                    {
                        await CreateReportFromTitleAsync(possibleName, appOverlay, useWaitWindow);
                        return;
                    }

                    var partsUnderscore = possibleName.Split('_');
                    if (partsUnderscore.Length >= 3)
                    {
                        string t = partsUnderscore[0];
                        string rid = partsUnderscore[1];
                        string run = partsUnderscore[2];
                        string sh = partsUnderscore.Length >= 4 ? partsUnderscore[3] : string.Empty;
                        string reconstructed = $"{t}|{rid}|{run}|{sh}";
                        await CreateReportFromTitleAsync(reconstructed, appOverlay, useWaitWindow);
                        return;
                    }
                }

                await RefreshListObjectAsync(listObjectName, appOverlay);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to refresh report from ListObject");
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => appOverlay?.ShowError(ex.Message));
            }
        }

        private static async Task CleanupAsync()
        {
            await UiDispatcher.RunAsync(DoCleanupOnUiThreadAsync);
            await ReleaseKeyboardFocusFromTaskPaneAsync();
            await ForceReleaseOutstandingComReferencesAsync();
        }

        /// <summary>
        /// COM-leak safety net, run once at the end of every report-generation/refresh/cancel/error
        /// path (CleanupAsync is the single choke point all of them funnel through). Explicit
        /// Marshal.ReleaseComObject calls were added to the hottest per-row loops (AddImageColumn,
        /// AddDrilldownHyperlinks, AddAttachmentColumn, AddHyperlinkColumn), but this forces the CLR
        /// to actually finalize anything still outstanding elsewhere (e.g. RefreshListObjectAsync's
        /// smaller per-column loops), rather than leaving it to an unpredictable future GC pass - this
        /// is Microsoft's own documented mitigation for Office interop leaks. It matters here
        /// specifically because XLEdge runs in-process inside excel.exe (not as an external Automation
        /// client): Excel's own shutdown sequence waits for every in-process COM reference to actually
        /// be released before the process can fully exit, so leftover un-finalized RCWs (Runtime
        /// Callable Wrappers) from a report run are exactly what was keeping excel.exe running in the
        /// background after the workbook was closed. Run via Task.Run so this blocking collection
        /// doesn't tie up the awaiting UI-thread continuation any longer than necessary.
        /// </summary>
        private static async Task ForceReleaseOutstandingComReferencesAsync()
        {
            try
            {
                await Task.Run(() =>
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"{nameof(ForceReleaseOutstandingComReferencesAsync)}: GC collection pass failed - {ex.Message}");
            }
        }

        /// <summary>
        /// Releases native keyboard focus from the embedded WebView2 task pane control back to Excel,
        /// by blurring WebView2's active element, re-activating Excel's main window, and sending a
        /// dummy F2/Esc keystroke pair so Excel genuinely re-acquires OS keyboard focus.
        /// </summary>
        // internal (not private): RibEdgeRefreshAll_OnClick (AddinModule.cs) also calls this directly,
        // once after its own aggregated summary message for a book-wide RefreshAll.
        internal static async Task ReleaseKeyboardFocusFromTaskPaneAsync()
        {
            try
            {
                try
                {
                    await UiDispatcher.RunAsync(() =>
                    {
                        var addinModule = XLEdge.AddinModule.CurrentInstance;
                        if (addinModule != null)
                        {
                            var pane = addinModule.GetPaneInstance();
                            if (pane != null)
                            {
                                pane.ReleaseFocusToExcel();
                            }
                        }
                    });
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ReleaseKeyboardFocusFromTaskPaneAsync: Failed to release focus from task pane - {ex.Message}");
                }

                try
                {
                    var excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                    if (excelApp != null)
                    {
                        ExcelWindowHelper.ActivateExcelMainWindow(excelApp);

                        // Nudge keyboard focus off the WebView2 control by actually selecting a
                        // different cell then reselecting the original one - a real COM selection
                        // change, not a synthetic keystroke. This used to be preceded by SendKeys
                        // {F2}/{ESC} "dummy key" presses, which were found to be flipping the user's
                        // NumLock state on every report run (SendKeys/Application.SendKeys shares the
                        // same low-level toggle-key-detection path implicated in that). Removed -
                        // the Sleep below still runs on a background thread so Excel's own
                        // STA/message-pump thread stays free to process the COM selection calls,
                        // which are marshalled onto it via UiDispatcher.Run.
                        await Task.Run(() =>
                        {
                            try
                            {
                                Thread.Sleep(50);

                                Excel.Range originalCell = null;
                                UiDispatcher.Run(() =>
                                {
                                    if (excelApp.ActiveCell != null)
                                    {
                                        originalCell = excelApp.ActiveCell;
                                        var target = originalCell.Offset[1, 0];
                                        target?.Select();
                                    }
                                });
                                Thread.Sleep(20);
                                UiDispatcher.Run(() => originalCell?.Select());
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogWarn($"ReleaseKeyboardFocusFromTaskPaneAsync: Background focus reset failed - {ex.Message}");
                            }
                        });
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ReleaseKeyboardFocusFromTaskPaneAsync: Focus activation failed - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ReleaseKeyboardFocusFromTaskPaneAsync: Failed to reset keyboard focus");
            }
        }

        private static async Task DoCleanupOnUiThreadAsync()
        {
            try
            {
                if (_appOverlay != null)
                {
                    await _appOverlay.HideBusyAsync();
                    return;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "CleanupAsync: HideBusyAsync failed");
                return;
            }

            try
            {
                _waitWindow?.RequestClose();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "CleanupAsync: RequestClose failed");
            }
        }

        private static string BuildCustomXml(string title, string listObjectName, string metaJson, string paramsJson, List<(string Original, string Modified, int RawIndex)> mappings)
        {
            var doc = new XDocument(new XElement("XLEdgeReport",
                new XElement("Title", title ?? string.Empty),
                new XElement("ListObjectName", listObjectName ?? string.Empty),
                new XElement("Meta", new XCData(metaJson ?? string.Empty)),
                new XElement("Params", new XCData(paramsJson ?? string.Empty)),
                new XElement("Columns",
                    mappings.Select(m =>
                        new XElement(ColumnElementName,
                            new XAttribute("original", m.Original ?? string.Empty),
                            new XAttribute("modified", m.Modified ?? string.Empty),
                            new XAttribute("rawIndex", m.RawIndex)
                        )
                    )
                )
            ));

            return doc.ToString(SaveOptions.DisableFormatting);
        }

        private static void SaveCustomXmlPart(Excel.Workbook wb, string xml, string listObjectName)
        {
            try
            {
                if (wb == null || string.IsNullOrWhiteSpace(xml)) return;

                var parts = wb.CustomXMLParts;
                try
                {
                    for (int i = parts.Count; i >= 1; i--)
                    {
                        var part = parts[i];
                        try
                        {
                            if (!string.IsNullOrWhiteSpace(part.XML) &&
                                (part.XML.Contains($"<ListObjectName>{listObjectName}</ListObjectName>") ||
                                 part.XML.Contains($"<InfoID>{listObjectName}</InfoID>")))
                            {
                                part.Delete();
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(SaveCustomXmlPart)}: failed to inspect/delete a CustomXMLPart for '{listObjectName}' - {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(SaveCustomXmlPart)}: failed to enumerate CustomXMLParts for cleanup for '{listObjectName}' - {ex.Message}");
                }

                parts.Add(xml);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SaveCustomXmlPart");
            }
        }

        public static async Task RefreshListObjectAsync(string listObjectName, AppOverlay appOverlay = null, bool useWaitWindow = true, string paramsJsonPayload = null, bool collectErrors = false)
        {
            using var excelBulkScope = new ExcelBulkOperationScope();

            if (string.IsNullOrWhiteSpace(listObjectName)) return;

            XLEdgeWaitWindow waitWindow = null;
            CancellationHelper cancelHelper = null;

            try
            {
                var excelApp = XLApp.App;
                if (excelApp == null) throw new InvalidOperationException("Excel instance not available.");

                // Check edit mode
                try
                {
                    var ac = excelApp.ActiveCell;
                    var _ = ac?.Address;
                }
                catch (Exception editModeEx)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: Excel appears to be in edit mode - {editModeEx.Message}");
                    await HandleFailureAsync("Excel is in edit mode. Please exit edit mode (press Enter or Esc) and try again.", null, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                if (useWaitWindow)
                {
                    var waitCancelHelper = new CancellationHelper();
                    cancelHelper = waitCancelHelper;
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            waitWindow = new XLEdgeWaitWindow(waitCancelHelper);
                            waitWindow.SetProcessTitle("Refreshing report", MahApps.Metro.IconPacks.PackIconFontAwesomeKind.SpinnerSolid);
                            waitWindow.SetProcessMessage("Preparing to refresh report...");
                            waitWindow.StartMonitoring();
                            waitWindow.Show();
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "Failed to show wait window for refresh");
                        }
                    });
                }
                else
                {
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => appOverlay?.ShowBusyasyn("Refreshing report..."));
                }

                var wb = excelApp.ActiveWorkbook;
                var sheet = excelApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                if (sheet == null) throw new InvalidOperationException("No active worksheet.");

                Microsoft.Office.Interop.Excel.ListObject lo = null;
                try { lo = sheet.ListObjects[listObjectName]; }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: ListObject '{listObjectName}' not found - {ex.Message}");
                }
                if (lo == null)
                {
                    await HandleFailureAsync($"Table '{listObjectName}' not found.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                DeleteReportShapes(sheet);

                // --- STEP 1: Get stored report data from CustomXMLParts (Meta Data) ---
                // This is always from cache - meta data NEVER changes during refresh
                if (!TryResolveReportXmlForRefresh(wb, listObjectName, lo, out ReportXmlRefreshResult xmlResult))
                {
                    await HandleFailureAsync("No metadata found for this table.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                string title = xmlResult.Title;
                string runId = xmlResult.RunId;
                string storedMetaJson = xmlResult.MetaJson;
                string storedParamsJson = xmlResult.ParamsJson;
                List<(string Original, string Modified, int RawIndex)> mappings = xmlResult.Mappings;

                LogUtility.LogDebug($"RefreshListObjectAsync|Using meta data from CustomXMLParts for table: {listObjectName}");

                string eeLoginUrl = XLEdgeAppState.Instance.LoginUrl;
                if (string.IsNullOrWhiteSpace(eeLoginUrl))
                {
                    await HandleFailureAsync("Login URL is not set.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                // --- STEP 2: Check if we have a payload from control sheet ---
                // paramsJsonPayload is already the raw payload (no display values)
                bool hasParamsPayload = !string.IsNullOrEmpty(paramsJsonPayload);
                string paramsWithLabels = hasParamsPayload ? paramsJsonPayload : null;

                if (hasParamsPayload)
                {
                    LogUtility.LogDebug($"RefreshListObjectAsync|Using control sheet payload for table: {listObjectName}");
                }
                else
                {
                    LogUtility.LogDebug($"RefreshListObjectAsync|No control sheet payload - using original params from CustomXMLParts");
                }


                // --- STEP 3: Fetch CSV data (with payload or empty) ---

                if (cancelHelper == null)
                {
                    cancelHelper = new CancellationHelper();
                }

                await SetRefreshMessage("Downloading report data...", waitWindow, appOverlay, useWaitWindow);

                string csvUrl = $"{eeLoginUrl.TrimEnd('/')}/rest/secure/report/runner?runId={runId}&type=csv";
                string csvResponse = null;
                try
                {
                    csvResponse = await ApiHelper.ServerAPI(csvUrl, "JSON", paramsWithLabels ?? string.Empty, "POST", cancelHelper.GetToken());
                }
                catch (OperationCanceledException)
                {
                    LogUtility.LogWarn("CSV fetch cancelled by user.");
                    await ApiHelper.NotifyCancelRunAsync(eeLoginUrl, runId);
                    await CancelCleanupAsync(waitWindow, appOverlay, useWaitWindow, collectErrors);
                    if (collectErrors) throw;
                    if (useWaitWindow) { try { waitWindow?.RequestClose(); } catch (Exception ex) { LogUtility.LogException(ex, "Failed to close wait window on cancel"); } }
                    else { await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () => { if (appOverlay != null) await appOverlay.HideBusyAsync(); }); }
                    return;
                }
                catch (ApiTimeoutException ex)
                {
                    LogUtility.LogException(ex, "RefreshListObjectAsync: CSV request timed out");
                    await HandleFailureAsync(RequestTimedOutMessage, waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                if (string.IsNullOrWhiteSpace(csvResponse))
                {
                    await HandleFailureAsync("Failed to download report for refresh.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                // --- STEP 4: Parse CSV data ---
                await SetRefreshMessage("Parsing report data...", waitWindow, appOverlay, useWaitWindow);
                var rows = ParseCsv(csvResponse).ToList();
                if (rows.Count == 0)
                {
                    await HandleFailureAsync("No data in report.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

                var rawHeader = rows[0];
                int rawCols = rawHeader.Count;
                int newDataCount = Math.Max(0, rows.Count - 1);

                // --- STEP 5: Add missing columns from raw data into the table ---
                try
                {
                    var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 1; i <= lo.ListColumns.Count; i++)
                    {
                        try { existingNames.Add(lo.ListColumns[i].Name); }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read ListColumn[{i}].Name - {ex.Message}");
                        }
                    }

                    for (int i = 1; i <= rawCols; i++)
                    {
                        bool mapped = mappings.Any(m => m.RawIndex == i);
                        if (mapped) continue;

                        string orig = rawHeader[i - 1] ?? string.Empty;
                        string baseName = orig.Trim();
                        if (string.IsNullOrWhiteSpace(baseName)) baseName = ColumnElementName + i;

                        string mod = baseName;
                        int suffix = 1;
                        while (existingNames.Contains(mod) || mappings.Any(m => string.Equals(m.Modified, mod, StringComparison.OrdinalIgnoreCase)))
                        {
                            mod = baseName + suffix.ToString();
                            suffix++;
                        }

                        try
                        {
                            var added = lo.ListColumns.Add();
                            try { added.Name = mod; }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to rename new column to '{mod}' - {ex.Message}");
                            }
                            existingNames.Add(mod);
                            mappings.Add((Original: orig, Modified: mod, RawIndex: i));
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "Failed to add missing column to table");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed while adding missing columns to table - {ex.Message}");
                }

                // --- STEP 6: Re-order columns in the table to match original CSV order ---
                try
                {
                    var desiredOrder = mappings.OrderBy(m => m.RawIndex).Select(m => m.Modified).ToList();
                    int desiredCount = desiredOrder.Count;

                    for (int pos = 1; pos <= desiredCount && pos <= lo.ListColumns.Count; pos++)
                    {
                        string desiredName = desiredOrder[pos - 1];
                        string currentName = string.Empty;
                        try { currentName = lo.ListColumns[pos].Name; }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read ListColumns[{pos}].Name during reorder - {ex.Message}");
                        }

                        if (string.Equals(currentName, desiredName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        int curIndex = -1;
                        for (int i = 1; i <= lo.ListColumns.Count; i++)
                        {
                            try
                            {
                                if (string.Equals(lo.ListColumns[i].Name, desiredName, StringComparison.OrdinalIgnoreCase))
                                {
                                    curIndex = i;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read ListColumns[{i}].Name while searching for '{desiredName}' - {ex.Message}");
                            }
                        }

                        if (curIndex == -1)
                            continue;

                        try
                        {
                            var rangeA = lo.ListColumns[pos].Range;
                            var rangeB = lo.ListColumns[curIndex].Range;
                            if (rangeA != null && rangeB != null)
                            {
                                var temp = rangeA.Value2;
                                rangeA.Value2 = rangeB.Value2;
                                rangeB.Value2 = temp;
                            }

                            try
                            {
                                var headerRange = lo.HeaderRowRange;
                                if (headerRange != null)
                                {
                                    var headerCellObj = headerRange.Cells[1, pos];
                                    if (headerCellObj is Excel.Range headerCell)
                                        headerCell.Value2 = desiredName;
                                }
                                try { lo.ListColumns[pos].Name = desiredName; }
                                catch (Exception ex)
                                {
                                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to rename ListColumns[{pos}] to '{desiredName}' - {ex.Message}");
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to sync header name for '{desiredName}' - {ex.Message}");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "Failed to reorder table columns");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed while reordering table columns - {ex.Message}");
                }

                // --- STEP 7: Ensure table has at least one data row ---
                int currentRows = 0;
                try { currentRows = lo.DataBodyRange?.Rows.Count ?? 0; }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read DataBodyRange.Rows.Count - {ex.Message}");
                    currentRows = 0;
                }
                if (currentRows == 0)
                {
                    try { lo.ListRows.Add(); }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"{nameof(RefreshListObjectAsync)}: failed to add an initial data row to the table - {ex.Message}");
                    }
                }

                // --- STEP 8: Capture first data row formulas ---
                var firstRowFormulas = new Dictionary<int, string>();
                try
                {
                    var firstRowRange = lo.DataBodyRange.Resize[1, lo.ListColumns.Count];
                    for (int c = 1; c <= lo.ListColumns.Count; c++)
                    {
                        try
                        {
                            var cell = firstRowRange.Cells[1, c] as Microsoft.Office.Interop.Excel.Range;
                            var formula = cell?.Formula as string;
                            if (!string.IsNullOrWhiteSpace(formula) && formula.StartsWith("=", StringComparison.Ordinal))
                                firstRowFormulas[c] = formula;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read first-row formula for column {c} - {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed while capturing first-row formulas - {ex.Message}");
                }

                int headerRowIdx = lo.HeaderRowRange.Row;
                int dataStartRow = headerRowIdx + 1;
                int targetTotalRows = Math.Max(1, newDataCount);

                // --- STEP 9: Adjust table rows ---
                try
                {
                    while ((lo.DataBodyRange?.Rows.Count ?? 0) < targetTotalRows)
                    {
                        lo.ListRows.Add();
                    }

                    while ((lo.DataBodyRange?.Rows.Count ?? 0) > targetTotalRows)
                    {
                        var last = lo.ListRows[lo.ListRows.Count];
                        last.Delete();
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"{nameof(RefreshListObjectAsync)}: failed to adjust table row count to {targetTotalRows} - {ex.Message}");
                }

                int tableCols = lo.ListColumns.Count;

                // --- STEP 10: Write data to table ---
                await SetRefreshMessage("Writing data to Excel...", waitWindow, appOverlay, useWaitWindow);

                if (newDataCount > 0)
                {
                    for (int tc = 1; tc <= tableCols; tc++)
                    {
                        string modifiedName = string.Empty;
                        try { modifiedName = lo.ListColumns[tc].Name; }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read ListColumns[{tc}].Name while writing refreshed data - {ex.Message}");
                            continue;
                        }

                        var map = mappings.FirstOrDefault(m => string.Equals(m.Modified, modifiedName, StringComparison.OrdinalIgnoreCase));
                        if (map.Modified == null)
                        {
                            continue;
                        }

                        int rawIndex = map.RawIndex;
                        if (rawIndex < 1 || rawIndex > rawCols)
                        {
                            continue;
                        }

                        bool hasRow1Formula = firstRowFormulas.ContainsKey(tc);
                        int colWriteStartRow = hasRow1Formula ? dataStartRow + 1 : dataStartRow;
                        int colRowCount = targetTotalRows - (hasRow1Formula ? 1 : 0);

                        if (colRowCount <= 0)
                        {
                            continue;
                        }

                        try
                        {
                            object[,] colArr = new object[colRowCount, 1];
                            for (int i = 0; i < colRowCount; i++)
                            {
                                int physicalRow = colWriteStartRow + i;
                                int csvRecordIndex = physicalRow - dataStartRow + 1;
                                var rowVals = (csvRecordIndex >= 1 && csvRecordIndex <= newDataCount) ? rows[csvRecordIndex] : null;
                                colArr[i, 0] = (rowVals != null && rawIndex - 1 < rowVals.Count) ? rowVals[rawIndex - 1] : string.Empty;
                            }

                            var colStartCell = (Excel.Range)sheet.Cells[colWriteStartRow, tc];
                            var colEndCell = (Excel.Range)sheet.Cells[colWriteStartRow + colRowCount - 1, tc];
                            sheet.Range[colStartCell, colEndCell].Value2 = colArr;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, $"Failed writing refreshed data for column {tc} ('{modifiedName}')");
                        }
                    }
                }

                // --- STEP 11: Fill down formulas from first data row where applicable ---
                try
                {
                    int lastRow = dataStartRow + targetTotalRows - 1;
                    for (int c = 1; c <= tableCols; c++)
                    {
                        if (firstRowFormulas.TryGetValue(c, out _))
                        {
                            var topCell = (Excel.Range)sheet.Cells[dataStartRow, c];
                            var fillRange = sheet.Range[topCell, sheet.Cells[lastRow, c]];
                            try { fillRange.FillDown(); }
                            catch (Exception ex)
                            {
                                LogUtility.LogWarn($"{nameof(RefreshListObjectAsync)}: failed to fill down formula for column {c} - {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed while filling down preserved formulas - {ex.Message}");
                }

                // --- STEP 12: Handle RefreshSync ---
                if (XLEdgeAppState.Instance.RefreshSync)
                {
                    try
                    {
                        for (int c = lo.ListColumns.Count; c >= 1; c--)
                        {
                            string colName = string.Empty;
                            try { colName = lo.ListColumns[c].Name; }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read ListColumns[{c}].Name during RefreshSync cleanup - {ex.Message}");
                            }
                            var map = mappings.FirstOrDefault(m => string.Equals(m.Modified, colName, StringComparison.OrdinalIgnoreCase));
                            bool hasMapping = map.Modified != null;

                            bool hasFormula = false;
                            try
                            {
                                var firstRowObj = lo.DataBodyRange.Resize[1, lo.ListColumns.Count];
                                var firstCell = firstRowObj?.Cells[1, c] as Microsoft.Office.Interop.Excel.Range;
                                var formula = firstCell?.Formula as string;
                                if (!string.IsNullOrWhiteSpace(formula) && formula.StartsWith("=", StringComparison.Ordinal))
                                    hasFormula = true;
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogDebug($"{nameof(RefreshListObjectAsync)}: failed to read first-row formula for column {c} during RefreshSync cleanup - {ex.Message}");
                            }

                            if (!hasMapping && !hasFormula)
                            {
                                try { lo.ListColumns[c].Delete(); } catch (Exception ex) { LogUtility.LogException(ex, "Failed to delete column"); }
                            }
                        }
                    }
                    catch (Exception ex) { LogUtility.LogException(ex, "Failed to delete columns not present in mapping and not formula columns"); }
                }

                // --- STEP 13: Re-embed drilldown/attachment/image columns ---
                // Use meta data from CustomXMLParts (storedMetaJson)
                await SetRefreshMessage("Embedding attachment or drilldown links...", waitWindow, appOverlay, useWaitWindow);
                ReportMeta reportMetaForLinks = null;
                try
                {
                    if (!string.IsNullOrWhiteSpace(storedMetaJson))
                    {
                        reportMetaForLinks = JsonSerializer.Deserialize<ReportMeta>(storedMetaJson, JsonGlobals.Options);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "RefreshListObjectAsync: failed to parse stored report metadata for hyperlink/image re-embed");
                }

                // Re-embed drilldown/attachment/image columns
                if (reportMetaForLinks != null)
                {
                    try
                    {
                        AddDrilldownHyperlinks(sheet, lo, reportMetaForLinks);
                        AddAttachmentAndImageColumns(sheet, lo, reportMetaForLinks);
                        LogUtility.LogDebug($"RefreshListObjectAsync|Re-embedded drilldown/attachment/image columns");
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogException(ex, "RefreshListObjectAsync: failed to re-embed drilldown/attachment/image columns");
                    }
                }

                // --- STEP 14: Update params data in worksheet (ONLY if control sheet exists) ---
                await SetRefreshMessage("Updating param data...", waitWindow, appOverlay, useWaitWindow);
                try
                {
                    if (hasParamsPayload && !string.IsNullOrEmpty(paramsWithLabels))
                    {
                        await SetRefreshMessage("Updating report parameters...", waitWindow, appOverlay, useWaitWindow);

                        // Prefer the richly-merged array-shape params (preserves label/type/componentType
                        // and carries forward untouched parameters) built by
                        // AddinModule.BuildRefreshParamsPayload; falls back to the bare request-shape
                        // payload if that merge is unavailable.
                        string mergedParamsForDisplayAndStorage = !string.IsNullOrWhiteSpace(XLEdgeAppState.Instance.UpdatedParamData)
                            ? XLEdgeAppState.Instance.UpdatedParamData
                            : paramsWithLabels;

                        // --- Update parameter sheet cells (IT4, IU4, IV4, IW4 + Parameters Section rows) ---
                        UpdateParameterSheetCells(sheet, mergedParamsForDisplayAndStorage, lo);

                        // --- Save the merged (label/type-preserving, untouched-params-preserving) params to CustomXMLParts ---
                        SaveUpdatedReportData(wb, listObjectName, title, storedMetaJson, mergedParamsForDisplayAndStorage);

                        // Clear cached data since we've saved it
                        XLEdgeAppState.Instance.ClearCachedRefreshData();

                        LogUtility.LogDebug($"RefreshListObjectAsync|Updated parameter sheet cells, saved merged params to CustomXMLParts, and cleared cached refresh data for table: {listObjectName}");
                    }
                    else
                    {
                        // No control sheet payload: rewrite the Parameters Section / IT4-IW4 cells from
                        // the report's own stored parameter metadata, so a manually cleared or damaged
                        // section is restored on every refresh. Nothing needs re-saving to the
                        // CustomXMLPart here - storedParamsJson is already what's persisted.
                        if (!string.IsNullOrWhiteSpace(storedParamsJson))
                        {
                            UpdateParameterSheetCells(sheet, storedParamsJson, lo);
                            LogUtility.LogDebug($"RefreshListObjectAsync|No control sheet payload - restored Parameters Section from stored params JSON");
                        }
                        else
                        {
                            LogUtility.LogDebug($"RefreshListObjectAsync|No control sheet payload and no stored params JSON - skipping parameter update");
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "RefreshListObjectAsync|Failed to update params/save to CustomXMLParts");
                }

                // --- STEP 15: Cleanup ---
                await SetRefreshMessage("Cleaning up residual...", waitWindow, appOverlay, useWaitWindow);
                try
                {
                    if (useWaitWindow)
                    {
                        try { waitWindow?.RequestClose(); } catch (Exception ex) { LogUtility.LogException(ex, "Failed to close wait window"); }
                    }
                    else
                    {
                        if (appOverlay != null)
                        {
                            try
                            {
                                await System.Windows.Application.Current.Dispatcher
                                    .InvokeAsync(() => appOverlay.HideBusyAsync())
                                    .Task.Unwrap();
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogException(ex, "Failed to hide busy overlay");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "RefreshListObjectAsync: failed to clean up wait window/busy overlay after successful refresh");
                }

                // Reclaim keyboard focus from the WebView2 task pane back to Excel.
                await ReleaseKeyboardFocusFromTaskPaneAsync();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "RefreshListObjectAsync");

                if (collectErrors)
                {
                    throw;
                }

                if (useWaitWindow)
                {
                    await ShowErrorAsync(ex.Message, waitWindow);
                }
                else
                {
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() => appOverlay?.ShowError(ex.Message));
                }

                await ReleaseKeyboardFocusFromTaskPaneAsync();
            }
        }
        /// <summary>
        /// Cancellation cleanup
        /// </summary>
        private static async Task CancelCleanupAsync(XLEdgeWaitWindow waitWindow, AppOverlay appOverlay, bool useWaitWindow, bool collectErrors)
        {
            if (collectErrors)
            {
                throw new OperationCanceledException("Operation cancelled by user.");
            }

            if (useWaitWindow)
            {
                try { waitWindow?.RequestClose(); }
                catch (Exception ex) { LogUtility.LogException(ex, "CancelCleanupAsync|Failed to close wait window"); }
            }
            else
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (appOverlay != null) await appOverlay.HideBusyAsync();
                });
            }

            // Reclaim keyboard focus from the WebView2 task pane back to Excel.
            await ReleaseKeyboardFocusFromTaskPaneAsync();
        }
        /// <summary>
        /// Sets the refresh message on the wait window or busy overlay
        /// </summary>
        private static async Task SetRefreshMessage(string message, XLEdgeWaitWindow waitWindow, AppOverlay appOverlay, bool useWaitWindow)
        {
            try
            {
                if (useWaitWindow && waitWindow != null)
                {
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try { waitWindow.SetProcessMessage(message); }
                        catch (Exception ex) { LogUtility.LogDebug($"SetRefreshMessage|Failed to set wait window message: {ex.Message}"); }
                    });
                }
                else if (appOverlay != null)
                {
                    await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try { appOverlay.ShowBusyasyn(message); }
                        catch (Exception ex) { LogUtility.LogDebug($"SetRefreshMessage|Failed to set overlay message: {ex.Message}"); }
                    });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"SetRefreshMessage|Failed to set message: {ex.Message}");
            }
        }
        /// <summary>
        /// Saves updated report data to CustomXMLParts
        /// This method is called from RefreshListObjectAsync after successful refresh
        /// </summary>
        private static void SaveUpdatedReportData(Excel.Workbook workbook, string listObjectName, string title, string metaJson, string paramsJson)
        {
            try
            {
                if (workbook == null || string.IsNullOrEmpty(listObjectName))
                {
                    LogUtility.LogDebug("SaveUpdatedReportData|workbook or listObjectName is null");
                    return;
                }

                // Get column mappings for the table
                var mappings = GetColumnMappingsForTable(workbook, listObjectName);

                // Build and save custom XML
                string xml = BuildCustomXml(title, listObjectName, metaJson, paramsJson, mappings);
                SaveCustomXmlPart(workbook, xml, listObjectName);

                LogUtility.LogDebug($"SaveUpdatedReportData|Successfully saved data for {listObjectName}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "SaveUpdatedReportData");
            }
        }

        /// <summary>
        /// Gets column mappings for a table by name
        /// This is called from SaveUpdatedReportData when we only have the table name
        /// </summary>
        private static List<(string Original, string Modified, int RawIndex)> GetColumnMappingsForTable(Excel.Workbook workbook, string listObjectName)
        {
            var mappings = new List<(string Original, string Modified, int RawIndex)>();

            try
            {
                if (workbook == null || string.IsNullOrEmpty(listObjectName))
                {
                    LogUtility.LogDebug("GetColumnMappingsForTable|workbook or listObjectName is null");
                    return mappings;
                }

                Excel.ListObject lo = null;
                Excel.Sheets allSheets = workbook.Worksheets;
                try
                {
                    foreach (Excel.Worksheet ws in allSheets)
                    {
                        try
                        {
                            if (ws.ListObjects.Count > 0 && string.Equals(ws.ListObjects[1].Name, listObjectName, StringComparison.OrdinalIgnoreCase))
                            {
                                lo = ws.ListObjects[1];
                                break;
                            }
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

                if (lo?.HeaderRowRange != null)
                {
                    int col = 1;
                    foreach (Excel.Range headerCell in lo.HeaderRowRange.Cells)
                    {
                        string headerText = Convert.ToString(headerCell.Value) ?? string.Empty;
                        mappings.Add((headerText, headerText, col));
                        col++;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"GetColumnMappingsForTable: Failed to get column mappings - {ex.Message}");
            }

            return mappings;
        }
        /// <summary>
        /// Updates the parameter sheet's "Parameters Section:" rows and IT4/IU4/IV4/IW4/IT5 bookkeeping
        /// cells from a refresh payload, resolving the same target sheet/mode used at report creation
        /// and delegating the actual write to RewriteParameterSectionRows.
        /// </summary>
        private static void UpdateParameterSheetCells(Excel.Worksheet dataSheet, string paramsJson, Excel.ListObject tableObj)
        {
            const string MethodName = "UpdateParameterSheetCells";

            try
            {
                // Determine if we need to update a companion parameter sheet or same-sheet cells
                Excel.Worksheet paramSheet;
                bool sameSheetMode;

                if (tableObj.HeaderRowRange != null && tableObj.HeaderRowRange.Offset[1, 0].Row == 2)
                {
                    // Separate sheet mode - find companion parameter sheet
                    string paramSheetName = $"P_{dataSheet.Name}";
                    if (paramSheetName.Length >= 29)
                    {
                        paramSheetName = paramSheetName.Substring(0, 28);
                    }
                    paramSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableObj.Name);
                    sameSheetMode = false;
                }
                else
                {
                    // Same-sheet mode - use the data sheet itself
                    paramSheet = dataSheet;
                    sameSheetMode = true;
                }

                if (paramSheet == null)
                {
                    LogUtility.LogDebug($"{MethodName}|Parameter sheet not found");
                    return;
                }

                // The refresh payload (XLEdgeParamsBuilder.BuildParamData's output) is a JSON OBJECT
                // shape ({"reportId":..,"parameters":[...],"extraParameters":{...}}) - different from
                // the JSON ARRAY shape RewriteParameterSectionRows/ParseParamDisplayRows expect (the
                // same array-of-parameter-objects shape used at report-creation time). Adapt it first so
                // the shared writer sees the shape it already knows how to parse.
                string displayRowsJson = AdaptRefreshPayloadToDisplayRowsJson(paramsJson);

                RewriteParameterSectionRows(paramSheet, displayRowsJson, tableObj.Name, sameSheetMode);

                LogUtility.LogDebug($"{MethodName}|Successfully updated parameter sheet rows/cells");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Failed to update parameter sheet cells");
            }
        }

        /// <summary>
        /// Reshapes XLEdgeParamsBuilder.BuildParamData's refresh-payload JSON object
        /// ({"reportId":..,"parameters":[...],"extraParameters":{...}}) into the JSON array-of-
        /// parameter-objects shape ParseParamDisplayRows/RewriteParameterSectionRows expect (the same
        /// shape used at report-creation time, where each array entry may carry its own embedded
        /// "extraParameters"). Attaches the top-level extraParameters object onto the first parameter
        /// entry (synthesizing a placeholder entry if there are no parameters at all but extraParameters
        /// exist), which is exactly what ParseParamDisplayRows's own per-item scan already expects and
        /// handles - no changes needed there. If the input is already array-shaped (or isn't valid
        /// JSON), returns it unchanged so the caller degrades gracefully instead of throwing.
        /// </summary>
        private static string AdaptRefreshPayloadToDisplayRowsJson(string refreshPayloadJson)
        {
            if (string.IsNullOrWhiteSpace(refreshPayloadJson))
            {
                return refreshPayloadJson;
            }

            try
            {
                using JsonDocument doc = JsonDocument.Parse(refreshPayloadJson);
                if (doc.RootElement.ValueKind != JsonValueKind.Object)
                {
                    // Already array-shaped (or something else entirely) - pass through unchanged.
                    return refreshPayloadJson;
                }

                bool hasParameters = JsonHelper.TryGetProperty(doc.RootElement, "parameters", out JsonElement parametersEl)
                    && parametersEl.ValueKind == JsonValueKind.Array;

                bool hasExtraParams = JsonHelper.TryGetProperty(doc.RootElement, ExtraParametersKey, out JsonElement extraParamsEl)
                    && extraParamsEl.ValueKind == JsonValueKind.Object
                    && extraParamsEl.EnumerateObject().Any();

                var resultArray = new JsonArray();
                bool extraAttached = false;

                if (hasParameters)
                {
                    foreach (JsonElement item in parametersEl.EnumerateArray())
                    {
                        JsonNode itemNode = JsonNode.Parse(item.GetRawText());
                        if (!extraAttached && hasExtraParams && itemNode is JsonObject itemObj)
                        {
                            itemObj[ExtraParametersKey] = JsonNode.Parse(extraParamsEl.GetRawText());
                            extraAttached = true;
                        }

                        resultArray.Add(itemNode);
                    }
                }

                if (hasExtraParams && !extraAttached)
                {
                    var placeholder = new JsonObject
                    {
                        [ExtraParametersKey] = JsonNode.Parse(extraParamsEl.GetRawText())
                    };
                    resultArray.Add(placeholder);
                }

                return resultArray.ToJsonString();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(AdaptRefreshPayloadToDisplayRowsJson)}: failed to adapt refresh payload JSON to display-rows array shape - falling back to raw payload");
                return refreshPayloadJson;
            }
        }
        private static async Task HandleFailureAsync(string message, XLEdgeWaitWindow waitWindow, AppOverlay appOverlay, bool useWaitWindow, bool collectErrors)
        {
            if (collectErrors)
            {
                throw new InvalidOperationException(message);
            }

            if (useWaitWindow)
            {
                await ShowErrorAsync(message, waitWindow);
            }
            else
            {
                await System.Windows.Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (appOverlay != null)
                    {
                        await appOverlay.HideBusyAsync();
                        await appOverlay.ShowErrorAsync(message);
                    }
                });
            }

            // Reclaim keyboard focus from the WebView2 task pane back to Excel.
            await ReleaseKeyboardFocusFromTaskPaneAsync();
        }

        private static async Task ShowErrorAsync(string message, XLEdgeWaitWindow waitWindow)
        {
            await System.Windows.Application.Current.Dispatcher.InvokeAsync(() =>
            {
                try { waitWindow?.RequestClose(); } catch (Exception ex) { LogUtility.LogException(ex, "ShowErrorAsync: failed to close wait window"); }

                try
                {
                    var mw = new XLEdgeMessageWindow(message ?? "Error", System.Windows.Forms.MessageBoxIcon.Error);
                    mw.ShowDialog();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "ShowErrorAsync");
                }
            });
        }

        private static IEnumerable<List<string>> ParseCsv(string csv)
        {
            if (string.IsNullOrEmpty(csv)) yield break;

            using var reader = new StringReader(csv);
            string line;
            while ((line = reader.ReadLine()) != null)
            {
                var fields = new List<string>();
                var sb = new StringBuilder();
                bool inQuotes = false;
                // A while loop (rather than for) here so the escaped-quote branch's extra advance
                // doesn't read as mutating a for loop's own stop-condition variable - behavior is
                // unchanged: an escaped "" inside a quoted field advances by 2 (skipping both quote
                // characters), every other branch advances by 1.
                int i = 0;
                while (i < line.Length)
                {
                    char ch = line[i];
                    if (ch == '"')
                    {
                        if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i += 2;
                        }
                        else
                        {
                            inQuotes = !inQuotes;
                            i++;
                        }
                        continue;
                    }

                    if (ch == ',' && !inQuotes)
                    {
                        fields.Add(sb.ToString());
                        sb.Clear();
                        i++;
                        continue;
                    }

                    sb.Append(ch);
                    i++;
                }

                fields.Add(sb.ToString());
                yield return fields;
            }
        }
    }

    /// <summary>
    /// Ported from VB's FrmProcessBar_Load/FormProcessBar_Closing bracket
    /// </summary>
    internal sealed class ExcelBulkOperationScope : IDisposable
    {
        private bool _disposed;

        public ExcelBulkOperationScope()
        {
            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.RequireActiveExcelApplication();
                excelApp.EnableEvents = false;
                excelApp.ScreenUpdating = false;
                excelApp.DisplayAlerts = false;
                excelApp.Calculation = Excel.XlCalculation.xlCalculationManual;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to suspend Excel screen updating/events/alerts for report run");
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;

            try
            {
                Excel.Application excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                if (excelApp != null)
                {
                    excelApp.EnableEvents = true;
                    excelApp.ScreenUpdating = true;
                    excelApp.DisplayAlerts = true;
                    excelApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;

                    try
                    {
                        var activeCell = excelApp.ActiveCell;
                        if (activeCell != null)
                        {
                            excelApp.StatusBar = false;
                            System.Windows.Forms.Application.DoEvents();
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogError($"Failed to force Excel focus restoration - {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to restore Excel screen updating/events/alerts after report run");
            }

            try
            {
                XLEdge.AddinModule.ApplyRibbonState("ApplySheetActiveState");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to refresh ribbon state after report run");
            }
        }
    }
}