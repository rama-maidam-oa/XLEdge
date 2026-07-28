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
                await UiDispatcher.RunAsync(() =>
                {
                    _waitWindow.SetProcessMessage(message);
                    _waitWindow.RequestClose();
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
            extraParamsEl = default;
            extraParamsKey = null;

            foreach (JsonProperty prop in root.EnumerateObject())
            {
                if (prop.Name.Equals("extraParameters", StringComparison.OrdinalIgnoreCase))
                {
                    extraParamsEl = prop.Value;
                    extraParamsKey = prop.Name;
                    break;
                }
            }

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

            foreach (JsonProperty prop in extraParamsEl.EnumerateObject())
            {
                if (!displayKeys.Contains(prop.Name))
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
                    writer.WritePropertyName("extraParameters");
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

        public static async Task CreateReportFromTitleAsync(string title, AppOverlay appOverlay = null, bool useWaitWindow = false, string paramsJsonPayload = null)
        {
            using var excelBulkScope = new ExcelBulkOperationScope();

            _appOverlay = appOverlay;
            _ctsHelper = new CancellationHelper();

            bool isDrilldownRequest = !string.IsNullOrWhiteSpace(paramsJsonPayload);

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

            try
            {
                GetEdgeRequestFromTitle(title);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Failed to parse title for report generation: {title}");
                await DisplayErrorAsync($"Invalid title format for report generation. Title Format {title}");
                return;
            }

            string csvResponse = null;
            try
            {
                await SetMessage("Downloading report data...");
                string csvUrl = isDrilldownRequest
                    ? $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/runner?type=csv"
                    : $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/runner?runId={_edgeRequest.ReportRunId}&type=csv";

                // Strip display values from the payload before sending it to the CSV endpoint.
                string csvPayload = isDrilldownRequest ? StripExtraParameterDisplayValues(paramsJsonPayload) : null;

                csvResponse = await ApiHelper.ServerAPI(csvUrl, "JSON", csvPayload ?? string.Empty, "POST", _ctsHelper.GetToken());
                await Task.Delay(100);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("Report generation cancelled by user.");
                await ApiHelper.NotifyCancelRunAsync(XLEdgeAppState.Instance.LoginUrl, _edgeRequest?.ReportRunId);
                await DisplayErrorAsync("Report generation was cancelled by the user.");
                await CleanupAsync();
                return;
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Report generation request timed out");
                await DisplayErrorAsync("The request timed out. Please try again.");
                await CleanupAsync();
                return;
            }
            catch (Exception ex)
            {
                // Clean up the wait window/overlay and restore Excel focus on any unhandled error.
                LogUtility.LogException(ex, "Unhandled error in report generation");
                await DisplayErrorAsync($"An unexpected error occurred during report generation.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return;
            }

            if (string.IsNullOrWhiteSpace(csvResponse))
            {
                LogUtility.LogWarn("CSV response is empty. Cannot generate report.");
                await DisplayErrorAsync("Failed to download report data. The response was empty.");
                await CleanupAsync();
                return;
            }
            else
            {
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
                    return;
                }
            }

            // Reuse the already-supplied params payload directly instead of calling the API again.
            string paramsResponse = null;

            if (!string.IsNullOrEmpty(paramsJsonPayload))
            {
                // We already have the parameters from the control sheet - use it directly
                paramsResponse = paramsJsonPayload;
                LogUtility.LogDebug($"CreateReportFromTitleAsync|Using control-sheet params payload for report {_edgeRequest.ReportId}");
            }

            // Step-4: Fetch report definition (Meta) - always need this from API
            await SetMessage("Fetching report definition...");
            string metaUrl = $"{XLEdgeAppState.Instance.LoginUrl.TrimEnd('/')}/rest/secure/report/report-definition?reportId={_edgeRequest.ReportId}&runId={_edgeRequest.ReportRunId}&isDrillDown={(isDrilldownRequest ? "true" : "false")}";
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
                return;
            }
            catch (ApiTimeoutException ex)
            {
                LogUtility.LogException(ex, "Report definition fetch timed out");
                await DisplayErrorAsync("The request timed out. Please try again.");
                await CleanupAsync();
                return;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unhandled error fetching report definition");
                await DisplayErrorAsync($"An unexpected error occurred while fetching report definition.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return;
            }

            if (string.IsNullOrWhiteSpace(metaResponse))
            {
                LogUtility.LogWarn("Report definition response is empty. Cannot generate report.");
                await DisplayErrorAsync("Failed to fetch report definition. The response was empty.");
                await CleanupAsync();
                return;
            }

            LogResponsePayload("Report definition response (metaResponse)", metaResponse);

            // Only fetch params from the API if we don't already have them.
            if (string.IsNullOrEmpty(paramsResponse))
            {
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
                    return;
                }
                catch (ApiTimeoutException ex)
                {
                    LogUtility.LogException(ex, "Report parameters fetch timed out");
                    await DisplayErrorAsync("The request timed out. Please try again.");
                    await CleanupAsync();
                    return;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Unhandled error fetching report parameters");
                    await DisplayErrorAsync($"An unexpected error occurred while fetching report parameters.{Environment.NewLine}{ex.Message}");
                    await CleanupAsync();
                    return;
                }

                if (string.IsNullOrWhiteSpace(paramsResponse))
                {
                    LogUtility.LogWarn("Report parameters response is empty. Cannot generate report.");
                    await DisplayErrorAsync("Failed to fetch report parameters. The response was empty.");
                    await CleanupAsync();
                    return;
                }

                LogResponsePayload("Report parameters response (paramsResponse)", paramsResponse);
            }
            else
            {
                // Already logged above (line ~337) when the control-sheet payload was first accepted.
            }

            ReportMeta reportMeta;
            try
            {
                reportMeta = JsonSerializer.Deserialize<ReportMeta>(metaResponse, JsonGlobals.Options);
                if (reportMeta == null)
                {
                    await DisplayErrorAsync("Report definition could not be parsed.");
                    await CleanupAsync();
                    return;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to parse report definition JSON");
                await DisplayErrorAsync("Report definition is not in the expected format.");
                await CleanupAsync();
                return;
            }

            try
            {
                await SetMessage("Building report in Excel...");
                BuildReportTable(_edgeRequest, reportMeta, csvResponse, metaResponse, paramsResponse, title);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to build report table in Excel");
                await DisplayErrorAsync($"Failed to write the report into Excel.{Environment.NewLine}{ex.Message}");
                await CleanupAsync();
                return;
            }

            await CleanupAsync();
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
                await DisplayErrorAsync("The request timed out. Please try again.");
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

            string tableId = $"ORB_{request.ReportId}_{request.ReportRunId}_E";

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

            bool sameSheet = XLEdgeAppState.Instance.ParamDataSameSheet;
            int headerRow = sameSheet ? 8 : 1;
            int dataStartRow = headerRow + 1;
            string companionSheetToDelete = null;

            Excel.Worksheet sheet = FindSheetWithTable(workbook, tableId);
            if (sheet != null)
            {
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
                }
            }
            else
            {
                string sheetName = BuildSheetName(reportMeta);
                if (ExcelSheetHelper.SheetExists(sheetName, workbook))
                {
                    sheet = (Excel.Worksheet)workbook.Worksheets[sheetName];
                    sheet.Cells.Clear();
                    ResetLeftoverRowArtifacts(sheet);
                }
                else
                {
                    try
                    {
                        sheet = (Excel.Worksheet)workbook.Worksheets.Add();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(BuildReportTable)}: default Worksheets.Add() failed, falling back to append-at-end - {ex.Message}");
                        sheet = (Excel.Worksheet)workbook.Worksheets.Add(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
                    }
                    sheet.Name = sheetName;
                }
            }

            if (sameSheet && string.IsNullOrEmpty(companionSheetToDelete))
            {
                try
                {
                    string paramSheetName = $"P_{sheet.Name}";
                    if (paramSheetName.Length >= 29)
                    {
                        paramSheetName = paramSheetName.Substring(0, 28);
                    }

                    Excel.Worksheet oldParamSheet = ExcelSheetHelper.GetParameterSheet(paramSheetName, tableId);
                    if (oldParamSheet != null)
                    {
                        companionSheetToDelete = oldParamSheet.Name;
                        Marshal.ReleaseComObject(oldParamSheet);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to resolve orphaned companion parameter sheet before same-sheet write");
                }
            }

            sheet.Activate();

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

                Excel.Range startCell = (Excel.Range)sheet.Cells[dataStartRow, 1];
                Excel.Range writeRange = startCell.Resize[dataRowCount, mappings.Count];
                writeRange.Value2 = writeArr;
            }

            Excel.Range tableRange = sheet.Range[sheet.Cells[headerRow, 1], sheet.Cells[headerRow + rowsToReserve, mappings.Count]];
            Excel.ListObject listObject = sheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
            listObject.Name = tableId;
            listObject.TableStyle = "TableStyleLight9";

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

            string reportTitleText = XLEdgeValueFormatter.RemoveEquaSymbol(
                (XLEdgeAppState.Instance.FollowDrilldown && !string.IsNullOrWhiteSpace(XLEdgeAppState.Instance.ChildRptLabel))
                    ? XLEdgeAppState.Instance.ChildRptLabel
                    : (reportMeta.Name ?? request.ReportName));

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

            AddDrilldownHyperlinks(sheet, listObject, reportMeta);
            AddAttachmentAndImageColumns(sheet, listObject, reportMeta);

            try
            {
                string xml = BuildCustomXml(title, tableId, metaJson, paramsJson, mappings);
                SaveCustomXmlPart(workbook, xml, tableId);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to persist report metadata");
            }

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

            if (!string.IsNullOrEmpty(companionSheetToDelete))
            {
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
        private static void RewriteParameterSectionRows(Excel.Worksheet paramSheet, string paramsJson, string tableId, bool sameSheetMode)
        {
            List<(string Label, string ValueText)> paramRows = ParseParamDisplayRows(
                paramsJson, out string oracleRespId, out string oracleRespValue, out string segmentValues, out string segmentDisplayValues);

            if (sameSheetMode)
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
                    LogUtility.LogDebug($"{nameof(RewriteParameterSectionRows)}: failed to clear stale same-sheet parameter rows before rewrite - {ex.Message}");
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
            else
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
                    LogUtility.LogDebug($"{nameof(RewriteParameterSectionRows)}: failed to clear stale companion-sheet parameter rows before rewrite - {ex.Message}");
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

            // Each of IT4/IU4/IV4/IW4 is cleared first, then only re-populated if it has a value -
            // ensures a blank value this round leaves an actually-blank cell rather than stale content.
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

            // IW4 holds the segment display value, alongside IV4's raw segment value.
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
                    try
                    {
                        if (!JsonHelper.TryGetProperty(item, "extraParameters", out JsonElement extraEl) ||
                            extraEl.ValueKind != JsonValueKind.Object)
                        {
                            continue;
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

                foreach (JsonElement item in doc.RootElement.EnumerateArray())
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
                            continue;
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
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ParseParamDisplayRows));
            }

            return result;
        }

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
                            // Accepts either a real JSON object, or a string whose content itself
                            // parses as a JSON object (e.g. {"Company":"1000-5000","Department":"-",...}).
                            JsonElement? segmentObjectEl = null;
                            if (prop.Value.ValueKind == JsonValueKind.Object)
                            {
                                segmentObjectEl = prop.Value;
                            }
                            else if (prop.Value.ValueKind == JsonValueKind.String)
                            {
                                string rawText = prop.Value.GetString();
                                if (!string.IsNullOrWhiteSpace(rawText))
                                {
                                    try
                                    {
                                        using var innerDoc = JsonDocument.Parse(rawText);
                                        if (innerDoc.RootElement.ValueKind == JsonValueKind.Object)
                                        {
                                            segmentObjectEl = innerDoc.RootElement.Clone();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        LogUtility.LogDebug($"{nameof(ExtractExtraParams)}: ORACLE_GL_SEGMENT_DISPLAY_VALUES string value did not parse as a JSON object - {ex.Message}");
                                    }
                                }
                            }

                            if (segmentObjectEl.HasValue)
                            {
                                var segmentString = new StringBuilder();
                                foreach (JsonProperty innerProp in segmentObjectEl.Value.EnumerateObject())
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

                                glSegments = segmentString.ToString();
                            }
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
        private static string ComputeRawParamDisplayValue(JsonElement item, string componentType, string paramOperator, string paramType)
        {
            string paramValue;

            bool hasAnyProperty = item.ValueKind == JsonValueKind.Object && item.EnumerateObject().Any();

            if (!hasAnyProperty)
            {
                paramValue = string.Empty;
            }
            else if (paramOperator == null || paramType == null)
            {
                paramValue = string.Empty;
            }
            else if (paramOperator.Contains("NULL"))
            {
                paramValue = string.Empty;
            }
            else if (JsonHelper.TryGetProperty(item, "displayValue", out JsonElement dvEl) && dvEl.ValueKind != JsonValueKind.Null && dvEl.ValueKind != JsonValueKind.Undefined)
            {
                if (dvEl.ValueKind == JsonValueKind.Array)
                {
                    List<string> items = dvEl.EnumerateArray().Select(v => v.ToString()).ToList();
                    paramValue = items.Count > 0
                        ? string.Join(",", items.Select(v => JoinFormatted(v, paramType)))
                        : string.Empty;
                }
                else if (dvEl.ValueKind == JsonValueKind.Object)
                {
                    LogUtility.LogWarn($"Type of jToken as object is not handled yet. {dvEl}");
                    paramValue = string.Empty;
                }
                else
                {
                    paramValue = Convert.ToString(XLEdgeValueFormatter.FormatValue(dvEl.ToString(), paramType));
                }
            }
            else if (JsonHelper.TryGetProperty(item, "displayValues", out JsonElement dvsEl) && dvsEl.ValueKind == JsonValueKind.Array)
            {
                List<JsonElement> values = dvsEl.EnumerateArray().ToList();

                if (values.Count == 0)
                {
                    paramValue = string.Empty;
                }
                else if ((componentType != null && componentType.Contains("range")) ||
                         paramOperator == "BETWEEN" || paramOperator == "NOT BETWEEN")
                {
                    if (values.Count == 2)
                    {
                        paramValue = $"{XLEdgeValueFormatter.FormatValue(values[0].ToString(), paramType)} and {XLEdgeValueFormatter.FormatValue(values[1].ToString(), paramType)}";
                    }
                    else if (values.Count == 1)
                    {
                        paramValue = Convert.ToString(XLEdgeValueFormatter.FormatValue(values[0].ToString(), paramType));
                    }
                    else
                    {
                        paramValue = string.Empty;
                    }
                }
                else
                {
                    paramValue = string.Join(",", values.Select(v => JoinFormatted(v.ToString(), paramType)));
                }
            }
            else
            {
                paramValue = string.Empty;
            }

            return paramValue;
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

            return TryResolveReportXmlForRefresh(workbook, listObjectName, listObject, out title, out _, out _, out metaJson, out paramsJson, out _);
        }

        private static bool TryResolveReportXmlForRefresh(
            Excel.Workbook workbook,
            string listObjectName,
            Excel.ListObject listObject,
            out string title,
            out string reportId,
            out string runId,
            out string metaJson,
            out string paramsJson,
            out List<(string Original, string Modified, int RawIndex)> mappings)
        {
            title = null;
            reportId = null;
            runId = null;
            metaJson = null;
            paramsJson = null;
            mappings = new List<(string Original, string Modified, int RawIndex)>();

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
                        string xml = part.XML;
                        if (string.IsNullOrWhiteSpace(xml))
                        {
                            continue;
                        }

                        if (xml.Contains($"<ListObjectName>{listObjectName}</ListObjectName>"))
                        {
                            XDocument xdoc = XDocument.Parse(xml);
                            title = xdoc.Root?.Element("Title")?.Value ?? string.Empty;
                            metaJson = xdoc.Root?.Element("Meta")?.Value ?? string.Empty;
                            paramsJson = xdoc.Root?.Element("Params")?.Value ?? string.Empty;

                            string[] titleParts = title.Split('|');
                            if (titleParts.Length < 3)
                            {
                                continue;
                            }

                            reportId = titleParts[1];
                            runId = titleParts[2];

                            XElement colsElem = xdoc.Root?.Element("Columns");
                            if (colsElem != null)
                            {
                                foreach (XElement ce in colsElem.Elements("Column"))
                                {
                                    string orig = ce.Attribute("original")?.Value ?? string.Empty;
                                    string mod = ce.Attribute("modified")?.Value ?? string.Empty;
                                    int.TryParse(ce.Attribute("rawIndex")?.Value ?? "0", out int idx);
                                    mappings.Add((orig, mod, idx));
                                }
                            }

                            return true;
                        }

                        if (xml.IndexOf("<DataMeta>", StringComparison.OrdinalIgnoreCase) < 0 ||
                            xml.IndexOf(listObjectName, StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        XDocument legacyDoc;
                        try
                        {
                            legacyDoc = XDocument.Parse(xml);
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogException(ex, "TryResolveReportXmlForRefresh: failed to parse a legacy CustomXMLPart");
                            continue;
                        }

                        XElement dataElem = legacyDoc.Root?.Elements().FirstOrDefault(e => e.Name.LocalName == "Data");
                        if (dataElem == null)
                        {
                            continue;
                        }

                        string infoId = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "InfoID")?.Value ?? string.Empty;
                        if (!string.Equals(infoId, listObjectName, StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        Match tableNameMatch = Regex.Match(listObjectName, @"^ORB_(?<reportId>[^_]+)_(?<runId>[^_]+)_[EP]$", RegexOptions.IgnoreCase);
                        if (!tableNameMatch.Success)
                        {
                            LogUtility.LogWarn($"TryResolveReportXmlForRefresh|Legacy metadata found for '{listObjectName}' but its name doesn't match the expected ORB_<reportId>_<runId>_E/P pattern - cannot derive report/run id.");
                            continue;
                        }

                        metaJson = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "DataMeta")?.Value ?? string.Empty;
                        paramsJson = dataElem.Elements().FirstOrDefault(e => e.Name.LocalName == "DataParam")?.Value ?? string.Empty;
                        reportId = tableNameMatch.Groups["reportId"].Value;
                        runId = tableNameMatch.Groups["runId"].Value;
                        title = $"Edge|{reportId}|{runId}|{listObjectName}";

                        if (listObject?.HeaderRowRange != null)
                        {
                            int col = 1;
                            foreach (Excel.Range headerCell in listObject.HeaderRowRange.Cells)
                            {
                                string headerText = Convert.ToString(headerCell.Value) ?? string.Empty;
                                mappings.Add((headerText, headerText, col));
                                col++;
                            }
                        }

                        return true;
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
                baseName = "Column";
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

        private static void AddDrilldownHyperlinks(Excel.Worksheet sheet, Excel.ListObject listObject, ReportMeta reportMeta)
        {
            if (reportMeta.Drilldowns == null || reportMeta.Drilldowns.Length == 0 || listObject.DataBodyRange == null)
            {
                return;
            }

            const int maxHyperlinks = 65530;
            int hyperlinkCount = 0;

            var byColumn = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (RptDrilldown dd in reportMeta.Drilldowns)
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

                list.Add($"DRILLDOWN|{dd.DrillReportId}|{dd.DrillReportName}|{reportMeta.ReportId}");
            }

            foreach (KeyValuePair<string, List<string>> kvp in byColumn)
            {
                int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, kvp.Key);
                if (matchCol <= 0)
                {
                    continue;
                }

                string tooltip = string.Join(",", kvp.Value);
                if (tooltip.Length > 255)
                {
                    tooltip = tooltip.Substring(0, 250) + "...";
                }

                Excel.Range dataRange = listObject.DataBodyRange;
                for (int r = 1; r <= dataRange.Rows.Count; r++)
                {
                    if (hyperlinkCount >= maxHyperlinks)
                    {
                        LogUtility.LogWarn($"Reached maximum hyperlink limit of {maxHyperlinks}; stopping further drilldown hyperlinks.");
                        return;
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
                }
            }
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
                        hyperlinkCount = AddHyperlinkColumn(sheet, listObject, col, hyperlinkCount, maxHyperlinks);
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
            }

            return hyperlinkCount;
        }

        private static int AddHyperlinkColumn(Excel.Worksheet sheet, Excel.ListObject listObject, RptColumn col, int hyperlinkCount, int maxHyperlinks)
        {
            int matchCol = ExcelSheetHelper.HRMatch(listObject.HeaderRowRange, col.Name?.Trim() ?? string.Empty);
            if (matchCol <= 0)
            {
                return hyperlinkCount;
            }

            Excel.Range dataRange = listObject.DataBodyRange;
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
            }

            return hyperlinkCount;
        }

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

            for (int r = 1; r <= dataRange.Rows.Count; r++)
            {
                Excel.Range cell = (Excel.Range)dataRange.Cells[r, matchCol];
                string destinationPath = null;
                try
                {
                    object rawValue = cell.Value;
                    if (rawValue == null)
                    {
                        continue;
                    }

                    string url = Convert.ToString(rawValue);
                    cell.Clear();

                    if (string.IsNullOrWhiteSpace(url))
                    {
                        continue;
                    }

                    string fileName = url.Contains("/") ? url.Substring(url.LastIndexOf('/') + 1) : url;
                    foreach (char invalidChar in Path.GetInvalidFileNameChars())
                    {
                        fileName = fileName.Replace(invalidChar, '_');
                    }

                    string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                    destinationPath = Path.Combine(downloadsFolder, fileName);

                    bool downloaded = ImageDownloadHelper.TryDownloadImage(url, destinationPath);
                    if (!downloaded || !File.Exists(destinationPath))
                    {
                        continue;
                    }

                    Excel.Shape imgShape = sheet.Shapes.AddPicture(
                        destinationPath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue,
                        (float)cell.Left, (float)cell.Top, (float)imgHeight, (float)imgWidth);

                    int rowIndex = cell.Row;
                    int colIndex = cell.Column;

                    double actualRowHeight = Math.Min(imgShape.Height, 409);
                    if (!rowMaxHeights.TryGetValue(rowIndex, out double existingRowHeight) || actualRowHeight > existingRowHeight)
                    {
                        rowMaxHeights[rowIndex] = actualRowHeight;
                        cell.EntireRow.RowHeight = actualRowHeight;
                    }

                    double colWidthEstimate = imgShape.Width / 10.0;
                    double adjustedColWidth = colWidthEstimate + (colWidthEstimate - 1);
                    if (!colMaxWidths.TryGetValue(colIndex, out double existingColWidth) || adjustedColWidth > existingColWidth)
                    {
                        colMaxWidths[colIndex] = adjustedColWidth;
                        cell.EntireColumn.ColumnWidth = adjustedColWidth;
                    }

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
                }
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
                    throw new Exception("Excel instance not available.");

                var sheet = excelApp.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                if (sheet == null)
                    throw new Exception("No active worksheet.");

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

                        // The Sleep intervals below run on a background thread on purpose: Excel's
                        // own STA/message-pump thread must stay free to actually process the
                        // synthetic SendKeys input between each step. Only the Excel COM calls
                        // themselves are marshalled onto that STA thread via UiDispatcher.Run -
                        // calling them directly from this background thread would access the
                        // Excel.Application RCW from the wrong apartment.
                        await Task.Run(() =>
                        {
                            try
                            {
                                Thread.Sleep(50);
                                UiDispatcher.Run(() => excelApp.SendKeys("{F2}"));
                                Thread.Sleep(30);
                                UiDispatcher.Run(() => excelApp.SendKeys("{ESC}"));
                                Thread.Sleep(30);

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
                        new XElement("Column",
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
                if (excelApp == null) throw new Exception("Excel instance not available.");

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
                if (sheet == null) throw new Exception("No active worksheet.");

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
                if (!TryResolveReportXmlForRefresh(wb, listObjectName, lo, out string title, out string reportId, out string runId, out string storedMetaJson, out string storedParamsJson, out List<(string Original, string Modified, int RawIndex)> mappings))
                {
                    await HandleFailureAsync("No metadata found for this table.", waitWindow, appOverlay, useWaitWindow, collectErrors);
                    return;
                }

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
                    await HandleFailureAsync("The request timed out. Please try again.", waitWindow, appOverlay, useWaitWindow, collectErrors);
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
                        if (string.IsNullOrWhiteSpace(baseName)) baseName = "Column" + i;

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
                    try { lo.ListRows.Add(); currentRows = lo.DataBodyRange.Rows.Count; }
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
                        if (firstRowFormulas.TryGetValue(c, out var f))
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
        /// Gets column mappings from the table's header row
        /// This is called from BuildCustomXml when we have a ListObject reference
        /// </summary>
        private static List<(string Original, string Modified, int RawIndex)> GetColumnMappings(Excel.ListObject tableObj)
        {
            var mappings = new List<(string Original, string Modified, int RawIndex)>();

            try
            {
                if (tableObj?.HeaderRowRange == null)
                {
                    return mappings;
                }

                int col = 1;
                foreach (Excel.Range headerCell in tableObj.HeaderRowRange.Cells)
                {
                    string headerText = Convert.ToString(headerCell.Value) ?? string.Empty;
                    mappings.Add((headerText, headerText, col));
                    col++;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "GetColumnMappings");
            }

            return mappings;
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

                bool hasExtraParams = JsonHelper.TryGetProperty(doc.RootElement, "extraParameters", out JsonElement extraParamsEl)
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
                            itemObj["extraParameters"] = JsonNode.Parse(extraParamsEl.GetRawText());
                            extraAttached = true;
                        }

                        resultArray.Add(itemNode);
                    }
                }

                if (hasExtraParams && !extraAttached)
                {
                    var placeholder = new JsonObject
                    {
                        ["extraParameters"] = JsonNode.Parse(extraParamsEl.GetRawText())
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
                throw new Exception(message);
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
                        appOverlay.ShowError(message);
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
                for (int i = 0; i < line.Length; i++)
                {
                    char ch = line[i];
                    if (ch == '"')
                    {
                        if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                        {
                            sb.Append('"');
                            i++;
                        }
                        else
                        {
                            inQuotes = !inQuotes;
                        }
                        continue;
                    }

                    if (ch == ',' && !inQuotes)
                    {
                        fields.Add(sb.ToString());
                        sb.Clear();
                        continue;
                    }

                    sb.Append(ch);
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