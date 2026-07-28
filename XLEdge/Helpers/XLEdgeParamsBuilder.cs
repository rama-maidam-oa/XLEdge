using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using XLEdge.Models;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Ported from XLEdgeParamsData.vb. Builds the parameter payload from control sheet changes
    /// and updates the original parameter data with the changed values.
    /// </summary>
    public static class XLEdgeParamsBuilder
    {
        /// <summary>
        /// Builds both parameter JSON shapes needed for a report: RequestJson (the compact
        /// {reportId, parameters, extraParameters} payload sent to the CSV API) and MergedJson
        /// (the full array-of-parameter-objects shape, with original label/type/componentType
        /// preserved, edited values merged in, and untouched parameters carried forward unchanged -
        /// used for anything that displays or persists "the current parameters").
        /// </summary>
        public static (string RequestJson, string MergedJson) BuildParamData(Excel.Workbook workbook, Excel.Worksheet paramSheet, Excel.ListObject tableObj)
        {
            if (workbook == null || tableObj == null)
            {
                return (string.Empty, string.Empty);
            }

            if (!ReportGenerator.TryGetStoredReportXml(workbook, tableObj.Name, out _, out _, out string paramsJson) ||
                string.IsNullOrWhiteSpace(paramsJson))
            {
                LogUtility.LogWarn($"BuildParamData|No stored parameter metadata found for table '{tableObj.Name}'.");
                return (string.Empty, string.Empty);
            }

            Excel.ListObject controlTable = ParamsControlSheetBuilder.FindControlTable(workbook);
            if (controlTable?.DataBodyRange == null)
            {
                LogUtility.LogDebug("BuildParamData|Parameters control sheet does not exist.");
                return (string.Empty, string.Empty);
            }

            string[] tableNameParts = tableObj.Name.Split('_');
            if (tableNameParts.Length < 2 || string.IsNullOrWhiteSpace(tableNameParts[1]))
            {
                return (string.Empty, string.Empty);
            }

            string reportId = tableNameParts[1];
            List<ControlSheetRow> controlRows = ReadControlSheetRows(controlTable);

            return BuildJsonPayload(paramSheet, reportId, paramsJson, controlRows);
        }

        private class ControlSheetRow
        {
            public string ReportId { get; set; }
            public string ParameterType { get; set; }
            public string ParameterName { get; set; }
            public string ParameterDisplayName { get; set; }
            public string IsRequired { get; set; }
            public string DataType { get; set; }
            public string Operator { get; set; }
            public string Value1 { get; set; }       // RAW value from Column J
            public string Value2 { get; set; }       // RAW value from Column K
            public string DisplayValue { get; set; } // DISPLAY value from Column IA (235)
        }

        private static List<ControlSheetRow> ReadControlSheetRows(Excel.ListObject controlTable)
        {
            var result = new List<ControlSheetRow>();

            Excel.Range headerRange = controlTable.HeaderRowRange;
            var headers = new List<string>();
            foreach (Excel.Range cell in headerRange.Cells)
            {
                headers.Add(Convert.ToString(cell.Value) ?? string.Empty);
            }

            Excel.Worksheet controlSheet = controlTable.Parent as Excel.Worksheet;
            if (controlSheet == null)
            {
                LogUtility.LogDebug("ReadControlSheetRows|Control sheet parent is null");
                return result;
            }

            int rowNumber = 4;
            foreach (Excel.Range row in controlTable.DataBodyRange.Rows)
            {
                var rowData = new ControlSheetRow();

                for (int colIndex = 1; colIndex <= headers.Count; colIndex++)
                {
                    object cellValue = ((Excel.Range)row.Cells[1, colIndex]).Value;
                    string value = cellValue != null ? Convert.ToString(cellValue) : string.Empty;

                    string header = headers[colIndex - 1];
                    switch (header)
                    {
                        case "Report ID":
                            rowData.ReportId = value;
                            break;
                        case "Parameter Type":
                            rowData.ParameterType = value;
                            break;
                        case "Parameter Name":
                            rowData.ParameterName = value;
                            break;
                        case "Parameter Display Name":
                            rowData.ParameterDisplayName = value;
                            break;
                        case "Is Required":
                            rowData.IsRequired = value;
                            break;
                        case "Data Type":
                            rowData.DataType = value;
                            break;
                        case "Operator":
                            rowData.Operator = value;
                            break;
                        case "Value1":
                            rowData.Value1 = value;
                            break;
                        case "Value2":
                            rowData.Value2 = value;
                            break;
                    }
                }

                // Read display value from column IA (235)
                try
                {
                    Excel.Range displayValueCell = controlSheet.Cells[rowNumber, 235] as Excel.Range;
                    if (displayValueCell != null && displayValueCell.Value != null)
                    {
                        rowData.DisplayValue = Convert.ToString(displayValueCell.Value);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"Failed to read display value from column IA for row {rowNumber}: {ex.Message}");
                }

                result.Add(rowData);
                rowNumber++;
            }

            return result;
        }

        private static (string RequestJson, string MergedJson) BuildJsonPayload(Excel.Worksheet paramSheet, string reportId, string paramsJson, List<ControlSheetRow> controlRows)
        {
            const string MethodName = "XLEdgeParamsBuilder.BuildJsonPayload";
            LogUtility.LogDebug($"{MethodName}|Building parameter payload for reportId: {reportId}");

            if (string.IsNullOrWhiteSpace(paramsJson) || controlRows == null || controlRows.Count == 0)
            {
                return (string.Empty, string.Empty);
            }

            JsonElement paramMapping;
            try
            {
                using var doc = JsonDocument.Parse(paramsJson);
                paramMapping = doc.RootElement.Clone();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Invalid JSON format in paramsJson");
                return (string.Empty, string.Empty);
            }

            if (paramMapping.ValueKind != JsonValueKind.Array)
            {
                return (string.Empty, string.Empty);
            }

            var columnMappings = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
            foreach (JsonElement item in paramMapping.EnumerateArray())
            {
                if (JsonHelper.TryGetProperty(item, "label", out JsonElement labelEl) && labelEl.ValueKind != JsonValueKind.Null)
                {
                    columnMappings[labelEl.ToString()] = item;
                }
                if (JsonHelper.TryGetProperty(item, "name", out JsonElement nameEl) && nameEl.ValueKind != JsonValueKind.Null)
                {
                    columnMappings[nameEl.ToString()] = item;
                }
            }

            var paramDataTypes = new Dictionary<string, string>();
            var parameters = new List<ReportParameterValue>();

            // Build extraParameters from control sheet rows
            var extraParams = new Dictionary<string, object>();
            var processedExtraParams = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // First pass: process regular parameters
            foreach (var row in controlRows)
            {
                if (row == null) continue;

                if (!string.IsNullOrEmpty(row.ReportId) && row.ReportId != reportId)
                {
                    continue;
                }

                string paramType = row.ParameterType ?? string.Empty;

                if (paramType.Equals("extraParameters", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string paramColumn = row.ParameterName ?? string.Empty;
                string paramOperatorLabel = row.Operator ?? string.Empty;
                string value1 = row.Value1 ?? string.Empty;
                string value2 = row.Value2 ?? string.Empty;

                if (string.IsNullOrEmpty(paramColumn) || string.IsNullOrEmpty(paramOperatorLabel))
                {
                    continue;
                }

                string updatedOperator = XLEdgeOperatorMappings.Map.TryGetValue(
                    paramOperatorLabel.Replace("\"", string.Empty).ToLowerInvariant(), out string mapped)
                    ? mapped
                    : string.Empty;

                if (string.IsNullOrWhiteSpace(updatedOperator) && string.IsNullOrEmpty(value1) && string.IsNullOrEmpty(value2))
                {
                    continue;
                }

                if (updatedOperator == "IS NULL" || updatedOperator == "IS NOT NULL")
                {
                    value1 = string.Empty;
                    value2 = string.Empty;
                }

                if (!columnMappings.TryGetValue(paramColumn, out JsonElement mapping) ||
                    !JsonHelper.TryGetProperty(mapping, "name", out JsonElement nameEl))
                {
                    LogUtility.LogDebug($"{MethodName}|No mapping found for column: {paramColumn}");
                    continue;
                }

                string paramName = nameEl.ToString();
                string paramTypeFromMapping = JsonHelper.TryGetProperty(mapping, "type", out JsonElement typeEl) ? typeEl.ToString() : string.Empty;

                paramDataTypes[paramName] = paramTypeFromMapping;

                var values = new List<object>();

                if (!string.IsNullOrEmpty(value1))
                {
                    object formatted = FormatValue(value1, paramTypeFromMapping);

                    if (updatedOperator == "IN" || updatedOperator == "NOT IN")
                    {
                        if (paramTypeFromMapping.ToUpper() != "DATE" && paramTypeFromMapping.ToUpper() != "DATETIME" &&
                            formatted?.ToString()?.Contains(",") == true)
                        {
                            var valuesList = SplitRespectingQuotes(formatted.ToString());
                            foreach (var item in valuesList)
                            {
                                values.Add(FormatValue(item, paramTypeFromMapping));
                            }
                        }
                        else
                        {
                            values.Add(formatted);
                        }
                    }
                    else
                    {
                        values.Add(formatted);
                    }
                }

                if (!string.IsNullOrEmpty(value2) && (updatedOperator == "BETWEEN" || updatedOperator == "NOT BETWEEN"))
                {
                    values.Add(FormatValue(value2, paramTypeFromMapping));
                }

                var paramValue = new ReportParameterValue
                {
                    Name = paramName,
                    Operator = updatedOperator
                };

                if (updatedOperator == "IS NULL" || updatedOperator == "IS NOT NULL")
                {
                    paramValue.Value = null;
                }
                else if (updatedOperator == "BETWEEN" || updatedOperator == "NOT BETWEEN" ||
                         updatedOperator == "IN" || updatedOperator == "NOT IN")
                {
                    paramValue.Values = values.Count > 0 ? values : null;
                }
                else
                {
                    paramValue.Value = values.Count > 0 ? values[0] : null;
                }

                parameters.Add(paramValue);
            }

            // Second pass: process extraParameters
            foreach (var row in controlRows)
            {
                if (row == null) continue;

                if (!string.IsNullOrEmpty(row.ReportId) && row.ReportId != reportId)
                {
                    continue;
                }

                string paramType = row.ParameterType ?? string.Empty;

                if (paramType.Equals("extraParameters", StringComparison.OrdinalIgnoreCase))
                {
                    string extraParamName = row.ParameterName ?? string.Empty;

                    // RAW value comes from Value1; DISPLAY value comes from DisplayValue (column IA).
                    string rawValue = row.Value1 ?? string.Empty;
                    string displayValue = row.DisplayValue ?? string.Empty;

                    if (!string.IsNullOrEmpty(extraParamName) &&
                        !string.IsNullOrEmpty(rawValue) &&
                        !string.Equals(rawValue, "N/A", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!processedExtraParams.Contains(extraParamName))
                        {
                            processedExtraParams.Add(extraParamName);

                            // Store the raw value as the parameter value.
                            extraParams[extraParamName] = rawValue;

                            // Store the display value under a derived _DISPLAY_VALUE key.
                            if (!string.IsNullOrEmpty(displayValue))
                            {
                                if (extraParamName == "ORACLE_RESP_ID")
                                {
                                    extraParams["ORACLE_RESP_DISPLAY_VALUE"] = displayValue;
                                }
                                else if (extraParamName == "ORACLE_GL_SEGMENT_VALUES")
                                {
                                    extraParams["ORACLE_GL_SEGMENT_DISPLAY_VALUES"] = displayValue;
                                }
                            }
                        }
                    }
                }
            }

            string extraParamsSummary = processedExtraParams.Count > 0 ? string.Join(", ", processedExtraParams) : "none";
            LogUtility.LogDebug($"{MethodName}|Processed {processedExtraParams.Count} extraParameter(s): {extraParamsSummary}");

            var output = new ReportParameterRequest
            {
                ReportId = reportId,
                Parameters = parameters,
                ExtraParameters = extraParams.Count > 0 ? extraParams : new object()
            };

            try
            {
                var returnJson = JsonSerializer.Serialize(output, JsonGlobals.Options);
                LogUtility.LogDebug($"{MethodName}|Serialized JSON: {returnJson}");

                string mergedJson = UpdateParameterData(returnJson, paramsJson);

                return (returnJson, mergedJson);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Failed to serialize parameter request");
                return (string.Empty, string.Empty);
            }
        }

        /// <summary>
        /// Merges the control-sheet-derived request payload (<paramref name="outputJson"/>) back into
        /// the report's original parameter metadata (<paramref name="parameterJson"/>), producing the
        /// merged array-shape JSON. Parameters edited on the control sheet get their operator/value
        /// updated (using the already-mapped operator); every other original parameter is carried
        /// forward unchanged rather than dropped.
        /// </summary>
        private static string UpdateParameterData(string outputJson, string parameterJson)
        {
            const string MethodName = "XLEdgeParamsBuilder.UpdateParameterData";

            try
            {
                using var outputDoc = JsonDocument.Parse(outputJson);
                using var paramDoc = JsonDocument.Parse(parameterJson);

                var outputParameters = outputDoc.RootElement.GetProperty("parameters");
                var parameterData = paramDoc.RootElement.Clone();

                var paramArray = new List<JsonElement>();
                foreach (JsonElement item in parameterData.EnumerateArray())
                {
                    paramArray.Add(item);
                }

                var updatedParams = new List<object>();
                var handledKeys = new HashSet<string>(StringComparer.Ordinal);

                foreach (JsonElement outputParam in outputParameters.EnumerateArray())
                {
                    ProcessOutputParameter(outputParam, paramArray, MethodName, updatedParams, handledKeys);
                }

                AppendUnhandledOriginalParameters(paramArray, handledKeys, updatedParams);

                AppendExtraParametersEntry(outputDoc.RootElement, updatedParams);

                var finalJson = JsonSerializer.Serialize(updatedParams, JsonGlobals.Options);
                XLEdgeAppState.Instance.UpdatedParamData = finalJson;

                LogUtility.LogDebug($"{MethodName}|Updated parameter data stored with extraParameters ({updatedParams.Count} total entries).");

                return finalJson;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{MethodName}|Failed to update parameter data");
                XLEdgeAppState.Instance.UpdatedParamData = parameterJson;
                return parameterJson;
            }
        }

        /// <summary>
        /// Handles a single output parameter entry: finds its matching original entry in
        /// <paramref name="paramArray"/> (by name or label) and either merges the updated
        /// operator/value(s) into a copy of that original entry, or falls back to carrying the
        /// original entry forward unchanged if no operator/value merge target-shape applies.
        /// Adds the result to <paramref name="updatedParams"/> and records the matched key in
        /// <paramref name="handledKeys"/> so the caller's "carry forward untouched originals"
        /// pass doesn't duplicate it.
        /// </summary>
        private static void ProcessOutputParameter(JsonElement outputParam, List<JsonElement> paramArray, string methodName, List<object> updatedParams, HashSet<string> handledKeys)
        {
            string paramName = outputParam.GetProperty("name").GetString();
            string paramOperator = outputParam.TryGetProperty("operator", out JsonElement opEl) ? opEl.GetString() : null;

            JsonElement? targetParam = null;
            foreach (JsonElement item in paramArray)
            {
                string itemName = item.TryGetProperty("name", out JsonElement n) ? n.GetString() : null;
                string itemLabel = item.TryGetProperty("label", out JsonElement l) ? l.GetString() : null;

                if (string.Equals(itemName, paramName, StringComparison.Ordinal) ||
                    string.Equals(itemLabel, paramName, StringComparison.Ordinal))
                {
                    targetParam = item;
                    break;
                }
            }

            if (targetParam.HasValue)
            {
                var updatedParam = new Dictionary<string, object>();
                var target = targetParam.Value;

                foreach (JsonProperty prop in target.EnumerateObject())
                {
                    updatedParam[prop.Name] = prop.Value.ValueKind switch
                    {
                        JsonValueKind.String => prop.Value.GetString(),
                        JsonValueKind.Number => prop.Value.ToString(),
                        JsonValueKind.True => true,
                        JsonValueKind.False => false,
                        JsonValueKind.Null => null,
                        _ => prop.Value.ToString()
                    };
                }

                // Excel-to-application operator mapping was already applied when this
                // value was computed (BuildJsonPayload's `updatedOperator`, via
                // XLEdgeOperatorMappings) - written through here unchanged so the merged/
                // stored JSON reflects the same mapped operator that was actually sent
                // to the API, not the raw Excel-side operator label.
                if (paramOperator != null)
                {
                    updatedParam["operator"] = paramOperator;
                }

                if (outputParam.TryGetProperty("value", out JsonElement valueEl) && valueEl.ValueKind != JsonValueKind.Null)
                {
                    updatedParam["value"] = valueEl.ValueKind == JsonValueKind.Number
                        ? valueEl.ToString()
                        : valueEl.GetString();
                }
                else if (outputParam.TryGetProperty("values", out JsonElement valuesEl) && valuesEl.ValueKind == JsonValueKind.Array)
                {
                    var values = new List<object>();
                    foreach (JsonElement v in valuesEl.EnumerateArray())
                    {
                        values.Add(v.ValueKind == JsonValueKind.Number ? v.ToString() : v.GetString());
                    }
                    updatedParam["values"] = values;
                }

                updatedParams.Add(updatedParam);

                string handledKey = targetParam.Value.TryGetProperty("name", out JsonElement hn) ? hn.GetString() : null;
                handledKey ??= targetParam.Value.TryGetProperty("label", out JsonElement hl) ? hl.GetString() : null;
                if (handledKey != null)
                {
                    handledKeys.Add(handledKey);
                }
            }
            else
            {
                // Fallback: look up the original parameter by name; skip if no match is found.
                JsonElement? orig = paramArray.Cast<JsonElement?>().FirstOrDefault(p =>
                {
                    string n = p.Value.TryGetProperty("name", out JsonElement ne) ? ne.GetString() : null;
                    return string.Equals(n, paramName, StringComparison.Ordinal);
                });

                if (orig.HasValue)
                {
                    updatedParams.Add(orig.Value);
                    handledKeys.Add(paramName);
                }
                else
                {
                    LogUtility.LogDebug($"{methodName}: control-sheet parameter '{paramName}' has no matching entry in the original parameter JSON - skipped.");
                }
            }
        }

        /// <summary>
        /// Carries forward every original parameter the control sheet didn't touch this round,
        /// unchanged. Bare extraParameters marker entries are skipped since a fresh one is
        /// appended separately (see <see cref="AppendExtraParametersEntry"/>).
        /// </summary>
        private static void AppendUnhandledOriginalParameters(List<JsonElement> paramArray, HashSet<string> handledKeys, List<object> updatedParams)
        {
            foreach (JsonElement orig in paramArray)
            {
                string itemName = orig.TryGetProperty("name", out JsonElement on) ? on.GetString() : null;
                string itemLabel = orig.TryGetProperty("label", out JsonElement ol) ? ol.GetString() : null;
                string key = itemName ?? itemLabel;

                if (key != null && handledKeys.Contains(key))
                {
                    continue;
                }

                bool isBareExtraParamsMarker = key == null &&
                    orig.TryGetProperty("extraParameters", out JsonElement _) ;

                if (isBareExtraParamsMarker)
                {
                    continue;
                }

                updatedParams.Add(orig);
            }
        }

        /// <summary>
        /// Appends a fresh "extraParameters" marker entry (built from <paramref name="outputRoot"/>'s
        /// own extraParameters object) to <paramref name="updatedParams"/>, if present.
        /// </summary>
        private static void AppendExtraParametersEntry(JsonElement outputRoot, List<object> updatedParams)
        {
            if (outputRoot.TryGetProperty("extraParameters", out JsonElement extraParamsEl) &&
                extraParamsEl.ValueKind == JsonValueKind.Object)
            {
                var extraDict = new Dictionary<string, object>();
                foreach (JsonProperty prop in extraParamsEl.EnumerateObject())
                {
                    extraDict[prop.Name] = prop.Value.ValueKind == JsonValueKind.String
                        ? prop.Value.GetString()
                        : prop.Value.ToString();
                }

                updatedParams.Add(new Dictionary<string, object>
                {
                    ["extraParameters"] = extraDict
                });
            }
        }

        private static List<string> SplitRespectingQuotes(string input)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool insideQuotes = false;

            foreach (char ch in input)
            {
                if (ch == '"')
                {
                    insideQuotes = !insideQuotes;
                }
                else if (ch == ',' && !insideQuotes)
                {
                    result.Add(current.ToString().Trim());
                    current.Clear();
                }
                else
                {
                    current.Append(ch);
                }
            }

            if (current.Length > 0)
            {
                result.Add(current.ToString().Trim());
            }

            return result;
        }

        private static object FormatValue(object value, string columnType)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {
                return null;
            }

            try
            {
                string strVal = value.ToString().Trim();
                columnType = (columnType ?? "STRING").ToUpper();

                if (columnType == "DATE" || columnType == "DATETIME")
                {
                    return FormatDateValue(strVal);
                }
                else if (columnType == "INTEGER" && int.TryParse(strVal, out int intVal))
                {
                    return intVal;
                }
                else if ((columnType == "DECIMAL" || columnType == "NUMERIC") && strVal.Contains("."))
                {
                    if (decimal.TryParse(strVal, out decimal decVal))
                    {
                        return decVal;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"FormatValue: failed to format '{value}' - {ex.Message}");
            }

            return value;
        }

        private static string FormatDateValue(object value)
        {
            try
            {
                if (value is DateTime dt)
                {
                    return dt.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                }
                else if (value is double serialDate)
                {
                    var convertedDate = DateTime.FromOADate(serialDate);
                    return convertedDate.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                }
                else if (value is string strVal)
                {
                    string[] formats = {
                        "M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MMM-yyyy", "yyyy/MM/dd",
                        "d/M/yyyy", "yyyy.MM.dd", "MM-dd-yyyy", "dd-MM-yyyy", "yyyy MMM dd",
                        "yyyyMMdd", "MMM dd, yyyy", "dd MMM yyyy", "dd/MM/yyyy",
                        "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-ddTHH:mm:ss.fff"
                    };

                    if (DateTime.TryParseExact(strVal, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                    {
                        return parsedDate.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                    }
                    else if (DateTime.TryParse(strVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                    {
                        return parsedDate.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"FormatDateValue: failed - {ex.Message}");
            }

            return value?.ToString() ?? string.Empty;
        }
    }

    /// <summary>
    /// Represents a parameter value in the report request
    /// </summary>
    public class ReportParameterValue
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("value")]
        [JsonConverter(typeof(NumericJsonConverter))]
        public object Value { get; set; }

        [JsonPropertyName("values")]
        public List<object> Values { get; set; }

        [JsonPropertyName("operator")]
        public string Operator { get; set; }
    }

    /// <summary>
    /// Represents the complete report parameter request
    /// </summary>
    public sealed class ReportParameterRequest
    {
        [JsonPropertyName("reportId")]
        [JsonConverter(typeof(NumericJsonConverter))]
        public object ReportId { get; set; }

        [JsonPropertyName("parameters")]
        public List<ReportParameterValue> Parameters { get; set; } = new List<ReportParameterValue>();

        [JsonPropertyName("extraParameters")]
        public object ExtraParameters { get; set; }
    }
}