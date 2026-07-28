using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using XLEdge.Models;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Builds the JSON request body for a drilldown click, scoping the child report to the clicked
    /// row's parent parameter values instead of re-running it unfiltered. Also attaches extra
    /// parameter display values (ORACLE_RESP_DISPLAY_VALUE, ORACLE_GL_SEGMENT_DISPLAY_VALUES) from
    /// the parameter sheet's hidden cells (IT4=ORACLE_RESP_ID, IU4=ORACLE_RESP_DISPLAY_VALUE,
    /// IV4=ORACLE_GL_SEGMENT_VALUES, IW4=ORACLE_GL_SEGMENT_DISPLAY_VALUES) so they're included in
    /// the drilldown request payload.
    /// </summary>
    public static class DrilldownRequestBuilder
    {
        /// <summary>
        /// Looks up a report column's declared "datatype" by name, defaulting to "STRING" if not found.
        /// </summary>
        public static string GetColumnType(ReportMeta reportMeta, string columnName)
        {
            try
            {
                RptColumn match = reportMeta?.Columns?.FirstOrDefault(c => string.Equals(c.Name, columnName, StringComparison.Ordinal));
                return match?.DataType ?? "STRING";
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(GetColumnType));
                return "STRING";
            }
        }

        /// <summary>
        /// Finds this report's own stored parameter value (from the "Param" CustomXMLPart JSON
        /// captured when the report was first run) for a given parameter name, so a "PARAM"-type
        /// drilldown parameter definition can be resolved to whatever value the parent report was
        /// actually run with - not the clicked row's cell value.
        /// </summary>
        public static ReportParameterValue ResolveStoredParamValue(string storedParamsJson, string colLabel)
        {
            var result = new ReportParameterValue { Name = colLabel, Value = null, Values = null };

            if (string.IsNullOrWhiteSpace(storedParamsJson) || string.IsNullOrWhiteSpace(colLabel))
            {
                return result;
            }

            try
            {
                using var doc = JsonDocument.Parse(storedParamsJson);
                if (doc.RootElement.ValueKind != JsonValueKind.Array)
                {
                    return result;
                }

                string trimmedLabel = colLabel.Trim();

                foreach (JsonElement item in doc.RootElement.EnumerateArray())
                {
                    if (!JsonHelper.TryGetProperty(item, "name", out JsonElement nameEl))
                    {
                        continue;
                    }

                    string itemName = nameEl.ValueKind == JsonValueKind.String ? nameEl.GetString() : nameEl.ToString();
                    if (!string.Equals(itemName?.Trim(), trimmedLabel, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    // Priority order matches VB exactly: values, values1, value, value1.
                    if (JsonHelper.TryGetProperty(item, "values", out JsonElement valuesEl) && valuesEl.ValueKind == JsonValueKind.Array)
                    {
                        result.Values = FormatArray(valuesEl);
                    }
                    else if (JsonHelper.TryGetProperty(item, "values1", out JsonElement values1El) && values1El.ValueKind == JsonValueKind.Array)
                    {
                        result.Values = FormatArray(values1El);
                    }
                    else if (JsonHelper.TryGetProperty(item, "value", out JsonElement valueEl) && valueEl.ValueKind != JsonValueKind.Null && valueEl.ValueKind != JsonValueKind.Undefined)
                    {
                        AssignScalarOrArray(result, valueEl);
                    }
                    else if (JsonHelper.TryGetProperty(item, "value1", out JsonElement value1El) && value1El.ValueKind != JsonValueKind.Null && value1El.ValueKind != JsonValueKind.Undefined)
                    {
                        AssignScalarOrArray(result, value1El);
                    }

                    return result;
                }

                return result;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ResolveStoredParamValue));
                return result;
            }
        }

        private static void AssignScalarOrArray(ReportParameterValue result, JsonElement valueEl)
        {
            if (valueEl.ValueKind == JsonValueKind.Array)
            {
                result.Values = FormatArray(valueEl);
            }
            else
            {
                object raw = RawValue(valueEl);
                result.Value = XLEdgeValueFormatter.FormatDrilldownValue(raw, XLEdgeValueFormatter.InferDrilldownDataType(raw));
            }
        }

        private static List<object> FormatArray(JsonElement arrayEl)
        {
            var list = new List<object>();
            foreach (JsonElement v in arrayEl.EnumerateArray())
            {
                object raw = RawValue(v);
                list.Add(XLEdgeValueFormatter.FormatDrilldownValue(raw, XLEdgeValueFormatter.InferDrilldownDataType(raw)));
            }
            return list;
        }

        private static object RawValue(JsonElement el)
        {
            switch (el.ValueKind)
            {
                case JsonValueKind.String: return el.GetString();
                case JsonValueKind.Number: return el.ToString();
                case JsonValueKind.True: return true;
                case JsonValueKind.False: return false;
                case JsonValueKind.Null:
                case JsonValueKind.Undefined: return null;
                default: return el.ToString();
            }
        }

        /// <summary>
        /// Reads stored extra parameter display values from the parameter sheet's hidden cells.
        /// These are written by ReportGenerator when the report is first created or refreshed.
        ///
        /// Cell mappings:
        /// - IT4: ORACLE_RESP_ID (raw responsibility ID)
        /// - IU4: ORACLE_RESP_DISPLAY_VALUE (human-readable responsibility name)
        /// - IV4: ORACLE_GL_SEGMENT_VALUES (raw GL segment values)
        /// - IW4: ORACLE_GL_SEGMENT_DISPLAY_VALUES (formatted GL segment display values)
        /// </summary>
        private static Dictionary<string, string> ReadStoredExtraDisplayValues(Excel.Worksheet paramSheet)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (paramSheet == null)
            {
                return result;
            }

            try
            {
                // Read Responsibility display value (IU4)
                try
                {
                    Excel.Range respDisplayCell = paramSheet.Range["IU4"];
                    if (respDisplayCell != null && respDisplayCell.Value != null)
                    {
                        string respDisplayValue = Convert.ToString(respDisplayCell.Value);
                        if (!string.IsNullOrWhiteSpace(respDisplayValue))
                        {
                            result["ORACLE_RESP_DISPLAY_VALUE"] = respDisplayValue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ReadStoredExtraDisplayValues: Failed to read IU4 - {ex.Message}");
                }

                // Read Responsibility ID (IT4)
                try
                {
                    Excel.Range respIdCell = paramSheet.Range["IT4"];
                    if (respIdCell != null && respIdCell.Value != null)
                    {
                        string respId = Convert.ToString(respIdCell.Value);
                        if (!string.IsNullOrWhiteSpace(respId))
                        {
                            result["ORACLE_RESP_ID"] = respId;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ReadStoredExtraDisplayValues: Failed to read IT4 - {ex.Message}");
                }

                // Read GL Segment display values (IW4)
                try
                {
                    Excel.Range glSegmentCell = paramSheet.Range["IW4"];
                    if (glSegmentCell != null && glSegmentCell.Value != null)
                    {
                        string glSegmentValue = Convert.ToString(glSegmentCell.Value);
                        if (!string.IsNullOrWhiteSpace(glSegmentValue))
                        {
                            result["ORACLE_GL_SEGMENT_DISPLAY_VALUES"] = glSegmentValue;
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ReadStoredExtraDisplayValues: Failed to read IW4 - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ReadStoredExtraDisplayValues: Unexpected error");
            }

            return result;
        }

        /// <summary>
        /// Builds the full drilldown request body: matches every drilldown definition for the clicked
        /// column against the child report id, resolves each of its parameters (PARAM/STATIC/cell-value
        /// types), attaches the "Responsibility" extra parameter (IT4/ORACLE_RESP_ID) and its display
        /// value (IU4/ORACLE_RESP_DISPLAY_VALUE) from the resolved parameter sheet, along with GL
        /// segment display values (IV4/ORACLE_GL_SEGMENT_DISPLAY_VALUES), and serializes the result.
        /// Returns null if nothing could be built (caller should fall back to an unfiltered drilldown
        /// rather than block the user).
        /// </summary>
        public static string BuildDrilldownRequestJson(
            ReportMeta reportMeta,
            string storedParamsJson,
            string childReportId,
            string columnName,
            Excel.Range headerRange,
            Excel.Worksheet dataSheet,
            Excel.Range clickedCell,
            Excel.Worksheet parameterSheetForExtras)
        {
            try
            {
                var request = new ReportParameterRequest
                {
                    ReportId = childReportId,
                    Parameters = new List<ReportParameterValue>()
                };

                IEnumerable<RptDrilldown> matches = (reportMeta?.Drilldowns ?? Array.Empty<RptDrilldown>())
                    .Where(d => string.Equals(d.DrillReportId.ToString(), childReportId, StringComparison.Ordinal) &&
                                string.Equals(d.DrillColumnName?.Trim(), columnName?.Trim(), StringComparison.OrdinalIgnoreCase));

                foreach (RptDrilldown drilldown in matches)
                {
                    if (drilldown.Parameters == null)
                    {
                        continue;
                    }

                    foreach (ChildParameter param in drilldown.Parameters)
                    {
                        string paramType = (param.Type ?? string.Empty).ToUpperInvariant();

                        if (paramType == "PARAM")
                        {
                            request.Parameters.Add(ResolveStoredParamValue(storedParamsJson, param.Name));
                        }
                        else if (paramType == "STATIC")
                        {
                            object staticVal = param.StaticValue;
                            if (staticVal == null)
                            {
                                continue;
                            }

                            string inferredType = XLEdgeValueFormatter.InferDrilldownDataType(staticVal);
                            object formatted = XLEdgeValueFormatter.FormatDrilldownValue(staticVal, inferredType);
                            var staticParam = new ReportParameterValue { Name = param.ParamName, Value = formatted };

                            if (staticParam.Values != null || staticParam.Value != null)
                            {
                                request.Parameters.Add(staticParam);
                            }
                        }
                        else
                        {
                            object cellValue = null;
                            try
                            {
                                int matchCol = ExcelSheetHelper.HRMatch(headerRange, param.Name);
                                if (matchCol > 0 && clickedCell != null)
                                {
                                    object raw = ((Excel.Range)dataSheet.Cells[clickedCell.Row, matchCol]).Value;
                                    if (raw != null)
                                    {
                                        string inferredType = XLEdgeValueFormatter.InferDrilldownDataType(raw);
                                        cellValue = XLEdgeValueFormatter.FormatDrilldownValue(raw, inferredType);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogException(ex, "BuildDrilldownRequestJson: failed to resolve cell-value drilldown parameter");
                            }

                            request.Parameters.Add(new ReportParameterValue { Name = param.ParamName, Value = cellValue });
                        }
                    }
                }

                // Read stored extra parameters (Responsibility and GL Segment, raw + display values)
                // from the parameter sheet's hidden cells and attach them to the request.

                var extraParams = new Dictionary<string, object>();

                try
                {
                    object it4Value = parameterSheetForExtras?.Range["IT4"]?.Value;

                    if (it4Value != null)
                    {
                        string respId = Convert.ToString(it4Value);
                        extraParams["ORACLE_RESP_ID"] = respId;
                    }

                    object iu4Value = parameterSheetForExtras?.Range["IU4"]?.Value;
                    if (iu4Value != null && !string.IsNullOrWhiteSpace(iu4Value.ToString()))
                    {
                        string respDisplayValue = iu4Value.ToString();
                        extraParams["ORACLE_RESP_DISPLAY_VALUE"] = respDisplayValue;
                    }

                    object iv4Value = parameterSheetForExtras?.Range["IV4"]?.Value;
                    if (iv4Value != null && !string.IsNullOrWhiteSpace(iv4Value.ToString()))
                    {
                        string glSegmentValues = iv4Value.ToString();
                        extraParams["ORACLE_GL_SEGMENT_VALUES"] = glSegmentValues;
                    }

                    object iw4Value = parameterSheetForExtras?.Range["IW4"]?.Value;
                    if (iw4Value != null && !string.IsNullOrWhiteSpace(iw4Value.ToString()))
                    {
                        string glSegmentDisplayValues = iw4Value.ToString();
                        extraParams["ORACLE_GL_SEGMENT_DISPLAY_VALUES"] = glSegmentDisplayValues;
                    }

                    // Single summary log of what was actually attached, instead of one line per field -
                    // useful for correlating a future "missing extra parameter in drilldown payload" report.
                    LogUtility.LogDebug($"BuildDrilldownRequestJson: attached {extraParams.Count} extra parameter(s) - {string.Join(", ", extraParams.Keys)}");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "BuildDrilldownRequestJson: exception reading parameters (IT4/IU4/IV4/IW4)");
                }

                // extraParameters is always a Dictionary<string, object>; an empty dictionary
                // serializes as "{}".
                request.ExtraParameters = extraParams;

                return ReportParameterRequestSerializer.Serialize(request);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(BuildDrilldownRequestJson));
                return null;
            }
        }
        /// <summary>
        /// Builds extra parameters for drilldown request with preserved display values.
        /// </summary>
        private static Dictionary<string, object> BuildExtraParametersForDrilldown(Excel.Worksheet paramSheet)
        {
            var extraParams = new Dictionary<string, object>();

            try
            {
                // Read stored display values
                var storedDisplayValues = ReadStoredExtraDisplayValues(paramSheet);

                // Read IT4 (Responsibility ID)
                object it4Value = paramSheet?.Range["IT4"]?.Value;
                if (it4Value != null)
                {
                    string respId = Convert.ToString(it4Value);
                    extraParams["ORACLE_RESP_ID"] = respId;

                    // CRITICAL: Preserve display value from IU4
                    if (storedDisplayValues.TryGetValue("ORACLE_RESP_DISPLAY_VALUE", out string respDisplay))
                    {
                        extraParams["ORACLE_RESP_DISPLAY_VALUE"] = respDisplay;
                    }
                }

                // Read IV4 (GL Segment values)
                object iv4Value = paramSheet?.Range["IV4"]?.Value;
                if (iv4Value != null && !string.IsNullOrWhiteSpace(iv4Value.ToString()))
                {
                    string glSegmentValues = iv4Value.ToString();
                    extraParams["ORACLE_GL_SEGMENT_VALUES"] = glSegmentValues;

                    // CRITICAL: Preserve GL segment display values from stored params
                    if (storedDisplayValues.TryGetValue("ORACLE_GL_SEGMENT_DISPLAY_VALUES", out string glDisplay))
                    {
                        extraParams["ORACLE_GL_SEGMENT_DISPLAY_VALUES"] = glDisplay;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "BuildExtraParametersForDrilldown: Failed to read extra parameters");
            }

            return extraParams.Count > 0 ? extraParams : null;
        }
    }
}