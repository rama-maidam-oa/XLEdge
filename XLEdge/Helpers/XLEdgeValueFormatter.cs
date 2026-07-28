using System;
using System.Globalization;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Value/date formatting utilities ported from XLEdgeProcedures.vb
    /// (XLEdgeFormatDateValue, XLEdgeFormatValue, RemoveEquaSymbol).
    /// Pure string logic - no Excel interop, no COM objects involved.
    /// </summary>
    public static class XLEdgeValueFormatter
    {
        private static readonly string[] KnownDateFormats =
        {
            "M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MMM-yyyy", "yyyy/MM/dd", "d/M/yyyy", "yyyy.MM.dd",
            "MM-dd-yyyy", "dd-MM-yyyy", "yyyy MMM dd", "yyyyMMdd", "MMM dd, yyyy", "dd MMM yyyy", "dd/MM/yyyy",
            "yyyyMMddHHmmss", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "dd-MM-yyyy HH:mm:ss", "M/d/yyyy h:mm:ss tt", "MM-dd-yyyy h:mm:ss tt",
            "d-M-yyyy", "M-d-yyyy", "yyyy.MM.dd HH:mm:ss", "yyyy MMMM dd", "dd MMMM yyyy",
            "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss.fffZ",
            "M/d/yyyy h:mm:ss tt", "MM/dd/yyyy h:mm:ss tt", "dd/MM/yyyy HH:mm:ss", "dd-MM-yyyy HH:mm:ss", "yyyy-MM-ddTHH:mm:ss.fffK"
        };

        /// <summary>
        /// Attempts to parse a value against a fixed list of known incoming date formats and,
        /// if successful, re-renders it as "dd-MMM-yyyy". Returns the original value unchanged
        /// if no known format matches.
        /// </summary>
        public static string FormatDateValue(string value)
        {
            try
            {
                if (DateTime.TryParseExact(value, KnownDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                {
                    return parsedDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture);
                }

                return value;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(FormatDateValue));
                return value;
            }
        }

        /// <summary>
        /// Formats a raw report cell value according to its declared column type.
        /// Only DATE/DATETIME columns are reformatted today; everything else passes through.
        /// </summary>
        public static object FormatValue(object value, string columnType)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString()) ||
                string.IsNullOrEmpty(columnType))
            {
                return null;
            }

            string trimmed;
            try
            {
                trimmed = value.ToString().Trim();
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort trim; falls back to the untrimmed ToString().
                LogUtility.LogDebug($"{nameof(FormatValue)}: failed to trim value, using untrimmed - {ex.Message}");
                trimmed = value.ToString();
            }

            string upperType = columnType.ToUpperInvariant();
            if (upperType == "DATE" || upperType == "DATETIME")
            {
                return FormatDateValue(trimmed);
            }

            return trimmed;
        }

        // Ported from AddinModule.vb's InferDataType (line ~1933) - distinct from the report-column
        // "datatype" metadata: this guesses a type from the raw value itself, used only for the
        // drilldown-request parameter building in AdxExcelAppEvents1_SheetFollowHyperlink (STATIC
        // values and values pulled from the report's own stored parameter JSON via ReportHLink_Param).
        private static readonly string[] InferDateFormats =
        {
            "yy/MM/dd", "yyyy/MM/dd", "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd",
            "dd-MM-yyyy", "MM-dd-yyyy", "yyyyMMdd", "dd MMM yyyy", "yyyy-MM-ddTHH:mm:ss"
        };

        public static string InferDrilldownDataType(object value)
        {
            if (value == null)
            {
                return "STRING";
            }

            string strVal = value.ToString().Trim();

            if (DateTime.TryParse(strVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                return "DATE";
            }

            if (DateTime.TryParseExact(strVal, InferDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                return "DATE";
            }

            if (int.TryParse(strVal, out _))
            {
                return "INTEGER";
            }

            if (decimal.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out _))
            {
                return "DECIMAL";
            }

            return "STRING";
        }

        // Ported from AddinModule.vb's FormatValue1 (line ~1816). Distinct from FormatValue above:
        // that one only reformats DATE/DATETIME strings for on-sheet display; this one is used when
        // building a drilldown request body, where DATE/DATETIME must round-trip as an ISO string the
        // server expects, and INTEGER/DECIMAL/NUMERIC values must be written as genuine JSON numbers
        // (int/decimal/BigInteger), not strings, matching the VB original exactly.
        public static object FormatDrilldownValue(object value, string columnType)
        {
            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {
                return null;
            }

            try
            {
                value = value.ToString().Trim();
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort trim; use value as-is.
                LogUtility.LogDebug($"{nameof(FormatDrilldownValue)}: failed to trim value, using as-is - {ex.Message}");
            }

            switch ((columnType ?? string.Empty).ToUpperInvariant())
            {
                case "DATE":
                case "DATETIME":
                    return FormatDrilldownDateIso(value);

                case "INTEGER":
                    if (int.TryParse(Convert.ToString(value), out int intValue))
                    {
                        return intValue;
                    }
                    break;

                case "DECIMAL":
                case "NUMERIC":
                    string strVal = Convert.ToString(value);
                    if (strVal.Contains("."))
                    {
                        if (decimal.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal decimalValue))
                        {
                            return decimalValue;
                        }
                    }
                    else if (System.Numerics.BigInteger.TryParse(strVal, out System.Numerics.BigInteger bigIntValue))
                    {
                        return bigIntValue;
                    }
                    break;
            }

            // Default: pass through as a (trimmed) string - matches VB falling out of the Select Case.
            return value;
        }

        private static readonly string[] DrilldownDateIsoFormats =
        {
            "M/d/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MMM-yyyy", "yyyy/MM/dd", "d/M/yyyy", "yyyy.MM.dd",
            "MM-dd-yyyy", "dd-MM-yyyy", "yyyy MMM dd", "yyyyMMdd", "MMM dd, yyyy", "dd MMM yyyy", "dd/MM/yyyy",
            "yyyyMMddHHmmss", "yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss", "dd-MM-yyyy HH:mm:ss",
            "d-M-yyyy", "M-d-yyyy", "yyyy.MM.dd HH:mm:ss", "yyyy MMMM dd", "dd MMMM yyyy",
            "yyyy-MM-ddTHH:mm:ss", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-ddTHH:mm:ss.fff", "yyyy-MM-ddTHH:mm:ss.fffZ",
            "yy/MM/dd", "yy-MM-dd", "yyMMdd"
        };

        // Ported from AddinModule.vb's FormatDateValue1 - renders as ISO "yyyy-MM-ddTHH:mm:ss" for the
        // drilldown request body (vs. FormatDateValue above, which renders "dd-MMM-yyyy" for display).
        private static object FormatDrilldownDateIso(object value)
        {
            if (value is DateTime dtValue)
            {
                return dtValue.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
            }

            if (value is string strValue)
            {
                string dateString = strValue.Trim();

                try
                {
                    if (DateTime.TryParseExact(dateString, DrilldownDateIsoFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedExact))
                    {
                        return parsedExact.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(FormatDrilldownDateIso));
                }

                try
                {
                    if (DateTime.TryParse(dateString, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedGeneric))
                    {
                        return parsedGeneric.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(FormatDrilldownDateIso));
                }

                return value;
            }

            // Excel serial-date numeric fallback (matches VB's IsNumeric(value) branch).
            if (double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out double serialDate))
            {
                try
                {
                    return DateTime.FromOADate(serialDate).ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, nameof(FormatDrilldownDateIso));
                    return string.Empty;
                }
            }

            return value;
        }

        /// <summary>
        /// Approximates VB's runtime IsNumeric() function for the string-vs-numeric checks used when
        /// deciding whether a list value needs comma/quote escaping (ReportParamValue's inline checks).
        /// </summary>
        public static bool IsNumeric(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out _);
        }

        /// <summary>
        /// Abbreviates the local time zone's DST-aware name to its initials
        /// (e.g. "Eastern Standard Time" -> "EST").
        /// </summary>
        public static string PrintTimeZone(DateTime dt)
        {
            try
            {
                TimeZoneInfo local = TimeZoneInfo.Local;
                string name = local.IsDaylightSavingTime(dt) ? local.DaylightName : local.StandardName;

                if (string.IsNullOrEmpty(name))
                {
                    return string.Empty;
                }

                var initials = new System.Text.StringBuilder();
                foreach (string word in name.Split(' '))
                {
                    if (word.Length >= 1)
                    {
                        initials.Append(word.Substring(0, 1));
                    }
                }

                return initials.ToString();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(PrintTimeZone));
                return string.Empty;
            }
        }

        /// <summary>
        /// Prefixes values containing "=" with a single quote so Excel treats them as text
        /// instead of attempting to evaluate them as a formula.
        /// </summary>
        public static string RemoveEquaSymbol(string value)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(value) && value.Contains("="))
                {
                    return "'" + value;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(RemoveEquaSymbol));
            }

            return value;
        }
    }
}
