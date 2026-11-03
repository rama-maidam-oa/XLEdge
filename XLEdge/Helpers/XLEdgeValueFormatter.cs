using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
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

        /// <summary>
        /// Builds the value written into a report cell. Same as <see cref="FormatValue"/>, except a
        /// numeric-typed column's value is parsed into a genuine double (server values always use "."
        /// decimal / "," thousands) instead of being passed through as a string - a string like
        /// "1234.56" is re-parsed by Excel using the user's locale, so on a European-locale Excel it
        /// became text or the wrong number. Ported from FormProcessBar.vb's GetExcelCellValue and the
        /// TextToColumns(DecimalSeparator:=".", ThousandsSeparator:=",") fix.
        /// DATE/TIME-typed values are likewise written as an OLE Automation date serial (as VB does)
        /// rather than a "dd-MMM-yyyy" string, which a non-English Excel can't parse ("Oct" vs "Okt").
        /// </summary>
        public static object ToExcelCellValue(object value, string columnType)
        {
            return ToExcelCellValue(value, columnType, null);
        }

        /// <summary>
        /// Same as <see cref="ToExcelCellValue(object, string)"/>; <paramref name="columnFormat"/> is the column's Excel number
        /// format (the report's fmt), used to write a date before 1900 as text (OISR-22697, see <see cref="ToExcelDateValue"/>).
        /// </summary>
        public static object ToExcelCellValue(object value, string columnType, string columnFormat)
        {
            object formatted = FormatValue(value, columnType);
            if (!(formatted is string text) || text.Length == 0 ||
                string.Equals(columnType, "STRING", StringComparison.OrdinalIgnoreCase))
            {
                return formatted;
            }

            if (columnType.IndexOf("DATE", StringComparison.OrdinalIgnoreCase) >= 0 ||
                columnType.IndexOf("TIME", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                // Parse the raw value, not FormatValue's "dd-MMM-yyyy" output, so a DATETIME keeps its
                // time component. Known formats first (same order FormatDateValue uses), then VB's
                // GetExcelCellValue fallbacks (invariant, then current culture).
                string raw = value.ToString().Trim();
                if (DateTime.TryParseExact(raw, KnownDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date) ||
                    DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.NoCurrentDateDefault, out date) ||
                    DateTime.TryParse(raw, CultureInfo.CurrentCulture, DateTimeStyles.NoCurrentDateDefault, out date))
                {
                    return ToExcelDateValue(date, columnFormat);
                }

                return formatted;
            }

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double number) ||
                double.TryParse(text, NumberStyles.Number, CultureInfo.CurrentCulture, out number))
            {
                return number;
            }

            return formatted;
        }

        private static readonly DateTime ExcelFirstDate = new DateTime(1900, 1, 1);
        private static readonly DateTime ExcelLeapYearBugEnd = new DateTime(1900, 3, 1);
        private const string Pre1900FallbackFormat = "dd-MMM-yyyy";

        /// <summary>
        /// The value written into a cell for a parsed date (OISR-22697). Excel's dates start on 1-Jan-1900 and Excel
        /// counts a 29-Feb-1900 that never was, so an OLE Automation serial is wrong there:
        /// - a time-only value (day zero) stays a pure time fraction (OISR-22663);
        /// - before 1-Jan-1900 the serial is negative and Excel shows "####": the date is written as text in the
        ///   column's format instead (as the server's Export to Excel does with "Year Prior to 1900"), with a leading
        ///   apostrophe so Excel keeps it as text;
        /// - 1-Jan..28-Feb-1900 the OLE serial is one more than Excel's, so it is written one lower;
        /// - from 1-Mar-1900 the OLE serial is Excel's.
        /// </summary>
        public static object ToExcelDateValue(DateTime date, string columnFormat)
        {
            if (date.Ticks < TimeSpan.TicksPerDay)
            {
                return date.ToOADate();
            }

            if (date < ExcelFirstDate)
            {
                return "'" + FormatPre1900Date(date, columnFormat);
            }

            double serial = date.ToOADate();
            return date < ExcelLeapYearBugEnd ? serial - 1 : serial;
        }

        /// <summary>A date Excel cannot hold, as the text the column's Excel format would show (dd-MMM-yyyy when it can't be read).</summary>
        public static string FormatPre1900Date(DateTime date, string columnFormat)
        {
            string format = ExcelDateFormatToDotNet(columnFormat) ?? Pre1900FallbackFormat;
            try
            {
                return date.ToString(format, CultureInfo.InvariantCulture);
            }
            catch (FormatException ex)
            {
                LogUtility.LogDebug($"{nameof(FormatPre1900Date)}: format '{columnFormat}' -> '{format}' failed, using {Pre1900FallbackFormat} - {ex.Message}");
                return date.ToString(Pre1900FallbackFormat, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>
        /// Converts an Excel date/time number format (e.g. "dd-MMM-YYYY", "yyyy/mm/dd hh:mm AM/PM") into a .NET custom
        /// format. Excel's codes are case-insensitive; "m"/"mm" are minutes after an hour or before a second. Returns
        /// null for anything else (General, "@", numbers, colours, conditions, ...), so the caller falls back.
        /// </summary>
        public static string ExcelDateFormatToDotNet(string excelFormat)
        {
            if (string.IsNullOrWhiteSpace(excelFormat))
            {
                return null;
            }

            string section = excelFormat.Split(';')[0];
            bool twelveHour = section.IndexOf("AM/PM", StringComparison.OrdinalIgnoreCase) >= 0 ||
                              section.IndexOf("A/P", StringComparison.OrdinalIgnoreCase) >= 0;

            // Each token: a code letter ('d','m','y','h','s','t' for AM/PM, 'a' for A/P) with its run length, or 'L' with literal text.
            var tokens = new List<KeyValuePair<char, string>>();
            int i = 0;
            while (i < section.Length)
            {
                char c = section[i];
                if (c == '"')
                {
                    int end = section.IndexOf('"', i + 1);
                    if (end < 0)
                    {
                        return null;
                    }

                    tokens.Add(new KeyValuePair<char, string>('L', section.Substring(i + 1, end - i - 1)));
                    i = end + 1;
                    continue;
                }

                if (c == '\\')
                {
                    if (i + 1 >= section.Length)
                    {
                        return null;
                    }

                    tokens.Add(new KeyValuePair<char, string>('L', section[i + 1].ToString()));
                    i += 2;
                    continue;
                }

                if (string.Compare(section, i, "AM/PM", 0, 5, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    tokens.Add(new KeyValuePair<char, string>('t', "tt"));
                    i += 5;
                    continue;
                }

                if (string.Compare(section, i, "A/P", 0, 3, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    tokens.Add(new KeyValuePair<char, string>('a', "t"));
                    i += 3;
                    continue;
                }

                char code = char.ToLowerInvariant(c);
                if ("dmyhs".IndexOf(code) >= 0)
                {
                    int j = i;
                    while (j < section.Length && char.ToLowerInvariant(section[j]) == code)
                    {
                        j++;
                    }

                    tokens.Add(new KeyValuePair<char, string>(code, new string(code, j - i)));
                    i = j;
                    continue;
                }

                if (char.IsLetterOrDigit(c) || "[@*_?#".IndexOf(c) >= 0)
                {
                    return null;
                }

                tokens.Add(new KeyValuePair<char, string>('L', c.ToString()));
                i++;
            }

            var result = new StringBuilder();
            bool hasDatePart = false;
            for (int t = 0; t < tokens.Count; t++)
            {
                char code = tokens[t].Key;
                int length = tokens[t].Value.Length;
                switch (code)
                {
                    case 'L':
                        foreach (char literal in tokens[t].Value)
                        {
                            result.Append('\\').Append(literal);
                        }
                        break;
                    case 'd':
                        result.Append(new string('d', Math.Min(length, 4)));
                        hasDatePart = true;
                        break;
                    case 'y':
                        result.Append(length <= 2 ? "yy" : "yyyy");
                        hasDatePart = true;
                        break;
                    case 'h':
                        result.Append(new string(twelveHour ? 'h' : 'H', Math.Min(length, 2)));
                        break;
                    case 's':
                        result.Append(new string('s', Math.Min(length, 2)));
                        break;
                    case 't':
                    case 'a':
                        result.Append(tokens[t].Value);
                        break;
                    case 'm':
                        bool minutes = length <= 2 && (NeighbourCode(tokens, t, -1) == 'h' || NeighbourCode(tokens, t, +1) == 's');
                        if (minutes)
                        {
                            result.Append(new string('m', length));
                        }
                        else if (length <= 4)
                        {
                            result.Append(new string('M', length));
                            hasDatePart = true;
                        }
                        else
                        {
                            return null;
                        }
                        break;
                }
            }

            return hasDatePart ? result.ToString() : null;
        }

        // The nearest date/time code before (direction -1) or after (+1) token t, skipping literals; '\0' if none.
        private static char NeighbourCode(List<KeyValuePair<char, string>> tokens, int t, int direction)
        {
            for (int k = t + direction; k >= 0 && k < tokens.Count; k += direction)
            {
                if (tokens[k].Key != 'L')
                {
                    return tokens[k].Key;
                }
            }

            return '\0';
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
