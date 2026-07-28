using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Windows;
using XLEdge.Helpers;
using XLEdge.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Views
{
    public partial class XLEdgeGLAccountsWindow : DpiAwareWindow
    {
        private Excel.Worksheet _controlSheet;
        private int _rowNumber;
        private List<SegmentData> _segments;

        public class SegmentData
        {
            public string SegmentName { get; set; }
            public string Low { get; set; } = string.Empty;
            public string High { get; set; } = string.Empty;
        }

        public XLEdgeGLAccountsWindow(Excel.Worksheet controlSheet, int rowNumber, string displayValuesJson)
        {
            InitializeComponent();
            _controlSheet = controlSheet;
            _rowNumber = rowNumber;

            LoadSegments(displayValuesJson);
        }

        private void LoadSegments(string displayValuesJson)
        {
            _segments = new List<SegmentData>();

            if (string.IsNullOrEmpty(displayValuesJson))
            {
                LogUtility.LogDebug($"{nameof(LoadSegments)}: displayValuesJson is null or empty for row {_rowNumber}");
                return;
            }

            try
            {
                // Parse the JSON from IA cell: {"Company":"00-T","Department":"-","Account":"1000-zzzz","Sub-Account":"-","Product":"-"}
                using var doc = JsonDocument.Parse(displayValuesJson);
                var root = doc.RootElement;

                if (root.ValueKind == JsonValueKind.Object)
                {
                    foreach (var prop in root.EnumerateObject())
                    {
                        var segment = new SegmentData
                        {
                            SegmentName = prop.Name
                        };

                        string value = prop.Value.ToString();
                        if (!string.IsNullOrEmpty(value) && value != "-")
                        {
                            var parts = value.Split('-');
                            if (parts.Length >= 2)
                            {
                                segment.Low = parts[0] != "-" ? parts[0] : string.Empty;
                                segment.High = parts[1] != "-" ? parts[1] : string.Empty;
                            }
                            else
                            {
                                segment.Low = value != "-" ? value : string.Empty;
                            }
                        }
                        _segments.Add(segment);
                    }

                    LogUtility.LogDebug($"{nameof(LoadSegments)}: Successfully loaded {_segments.Count} segments from JSON for row {_rowNumber}");
                }
                else
                {
                    LogUtility.LogDebug($"{nameof(LoadSegments)}: Root element is not an object, trying string format parsing for row {_rowNumber}");
                    ParseFromStringFormat(displayValuesJson);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(LoadSegments)}: Failed to parse JSON for row {_rowNumber}. JSON: {displayValuesJson}");
                ParseFromStringFormat(displayValuesJson);
            }

            SegmentItemsControl.ItemsSource = _segments;
        }

        private void ParseFromStringFormat(string displayValuesStr)
        {
            try
            {
                if (string.IsNullOrEmpty(displayValuesStr))
                {
                    LogUtility.LogDebug($"{nameof(ParseFromStringFormat)}: displayValuesStr is null or empty for row {_rowNumber}");
                    return;
                }

                // Parse format: {----} {----} or {00--1000--} {T--zzzz--}
                var parts = displayValuesStr.Split(new[] { "   " }, StringSplitOptions.None);
                if (parts.Length != 2)
                {
                    LogUtility.LogDebug($"{nameof(ParseFromStringFormat)}: Expected 2 parts separated by '   ', got {parts.Length} for row {_rowNumber}");
                    return;
                }

                string lowPart = parts[0].Trim().TrimStart('{').TrimEnd('}');
                string highPart = parts[1].Trim().TrimStart('{').TrimEnd('}');

                // Determine segment count
                int segmentCount = 0;

                // Check if the string contains only '-' characters (all empty values)
                bool allEmptyLow = lowPart.All(c => c == '-');
                bool allEmptyHigh = highPart.All(c => c == '-');

                if (allEmptyLow && allEmptyHigh)
                {
                    // All segments are empty
                    // The number of segments = number of hyphens + 1
                    // For "----": 4 hyphens + 1 = 5 segments
                    segmentCount = lowPart.Length + 1;

                    LogUtility.LogDebug($"{nameof(ParseFromStringFormat)}: All segments empty. Detected {segmentCount} segments from '{lowPart}' (length {lowPart.Length} + 1)");
                }
                else
                {
                    // There are real values, count delimiters + 1
                    // For "00--1000--": split by '-' gives ["00", "", "1000", "", ""] = 5 segments
                    var splitLow = lowPart.Split(new[] { '-' }, StringSplitOptions.None);
                    var splitHigh = highPart.Split(new[] { '-' }, StringSplitOptions.None);
                    segmentCount = Math.Max(splitLow.Length, splitHigh.Length);

                    LogUtility.LogDebug($"{nameof(ParseFromStringFormat)}: Detected {segmentCount} segments from values");
                }

                // Parse the actual values
                var lowValuesParsed = lowPart.Split(new[] { '-' }, StringSplitOptions.None);
                var highValuesParsed = highPart.Split(new[] { '-' }, StringSplitOptions.None);

                // Create segments
                for (int i = 0; i < segmentCount; i++)
                {
                    string segmentName = $"SEGMENT{i + 1}";

                    string low = i < lowValuesParsed.Length ? lowValuesParsed[i] : "-";
                    string high = i < highValuesParsed.Length ? highValuesParsed[i] : "-";

                    var segment = new SegmentData
                    {
                        SegmentName = segmentName,
                        Low = low != "-" ? low : string.Empty,
                        High = high != "-" ? high : string.Empty
                    };
                    _segments.Add(segment);
                }

                LogUtility.LogDebug($"{nameof(ParseFromStringFormat)}: Successfully parsed {_segments.Count} segments from string format for row {_rowNumber}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(ParseFromStringFormat)}: Failed to parse string format for row {_rowNumber}. Input: {displayValuesStr}");
            }
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var lowValueList = new List<string>();
                var highValueList = new List<string>();

                bool allLowEmpty = true;
                bool allHighEmpty = true;

                foreach (var segment in _segments)
                {
                    // Store actual values, empty as empty string
                    string low = segment.Low?.Trim() ?? string.Empty;
                    string high = segment.High?.Trim() ?? string.Empty;
                    lowValueList.Add(low);
                    highValueList.Add(high);

                    if (!string.IsNullOrEmpty(low)) allLowEmpty = false;
                    if (!string.IsNullOrEmpty(high)) allHighEmpty = false;
                }

                string lowResult;
                string highResult;

                // Special handling for all-empty segments
                if (allLowEmpty)
                {
                    // All low values are empty - create string with hyphens only
                    // For 5 segments: "----" (4 hyphens = 5 segments)
                    lowResult = "{" + new string('-', _segments.Count - 1) + "}";
                }
                else
                {
                    // Join with '-' delimiter - empty values become empty strings
                    // For ["10", "", "", "", ""] -> "10----" (4 hyphens)
                    lowResult = "{" + string.Join("-", lowValueList) + "}";
                }

                if (allHighEmpty)
                {
                    // All high values are empty - create string with hyphens only
                    // For 5 segments: "----" (4 hyphens = 5 segments)
                    highResult = "{" + new string('-', _segments.Count - 1) + "}";
                }
                else
                {
                    highResult = "{" + string.Join("-", highValueList) + "}";
                }

                string result = lowResult + "   " + highResult;

                LogUtility.LogDebug($"{nameof(BtnOk_Click)}: Generated segment string for row {_rowNumber}: {result}");

                Excel.Range valueCell = _controlSheet.Cells[_rowNumber, 10] as Excel.Range;
                if (valueCell != null)
                {
                    valueCell.Value2 = result;
                }
                else
                {
                    // No exception is thrown here, so a LogWarn (always written regardless of
                    // debug mode) is the only record that Column J was silently left unwritten.
                    LogUtility.LogWarn($"{nameof(BtnOk_Click)}: Failed to get Column J cell for row {_rowNumber} - value not written.");
                }

                UpdateIACell();

                DialogResult = true;
                Close();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(BtnOk_Click)}: Failed to update segment values for row {_rowNumber}");
                MessageFunctions.XLEdgeMessage("Error updating segment values.",
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void UpdateIACell()
        {
            try
            {
                Excel.Range iaCell = _controlSheet.Cells[_rowNumber, 235] as Excel.Range;
                if (iaCell == null)
                {
                    // No exception is thrown here, so a LogWarn (always written regardless of
                    // debug mode) is the only record that the IA cell was silently left unwritten.
                    LogUtility.LogWarn($"{nameof(UpdateIACell)}: Failed to get IA cell for row {_rowNumber} - value not written.");
                    return;
                }

                // Get the original segment names from IA cell
                List<string> segmentNames = GetSegmentNames(iaCell?.Value2?.ToString());

                var segmentDict = new Dictionary<string, string>();

                // Match UI segments with the order from IA cell
                for (int i = 0; i < segmentNames.Count && i < _segments.Count; i++)
                {
                    var segment = _segments[i];
                    string low = string.IsNullOrEmpty(segment.Low?.Trim()) ? "-" : segment.Low.Trim();
                    string high = string.IsNullOrEmpty(segment.High?.Trim()) ? "-" : segment.High.Trim();

                    string op;

                    if (low == "-" && high == "-")
                        op = "-";
                    else
                    {
                        op= $"{low}-{high}";
                    }

                    segmentDict[segmentNames[i]] = op;
                }

                // If there are more segments in UI than in original, add them with generic names
                for (int i = segmentNames.Count; i < _segments.Count; i++)
                {
                    var segment = _segments[i];
                    string low = string.IsNullOrEmpty(segment.Low?.Trim()) ? "-" : segment.Low.Trim();
                    string high = string.IsNullOrEmpty(segment.High?.Trim()) ? "-" : segment.High.Trim();
                    segmentDict[$"SEGMENT{i + 1}"] = $"{low}-{high}";
                }

                //string json = System.Text.Json.JsonSerializer.Serialize(segmentDict, JsonGlobals.Options);
                string json = SerializationHelper.SerializeToJson(segmentDict);
                iaCell.Value2 = json;

                LogUtility.LogDebug($"{nameof(UpdateIACell)}: Updated IA cell for row {_rowNumber} with JSON: {json}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(UpdateIACell)}: Failed to update IA cell for row {_rowNumber}");
            }
        }

        private List<string> GetSegmentNames(string iaCellValue)
        {
            var names = new List<string>();
            try
            {
                if (!string.IsNullOrEmpty(iaCellValue))
                {
                    using var doc = JsonDocument.Parse(iaCellValue);
                    var root = doc.RootElement;
                    if (root.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var prop in root.EnumerateObject())
                        {
                            names.Add(prop.Name);
                        }
                        LogUtility.LogDebug($"{nameof(GetSegmentNames)}: Found {names.Count} segment names from IA cell");
                    }
                    else
                    {
                        LogUtility.LogDebug($"{nameof(GetSegmentNames)}: IA cell value is not a JSON object: {iaCellValue}");
                    }
                }
                else
                {
                    LogUtility.LogDebug($"{nameof(GetSegmentNames)}: IA cell value is null or empty");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"{nameof(GetSegmentNames)}: Failed to parse IA cell value: {iaCellValue}");
            }

            // If no names found, create generic names based on segment count from UI
            if (names.Count == 0)
            {
                LogUtility.LogDebug($"{nameof(GetSegmentNames)}: No segment names found, using generic names");
                for (int i = 0; i < _segments.Count; i++)
                {
                    names.Add($"SEGMENT{i + 1}");
                }
            }
            return names;
        }

        // Combined Close/Cancel button handler
        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}