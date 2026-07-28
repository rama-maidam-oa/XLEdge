using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace XLEdge.Models
{
    public class BroadcastMessage
    {
        public string MsgType { get; set; }
        public string Message { get; set; }
    }
    public sealed class XLEdgeUserPreferences
    {
        [JsonPropertyName("parameterValues")]
        public bool ParameterValues { get; set; }

        [JsonPropertyName("scheduledOutputs")]
        public bool ScheduledOutputs { get; set; }

        [JsonPropertyName("refreshSync")]
        public bool RefreshSync { get; set; }

        [JsonPropertyName("changeSheetName")]
        public bool ChangeSheetName { get; set; }

        [JsonPropertyName("calendarCtrlDisplay")]
        public bool CalendarCtrlDisplay { get; set; }

        [JsonPropertyName("overrideFormats")]
        public bool OverrideFormats { get; set; }

        public XLEdgeUserPreferences Clone()
        {
            return new XLEdgeUserPreferences
            {
                ParameterValues = ParameterValues,
                ScheduledOutputs = ScheduledOutputs,
                RefreshSync = RefreshSync,
                ChangeSheetName = ChangeSheetName,
                CalendarCtrlDisplay = CalendarCtrlDisplay,
                OverrideFormats = OverrideFormats
            };
        }
    }
    public class EdgeRequest
    {
        public string ReportType { get; set; }
        public string ReportId { get; set; }
        public string ReportRunId { get; set; }
        public string ReportName { get; set; }
        public string Extension { get; set; }
        public string Extra1 { get; set; }
        public string Extra2 { get; set; }
    }

}
