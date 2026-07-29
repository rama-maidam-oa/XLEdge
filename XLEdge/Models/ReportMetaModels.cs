using System.Collections.Generic;
using System.Text.Json.Serialization;
using XLEdge.Helpers;

namespace XLEdge.Models
{
    // Ported from ReportMetaInfo.vb. "BroadcastMessage" (braodcastMsg) and "XLEdgeUserPreferences"
    // (xledgeuserPreferences) already exist in AllModels.cs - not duplicated here.
    //
    // These describe the report-metadata JSON contract from the server (columns, drilldowns,
    // parameters) and the drill-submit request contract. Needed once the drilldown
    // (SheetFollowHyperlink) and parameter-control-sheet logic are ported - see MIGRATION_STATUS.md.

    public sealed class OutputProp
    {
        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("hlinkDisplayValue")]
        public string HlinkDisplayValue { get; set; }

        [JsonPropertyName("imgWidth")]
        public int ImgWidth { get; set; }

        [JsonPropertyName("imgHeight")]
        public int ImgHeight { get; set; }
    }

    public sealed class ColumnProperties
    {
        [JsonPropertyName("fmt")]
        public string Fmt { get; set; }

        [JsonPropertyName("hdn")]
        public bool Hidden { get; set; }

        [JsonPropertyName("outputprop")]
        public OutputProp OutputProp { get; set; }
    }

    public sealed class RptColumn
    {
        [JsonPropertyName("categoryId")]
        public int CategoryId { get; set; }

        [JsonPropertyName("columnId")]
        public int ColumnId { get; set; }

        [JsonPropertyName("datatype")]
        public string DataType { get; set; }

        [JsonPropertyName("dimension")]
        public string Dimension { get; set; }

        [JsonPropertyName("logicalColumnId")]
        public long LogicalColumnId { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("isFileAttached")]
        public bool IsFileAttached { get; set; }

        [JsonPropertyName("isHtmlEncoded")]
        public bool IsHtmlEncoded { get; set; }

        [JsonPropertyName("properties")]
        public ColumnProperties Properties { get; set; }
    }

    public sealed class ChildParameter
    {
        [JsonPropertyName("paramName")]
        public string ParamName { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("logicalColumnId")]
        public long LogicalColumnId { get; set; }

        [JsonPropertyName("staticValue")]
        [JsonConverter(typeof(NumericJsonConverter))]
        public object StaticValue { get; set; }

        [JsonPropertyName("formula")]
        [JsonConverter(typeof(NumericJsonConverter))]
        public object Formula { get; set; }
    }

    public sealed class RptDrilldown
    {
        [JsonPropertyName("drillReportId")]
        public int DrillReportId { get; set; }

        [JsonPropertyName("drillReportName")]
        public string DrillReportName { get; set; }

        [JsonPropertyName("drillColumnName")]
        public string DrillColumnName { get; set; }

        [JsonPropertyName("categoryId")]
        public int CategoryId { get; set; }

        [JsonPropertyName("logicalColumnId")]
        public long LogicalColumnId { get; set; }

        [JsonPropertyName("dimension")]
        public string Dimension { get; set; }

        [JsonPropertyName("parameters")]
        public ChildParameter[] Parameters { get; set; }
    }

    public sealed class RptParameter
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("label")]
        public string Label { get; set; }

        [JsonPropertyName("isRequired")]
        public bool IsRequired { get; set; }

        [JsonPropertyName("dataType")]
        public string DataType { get; set; }
    }

    public sealed class ReportMeta
    {
        [JsonPropertyName("reportId")]
        public int ReportId { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("businessObjectId")]
        public int BusinessObjectId { get; set; }

        [JsonPropertyName("baseReportName")]
        public string BaseReportName { get; set; }

        [JsonPropertyName("columns")]
        public RptColumn[] Columns { get; set; }

        [JsonPropertyName("drilldowns")]
        public RptDrilldown[] Drilldowns { get; set; }

        [JsonPropertyName("parameters")]
        public RptParameter[] Parameters { get; set; }

        // Original VB property name ("procesStartTime") had a typo; kept as-is via
        // JsonPropertyName so deserialization still matches whatever the server actually sends.
        [JsonPropertyName("procesStartTime")]
        [JsonConverter(typeof(NumericJsonConverter))]
        public object ProcessStartTime { get; set; }
        // Plain int - NumericJsonConverter is JsonConverter<object> and is only valid on
        // ProcessStartTime above (typed as object). Applying it here throws
        // "converter ... not compatible with System.Int32" on every deserialize.
        [JsonPropertyName("lockedColumnsCount")]
        public int LockedColumnsCount { get; set; } = 0;
    }

    public sealed class DrillParameter
    {
        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("values")]
        public string[] Values { get; set; }

        [JsonPropertyName("operator")]
        public string Operator { get; set; }
    }

    public sealed class ExtraParameters
    {
        [JsonPropertyName("ORACLE_RESP_ID")]
        public string OracleRespId { get; set; }
    }

    public sealed class DrillSubmit
    {
        [JsonPropertyName("reportId")]
        public string ReportId { get; set; }

        [JsonPropertyName("parameters")]
        public DrillParameter[] Parameters { get; set; }

        [JsonPropertyName("extraParameters")]
        public ExtraParameters ExtraParameters { get; set; }
    }

    /// <summary>Internal working model (not deserialized from server JSON) tracking how a raw
    /// CSV column maps to its renamed Excel column and any drilldowns defined for it.</summary>
    public sealed class ColumnMetadata
    {
        public string OriginalName { get; set; }
        public string RenamedName { get; set; }
        public long LogicalColumnId { get; set; }
        public string PropType { get; set; }
        public List<DrilldownInfo> Drilldowns { get; set; } = new List<DrilldownInfo>();
    }

    public sealed class DrilldownInfo
    {
        public long DrillReportId { get; set; }
        public string DrillReportName { get; set; }
        public long LogicalColumnId { get; set; }
        public string Dimension { get; set; }
        public List<DrilldownParameter> Parameters { get; set; }
    }

    public sealed class DrilldownParameter
    {
        public string ParamName { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
        public string StaticValue { get; set; }
        public long LogicalColumnId { get; set; }
        public string Formula { get; set; }
    }

    // Note: the outbound request body for refreshing/submitting a report with edited parameter
    // values ("ReportParameterRequest"/"ReportParameterValue") lives in XLEdge.Helpers
    // (XLEdgeParamsBuilder.cs) - that's the version every real call site (DrilldownRequestBuilder,
    // ReportParameterRequestSerializer) actually resolves to. A duplicate, unused pair of classes
    // with the same names used to live here too; removed to avoid the same-name shadowing trap
    // that caused the ExtraParameters/NumericJsonConverter bug.
}
