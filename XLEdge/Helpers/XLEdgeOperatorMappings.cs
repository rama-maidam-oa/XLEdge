using System.Collections.Generic;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Maps the human-readable parameter operator labels shown in the "orb_params_control"
    /// sheet to the operator tokens used in report-refresh request payloads. Ported verbatim
    /// from the "OperatorMappings" shared dictionary in AddinModule.vb - used by the not-yet-ported
    /// RibControlSheet_OnClick (writes these labels into the control sheet) and the parameter-data
    /// builder that reads user edits back out of it.
    /// </summary>
    public static class XLEdgeOperatorMappings
    {
        public static readonly IReadOnlyDictionary<string, string> Map = new Dictionary<string, string>
        {
            { "is equal to", "=" },
            { "does not equal", "<>" },
            { "begins with", "BEGINSWITH" },
            { "does not begin with", "NOT BEGINSWITH" },
            { "ends with", "ENDSWITH" },
            { "does not end with", "NOT ENDSWITH" },
            { "contains", "CONTAINS" },
            { "does not contain", "NOT CONTAINS" },
            { "like", "LIKE" },
            { "not like", "NOT LIKE" },
            { "is greater than", ">" },
            { "is greater than or equal to", ">=" },
            { "is less than", "<" },
            { "is less than or equal to", "<=" },
            { "is between", "BETWEEN" },
            { "is not between", "NOT BETWEEN" },
            { "is in list", "IN" },
            { "is not in list", "NOT IN" },
            { "top", "TOP" },
            { "bottom", "BOTTOM" },
            { "is null", "IS NULL" },
            { "is not null", "IS NOT NULL" },
        };
    }
}
