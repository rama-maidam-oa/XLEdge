using System;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace XLEdge.Helpers
{
#nullable enable
    /// <summary>
    /// Global reusable JSON configuration
    /// Used for Serialize / Deserialize operations
    /// </summary>
    public static class JsonGlobals
    {
        public static JsonSerializerOptions Options { get; }

        static JsonGlobals()
        {
            var opts = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                PropertyNamingPolicy = null,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
                ReadCommentHandling = JsonCommentHandling.Skip,
                WriteIndented = false,
                AllowTrailingCommas = true,
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                NumberHandling =
                    JsonNumberHandling.AllowReadingFromString |
                    JsonNumberHandling.AllowNamedFloatingPointLiterals
            };

            // Registered globally (not just via the per-property [JsonConverter] attributes already
            // used on Value/ReportId/ExtraParameters etc.) so bare object-typed collection elements -
            // e.g. ReportParameterValue.Values (List<object>), used for IN/BETWEEN drilldown/refresh
            // parameter values - also get correct numeric-vs-string JSON output instead of falling
            // through to default object serialization (which has no BigInteger support and would
            // otherwise write every element as a JSON string). Property-level attributes still take
            // precedence where present; this only fills the gap for object-typed values with none.
            opts.Converters.Add(new NumericJsonConverter());

#if NET8_0_OR_GREATER
            opts.MakeReadOnly();
#endif

            Options = opts;
        }
    }

    /// <summary>
    /// Safe JSON parsing helpers for dynamic API responses
    /// </summary>
    public static class JsonHelper
    {
        public static bool TryGetProperty(
            JsonElement element,
            string propertyName,
            out JsonElement value)
        {
            if (element.ValueKind != JsonValueKind.Object)
            {
                value = default;
                return false;
            }

            var match = element
                .EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name, propertyName,
                        StringComparison.OrdinalIgnoreCase));

            if (!match.Equals(default(JsonProperty)))
            {
                value = match.Value;
                return true;
            }

            value = default;
            return false;
        }

        public static bool TryGetDouble(JsonElement element, out double result)
        {
            result = 0;

            switch (element.ValueKind)
            {
                case JsonValueKind.Number:
                    return element.TryGetDouble(out result);

                case JsonValueKind.String:
                    return double.TryParse(
                        element.GetString(),
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out result);

                default:
                    return false;
            }
        }

        public static bool TryGetString(JsonElement element, out string? result)
        {
            result = null;

            if (element.ValueKind == JsonValueKind.String)
            {
                result = element.GetString();
                return !string.IsNullOrWhiteSpace(result);
            }

            return false;
        }
        public static string GetStringSafe(JsonElement element, string propertyName)
        {
            if (element.ValueKind != JsonValueKind.Object)
                return string.Empty;

            var property = element
                .EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name,
                                  propertyName,
                                  StringComparison.OrdinalIgnoreCase));

            if (!property.Equals(default(JsonProperty)))
            {
                var value = property.Value;

                return value.ValueKind == JsonValueKind.String
                    ? value.GetString() ?? string.Empty
                    : value.ToString();
            }

            return string.Empty;
        }
    }
#nullable restore
}
