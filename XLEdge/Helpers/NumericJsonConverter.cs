using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace XLEdge.Helpers
{
    [ComVisible(false)]
    public sealed class NumericJsonConverter : JsonConverter<object>
    {
        public override object Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
        {
            switch (reader.TokenType)
            {
                case JsonTokenType.Number:
                    if (reader.TryGetInt64(out long longValue))
                    {
                        return longValue;
                    }
                    return reader.GetDouble();

                case JsonTokenType.String:
                    return reader.GetString();

                case JsonTokenType.True:
                case JsonTokenType.False:
                    return reader.GetBoolean();

                case JsonTokenType.Null:
                    return null;

                default:
                    using (JsonDocument doc = JsonDocument.ParseValue(ref reader))
                    {
                        return doc.RootElement.Clone();
                    }
            }
        }

        public override void Write(Utf8JsonWriter writer, object value, JsonSerializerOptions options)
        {
            switch (value)
            {
                case null:
                    writer.WriteNullValue();
                    return;

                case int intValue:
                    writer.WriteNumberValue(intValue);
                    return;

                case long longValue:
                    writer.WriteNumberValue(longValue);
                    return;

                case System.Numerics.BigInteger bigIntValue:
                    writer.WriteRawValue(bigIntValue.ToString(CultureInfo.InvariantCulture));
                    return;

                case double doubleValue:
                    writer.WriteNumberValue(doubleValue);
                    return;

                case decimal decimalValue:
                    writer.WriteNumberValue(decimalValue);
                    return;

                case bool boolValue:
                    writer.WriteBooleanValue(boolValue);
                    return;

                case string stringValue:
                    WriteStringAsNumericIfPossible(writer, stringValue);
                    return;

                case Dictionary<string, object> dict:
                    WriteDictionary(writer, dict, options);
                    return;

                default:
                    // For other complex types, use fallback
                    JsonSerializer.Serialize(writer, value, value.GetType(), GetOptionsWithoutSelf(options));
                    return;
            }
        }

        private static void WriteDictionary(Utf8JsonWriter writer, Dictionary<string, object> dict, JsonSerializerOptions options)
        {
            writer.WriteStartObject();

            foreach (var kvp in dict)
            {
                writer.WritePropertyName(kvp.Key);

                // --- CRITICAL: ORACLE_RESP_ID is ALWAYS a string ---
                if (kvp.Key == "ORACLE_RESP_ID")
                {
                    writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                }
                // --- CRITICAL: ORACLE_GL_SEGMENT_VALUES is ALWAYS a string ---
                else if (kvp.Key == "ORACLE_GL_SEGMENT_VALUES")
                {
                    writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                }
                // --- CRITICAL: ORACLE_RESP_DISPLAY_VALUE is ALWAYS a string ---
                else if (kvp.Key == "ORACLE_RESP_DISPLAY_VALUE")
                {
                    writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                }
                // --- CRITICAL: ORACLE_GL_SEGMENT_DISPLAY_VALUES is ALWAYS a string ---
                else if (kvp.Key == "ORACLE_GL_SEGMENT_DISPLAY_VALUES")
                {
                    writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                }
                else if (kvp.Value is string strVal)
                {
                    WriteStringAsNumericIfPossible(writer, strVal);
                }
                else if (kvp.Value is int intVal)
                {
                    writer.WriteNumberValue(intVal);
                }
                else if (kvp.Value is long longVal)
                {
                    writer.WriteNumberValue(longVal);
                }
                else if (kvp.Value is decimal decVal)
                {
                    writer.WriteNumberValue(decVal);
                }
                else if (kvp.Value is double dblVal)
                {
                    writer.WriteNumberValue(dblVal);
                }
                else if (kvp.Value is bool boolVal)
                {
                    writer.WriteBooleanValue(boolVal);
                }
                else if (kvp.Value == null)
                {
                    writer.WriteNullValue();
                }
                else
                {
                    // For complex nested objects
                    JsonSerializer.Serialize(writer, kvp.Value, kvp.Value.GetType(), GetOptionsWithoutSelf(options));
                }
            }

            writer.WriteEndObject();
        }

        private static readonly ConditionalWeakTable<JsonSerializerOptions, JsonSerializerOptions> _fallbackOptionsCache = new();

        private static JsonSerializerOptions GetOptionsWithoutSelf(JsonSerializerOptions options)
        {
            return _fallbackOptionsCache.GetValue(options, static source =>
            {
                var clone = new JsonSerializerOptions(source);
                for (int i = clone.Converters.Count - 1; i >= 0; i--)
                {
                    if (clone.Converters[i] is NumericJsonConverter)
                    {
                        clone.Converters.RemoveAt(i);
                    }
                }
                return clone;
            });
        }

        private static void WriteStringAsNumericIfPossible(Utf8JsonWriter writer, string stringValue)
        {
            if (stringValue.Contains("."))
            {
                if (double.TryParse(stringValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double doubleValue))
                {
                    writer.WriteNumberValue(doubleValue);
                    return;
                }
            }
            else if (int.TryParse(stringValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out int intValue))
            {
                writer.WriteNumberValue(intValue);
                return;
            }

            writer.WriteStringValue(stringValue);
        }
    }
}