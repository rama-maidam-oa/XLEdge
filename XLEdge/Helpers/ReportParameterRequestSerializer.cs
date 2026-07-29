using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace XLEdge.Helpers
{
    public static class ReportParameterRequestSerializer
    {
        public static string Serialize(ReportParameterRequest request)
        {
            using var stream = new MemoryStream();
            using var writer = new Utf8JsonWriter(stream, new JsonWriterOptions { Indented = false });

            writer.WriteStartObject();

            // 1. reportId (lowercase)
            writer.WritePropertyName("reportId");
            writer.WriteStringValue(request.ReportId?.ToString() ?? string.Empty);

            // 2. parameters (lowercase)
            writer.WritePropertyName("parameters");
            writer.WriteStartArray();
            if (request.Parameters != null)
            {
                foreach (var param in request.Parameters)
                {
                    writer.WriteStartObject();

                    writer.WritePropertyName("name");
                    writer.WriteStringValue(param.Name);

                    // value or values
                    if (param.Values != null && param.Values.Count > 0)
                    {
                        writer.WritePropertyName("values");
                        writer.WriteStartArray();
                        foreach (var val in param.Values)
                        {
                            WriteValue(writer, val);
                        }
                        writer.WriteEndArray();
                    }
                    else
                    {
                        writer.WritePropertyName("value");
                        WriteValue(writer, param.Value);
                    }

                    writer.WritePropertyName("operator");
                    writer.WriteStringValue(param.Operator);

                    writer.WriteEndObject();
                }
            }
            writer.WriteEndArray();

            // 3. extraParameters (lowercase) - null when there's nothing to send, not "{}"
            writer.WritePropertyName("extraParameters");
            var extraParams = request.ExtraParameters as Dictionary<string, object>;
            if (extraParams != null && extraParams.Count > 0)
            {
                writer.WriteStartObject();
                foreach (var kvp in extraParams)
                {
                    writer.WritePropertyName(kvp.Key);
                    // --- CRITICAL: ORACLE_RESP_ID is ALWAYS a string ---
                    writer.WriteStringValue(kvp.Value?.ToString() ?? string.Empty);
                }
                writer.WriteEndObject();
            }
            else
            {
                writer.WriteNullValue();
            }

            writer.WriteEndObject();
            writer.Flush();

            return Encoding.UTF8.GetString(stream.ToArray());
        }

        private static void WriteValue(Utf8JsonWriter writer, object value)
        {
            switch (value)
            {
                case null:
                    writer.WriteNullValue();
                    break;
                case int i:
                    writer.WriteNumberValue(i);
                    break;
                case long l:
                    writer.WriteNumberValue(l);
                    break;
                case double d:
                    writer.WriteNumberValue(d);
                    break;
                case decimal m:
                    writer.WriteNumberValue(m);
                    break;
                case bool b:
                    writer.WriteBooleanValue(b);
                    break;
                case string s:
                    writer.WriteStringValue(s);
                    break;
                default:
                    writer.WriteStringValue(value.ToString());
                    break;
            }
        }
    }
}
