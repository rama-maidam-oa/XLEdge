using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using XLEdge.Models;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    public static class ApiResponseHelper
    {
        private static readonly string[] KnownPayloadNames = new[]
        {
            "records",
            "record",
            "preferences",
            "data",
            "items",
            "value",
            "values",
            "result"
        };

        // =====================================================
        // PUBLIC ENTRY POINT
        // =====================================================
        public static ApiResult<T> Parse<T>(string rawResponse, JsonSerializerOptions options)
        {
            if (string.IsNullOrWhiteSpace(rawResponse))
                return ApiResult<T>.Failure("Empty response from server.");

            if (!IsLikelyJson(rawResponse))
            {
                // Only log the full response body when Debug mode + "Include Output Data" are both
                // on, since a non-JSON response is often a large HTML error page; otherwise log a
                // short length-only warning.
                if (LogUtility.DebugMode && XLEdgeAppState.Instance.DebugOutputData)
                {
                    LogUtility.LogWarn($"ApiResponseHelper | Non-JSON response: {rawResponse}");
                }
                else
                {
                    LogUtility.LogWarn($"ApiResponseHelper | Non-JSON response received ({rawResponse.Length} character(s) - enable Debug mode + 'Include Output Data' to log the full body).");
                }

                string apiMessage = string.Empty;

                if (rawResponse.IndexOf("<!doctype html>", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Invalid response format. Received HTML instead of JSON.";
                }
                else if (rawResponse.IndexOf("InternalServerError", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    apiMessage = "Server encountered an error. Please try again later.";
                }
                else
                {
                    apiMessage = "Received non-JSON response from server.";
                }

                return ApiResult<T>.Failure(apiMessage);
            }

            try
            {
                using var doc = JsonDocument.Parse(rawResponse);
                var root = doc.RootElement;

                // -------------------------------------------------
                // CASE 1: Root is ARRAY → direct payload
                // -------------------------------------------------
                if (root.ValueKind == JsonValueKind.Array)
                {
                    return DeserializePayload<T>(root);
                }

                // -------------------------------------------------
                // CASE 2: Root is OBJECT
                // -------------------------------------------------
                if (root.ValueKind == JsonValueKind.Object)
                {
                    // If contains status → treat as wrapped response
                    if (ContainsProperty(root, "status"))
                    {
                        return HandleWrappedResponse<T>(root);
                    }

                    // No status → assume root is direct payload
                    return DeserializePayload<T>(root);
                }

                return ApiResult<T>.Failure("Unsupported JSON format.");
            }
            catch (JsonException ex)
            {
                LogUtility.LogException(ex, "Invalid JSON received.");
                LogUtility.LogRawJson("ApiResponseHelper.Parse", rawResponse);
                return ApiResult<T>.Failure("Invalid JSON response.");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Unexpected parsing error.");
                return ApiResult<T>.Failure("Unexpected response format.");
            }
        }

        // =====================================================
        // WRAPPED RESPONSE HANDLER
        // =====================================================
        private static ApiResult<T> HandleWrappedResponse<T>(JsonElement root)
        {
            string status = GetStringSafe(root, "status");
            string message = GetStringSafe(root, "message");

            if (string.IsNullOrWhiteSpace(message))
                message = GetStringSafe(root, "msg");

            bool success = string.Equals(
                status,
                "success",
                StringComparison.OrdinalIgnoreCase);

            if (!success)
            {
                return ApiResult<T>.Failure(
                    string.IsNullOrWhiteSpace(message)
                        ? "Server returned failure status."
                        : message);
            }

            // If T is JsonElement → return whole root
            if (typeof(T) == typeof(JsonElement))
            {
                object clone = root.Clone();
                return ApiResult<T>.Success((T)clone);
            }

            // First, try deserializing the entire wrapped object to T
            var wholeResult = DeserializePayload<T>(root);
            if (wholeResult.IsSuccess)
            {
                return wholeResult;
            }

            // Auto-detect payload
            var payload = DetectPayload(root);

            return DeserializePayload<T>(payload);
        }

        // =====================================================
        // AUTO PAYLOAD DETECTION
        // =====================================================
        private static JsonElement DetectPayload(JsonElement root)
        {
            // 1️⃣ Known names
            foreach (var name in KnownPayloadNames)
            {
                var prop = root.EnumerateObject()
                    .FirstOrDefault(p =>
                        string.Equals(p.Name, name,
                            StringComparison.OrdinalIgnoreCase));

                if (!prop.Equals(default(JsonProperty)))
                    return prop.Value;
            }

            // 2️⃣ First non-meta object/array
            foreach (var prop in root.EnumerateObject())
            {
                if (IsMetaProperty(prop.Name))
                    continue;

                if (prop.Value.ValueKind == JsonValueKind.Array ||
                    prop.Value.ValueKind == JsonValueKind.Object)
                    return prop.Value;
            }

            // 3️⃣ Fallback → entire root
            return root;
        }

        // =====================================================
        // DESERIALIZATION CORE
        // =====================================================
        private static ApiResult<T> DeserializePayload<T>(JsonElement element)
        {
            try
            {
                if (IsCollectionType(typeof(T)))
                {
                    if (element.ValueKind != JsonValueKind.Array)
                    {
                        var detected = DetectPayload(element);
                        if (detected.ValueKind == JsonValueKind.Array)
                        {
                            element = detected;
                        }
                        else
                        {
                            LogUtility.LogWarn("Expected array but received object.");
                            return ApiResult<T>.Failure("Invalid payload format for collection.");
                        }
                    }
                }
                else
                {
                    if (element.ValueKind == JsonValueKind.Array)
                    {
                        LogUtility.LogWarn("Expected object but received array.");
                    }
                }

                var result = JsonSerializer.Deserialize<T>(
                    element.GetRawText(),
                    JsonGlobals.Options);

                if (object.Equals(result, default(T)))
                    return ApiResult<T>.Failure("Failed to deserialize response.");

                return ApiResult<T>.Success(result);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Payload deserialization failed.");
                LogUtility.LogRawJson($"ApiResponseHelper.DeserializePayload<{typeof(T).Name}>", element.GetRawText());
                return ApiResult<T>.Failure("Invalid payload format.");
            }
        }

        // =====================================================
        // TYPE DETECTION
        // =====================================================
        private static bool IsCollectionType(Type type)
        {
            if (type == typeof(string))
                return false;

            if (type.IsArray)
                return true;

            return typeof(IEnumerable).IsAssignableFrom(type);
        }

        // =====================================================
        // HELPERS
        // =====================================================
        private static bool IsLikelyJson(string input)
        {
            input = input.TrimStart();
            return input.StartsWith("{") || input.StartsWith("[");
        }

        private static bool ContainsProperty(JsonElement element, string name)
        {
            return element.EnumerateObject()
                .Any(p => string.Equals(p.Name,
                                        name,
                                        StringComparison.OrdinalIgnoreCase));
        }

        private static bool IsMetaProperty(string name)
        {
            return string.Equals(name, "status", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "msg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "message", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "redirectURL", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "domain", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "contextPath", StringComparison.OrdinalIgnoreCase);
        }

        private static string GetStringSafe(JsonElement element, string propertyName)
        {
            var prop = element.EnumerateObject()
                .FirstOrDefault(p =>
                    string.Equals(p.Name,
                                  propertyName,
                                  StringComparison.OrdinalIgnoreCase));

            if (prop.Equals(default(JsonProperty)))
                return string.Empty;

            return prop.Value.ValueKind == JsonValueKind.String
                ? prop.Value.GetString() ?? string.Empty
                : prop.Value.ToString();
        }
    }
}
