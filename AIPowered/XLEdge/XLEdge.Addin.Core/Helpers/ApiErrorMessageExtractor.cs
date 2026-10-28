using System;
using System.Text.Json;
using System.Text.RegularExpressions;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Converts a raw error response body (HTML or JSON) from a failed API call into a short,
    /// readable message suitable for showing to the user.
    /// </summary>
    public static class ApiErrorMessageExtractor
    {
        private static readonly string[] ErrorKeywords =
        {
            "error", "exception", "timeout", "failed", "invalid", "not found", "unauthorized", "forbidden"
        };

        private static readonly string[] JsonObjectMessageKeys =
        {
            "message", "msg", "error", "errorMessage", "error_description", "detail",
            "Message", "Msg", "status", "statusMessage", "fault", "reason"
        };

        private static readonly string[] JsonArrayItemMessageKeys =
        {
            "message", "msg", "error", "description"
        };

        /// <summary>Ported from ExtractErrorMessage.</summary>
        public static string ExtractErrorMessage(string responseBody)
        {
            try
            {
                if (string.IsNullOrEmpty(responseBody))
                {
                    return "An unknown error occurred.";
                }

                string trimmedBody = responseBody.Trim();

                if (IsHtmlResponse(trimmedBody))
                {
                    return ExtractTextFromHtml(trimmedBody);
                }

                if (trimmedBody.StartsWith("{") || trimmedBody.StartsWith("["))
                {
                    return ExtractFromJson(trimmedBody);
                }

                string cleanedText = CleanPlainText(trimmedBody);
                if (!string.IsNullOrEmpty(cleanedText))
                {
                    return cleanedText;
                }

                return "An error occurred while processing the request.";
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ExtractErrorMessage));
                return "An unknown error occurred.";
            }
        }

        /// <summary>Ported from IsHtmlResponse.</summary>
        public static bool IsHtmlResponse(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            string lowerText = text.ToLowerInvariant().Trim();

            if (lowerText.StartsWith("<!doctype") ||
                lowerText.StartsWith("<html") ||
                lowerText.StartsWith("<head") ||
                lowerText.StartsWith("<body") ||
                lowerText.StartsWith("<?xml"))
            {
                return true;
            }

            if (lowerText.Contains("<html") ||
                lowerText.Contains("<head") ||
                lowerText.Contains("<body") ||
                lowerText.Contains("<!doctype"))
            {
                return true;
            }

            return false;
        }

        /// <summary>Ported from ExtractTextFromHtml.</summary>
        private static string ExtractTextFromHtml(string html)
        {
            try
            {
                // Remove HTML tags
                string plainText = Regex.Replace(html, "<[^>]+>", " ");

                // Decode HTML entities
                plainText = System.Net.WebUtility.HtmlDecode(plainText);

                // Clean up whitespace
                plainText = Regex.Replace(plainText, "\\s+", " ").Trim();

                // Look for common error indicators
                foreach (string keyword in ErrorKeywords)
                {
                    if (plainText.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string[] sentences = plainText.Split(new[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string sentence in sentences)
                        {
                            if (sentence.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return sentence.Trim();
                            }
                        }
                    }
                }

                // If no error keywords found, return first sentence
                if (plainText.Length > 0)
                {
                    string[] sentences = plainText.Split(new[] { '.', '!', '?' }, StringSplitOptions.RemoveEmptyEntries);
                    if (sentences.Length > 0)
                    {
                        string firstSentence = sentences[0].Trim();
                        if (firstSentence.Length > 0)
                        {
                            return firstSentence;
                        }
                    }

                    return plainText.Substring(0, Math.Min(200, plainText.Length)).Trim();
                }

                return "A server error occurred.";
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort HTML-error-page text extraction; a malformed/unexpected
                // HTML body just falls back to a generic message instead of surfacing to the user.
                LogUtility.LogDebug($"{nameof(ExtractTextFromHtml)}: failed to extract text from HTML error response - {ex.Message}");
                return "A server error occurred.";
            }
        }

        /// <summary>Ported from ExtractFromJson - uses System.Text.Json instead of Newtonsoft.</summary>
        private static string ExtractFromJson(string json)
        {
            try
            {
                if (json.StartsWith("{"))
                {
                    return ExtractFromJsonObject(json);
                }

                if (json.StartsWith("["))
                {
                    return ExtractFromJsonArray(json);
                }

                return "An error occurred.";
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort JSON error-message parse; falls back to returning the
                // cleaned raw text instead of surfacing to the user.
                LogUtility.LogDebug($"{nameof(ExtractFromJson)}: failed to parse JSON error response - {ex.Message}");
                return CleanPlainText(json);
            }
        }

        /// <summary>
        /// Extracted from ExtractFromJson's "{" branch: parses a JSON object body and checks, in order,
        /// the common message fields, "errors" array, "validationErrors" array, and "fieldErrors" object.
        /// Any parse failure here propagates up to ExtractFromJson's own catch block, matching the
        /// original (unsplit) behavior where this code ran inline inside that same try block.
        /// </summary>
        private static string ExtractFromJsonObject(string json)
        {
            using JsonDocument doc = JsonDocument.Parse(json);
            JsonElement root = doc.RootElement;

            string commonMessage = TryGetCommonErrorMessage(root);
            if (!string.IsNullOrEmpty(commonMessage))
            {
                return commonMessage;
            }

            string errorsArrayMessage = TryGetErrorsArrayMessage(root);
            if (errorsArrayMessage != null)
            {
                return errorsArrayMessage;
            }

            string validationErrorMessage = TryGetValidationErrorMessage(root);
            if (validationErrorMessage != null)
            {
                return validationErrorMessage;
            }

            string fieldErrorsMessage = TryGetFieldErrorsMessage(root);
            if (fieldErrorsMessage != null)
            {
                return fieldErrorsMessage;
            }

            return "An error occurred. Please check the server logs for details.";
        }

        /// <summary>Checks the common error-message fields (JsonObjectMessageKeys) on a JSON object root.</summary>
        private static string TryGetCommonErrorMessage(JsonElement root)
        {
            // Check for common error message fields
            foreach (string key in JsonObjectMessageKeys)
            {
                if (root.TryGetProperty(key, out JsonElement valueEl))
                {
                    string value = JsonElementToStringValue(valueEl);
                    if (!string.IsNullOrEmpty(value))
                    {
                        return CleanPlainText(value);
                    }
                }
            }

            return null;
        }

        /// <summary>Checks a JSON object root's "errors" array for a message. Returns null if not present/found.</summary>
        private static string TryGetErrorsArrayMessage(JsonElement root)
        {
            // Check for errors array
            if (root.TryGetProperty("errors", out JsonElement errorsEl) && errorsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement errorItem in errorsEl.EnumerateArray())
                {
                    if (errorItem.ValueKind == JsonValueKind.Object)
                    {
                        if (errorItem.TryGetProperty("message", out JsonElement msgEl))
                        {
                            string msg = JsonElementToStringValue(msgEl);
                            if (!string.IsNullOrEmpty(msg))
                            {
                                return CleanPlainText(msg);
                            }
                        }

                        if (errorItem.TryGetProperty("code", out JsonElement codeEl) &&
                            errorItem.TryGetProperty("message", out JsonElement msgEl2))
                        {
                            return $"{JsonElementToStringValue(codeEl)}: {JsonElementToStringValue(msgEl2)}";
                        }
                    }
                    else if (errorItem.ValueKind == JsonValueKind.String)
                    {
                        return errorItem.GetString();
                    }
                }
            }

            return null;
        }

        /// <summary>Checks a JSON object root's "validationErrors" array (first element only) for a message.</summary>
        private static string TryGetValidationErrorMessage(JsonElement root)
        {
            // Check for validation errors
            if (root.TryGetProperty("validationErrors", out JsonElement valErrorsEl) && valErrorsEl.ValueKind == JsonValueKind.Array)
            {
                foreach (JsonElement firstError in valErrorsEl.EnumerateArray())
                {
                    if (firstError.ValueKind == JsonValueKind.Object && firstError.TryGetProperty("message", out JsonElement fmEl))
                    {
                        return CleanPlainText(JsonElementToStringValue(fmEl));
                    }

                    break; // VB only ever looks at valErrors(0)
                }
            }

            return null;
        }

        /// <summary>Checks a JSON object root's "fieldErrors" object for the first field with a usable message.</summary>
        private static string TryGetFieldErrorsMessage(JsonElement root)
        {
            // Check for field errors
            if (root.TryGetProperty("fieldErrors", out JsonElement fieldErrorsEl) && fieldErrorsEl.ValueKind == JsonValueKind.Object)
            {
                foreach (JsonProperty prop in fieldErrorsEl.EnumerateObject())
                {
                    if (prop.Value.ValueKind == JsonValueKind.Array)
                    {
                        var arrayEnumerator = prop.Value.EnumerateArray();
                        if (arrayEnumerator.MoveNext())
                        {
                            string errorMsg = JsonElementToStringValue(arrayEnumerator.Current);
                            if (!string.IsNullOrEmpty(errorMsg))
                            {
                                return CleanPlainText(errorMsg);
                            }

                            return CleanPlainText($"{prop.Name} has an error");
                        }
                    }
                    else if (prop.Value.ValueKind == JsonValueKind.String)
                    {
                        return CleanPlainText(prop.Value.GetString());
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Extracted from ExtractFromJson's "[" branch: parses a JSON array body and checks the first
        /// element's common message fields (JsonArrayItemMessageKeys). Keeps its own try/catch exactly
        /// as in the original inline code.
        /// </summary>
        private static string ExtractFromJsonArray(string json)
        {
            try
            {
                using JsonDocument doc = JsonDocument.Parse(json);
                JsonElement root = doc.RootElement;
                var arrayEnumerator = root.EnumerateArray();
                if (arrayEnumerator.MoveNext())
                {
                    JsonElement firstItem = arrayEnumerator.Current;
                    if (firstItem.ValueKind == JsonValueKind.Object)
                    {
                        foreach (string key in JsonArrayItemMessageKeys)
                        {
                            if (firstItem.TryGetProperty(key, out JsonElement valueEl))
                            {
                                return CleanPlainText(JsonElementToStringValue(valueEl));
                            }
                        }
                    }
                }

                return "An error occurred.";
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort JSON-array error-message parse; malformed/unexpected
                // shape falls back to a generic message instead of surfacing to the user.
                LogUtility.LogDebug($"{nameof(ExtractFromJson)}: failed to parse JSON array error response - {ex.Message}");
                return "An error occurred.";
            }
        }

        /// <summary>Ported from CleanPlainText.</summary>
        private static string CleanPlainText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            // Remove extra whitespace, newlines, tabs
            string cleaned = Regex.Replace(text, "\\s+", " ").Trim();

            // Remove quotes if present
            cleaned = cleaned.Trim('"', '\'');

            // Decode HTML entities
            cleaned = System.Net.WebUtility.HtmlDecode(cleaned);

            // If the text is too long, truncate it
            if (cleaned.Length > 500)
            {
                cleaned = cleaned.Substring(0, 500) + "...";
            }

            return cleaned;
        }

        private static string JsonElementToStringValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString(),
                JsonValueKind.Null or JsonValueKind.Undefined => null,
                _ => element.GetRawText()
            };
        }
    }
}
