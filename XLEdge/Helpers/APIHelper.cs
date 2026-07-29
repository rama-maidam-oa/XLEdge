using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.NetworkInformation;
using System.Security.Authentication;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    public static class ApiHelper
    {
        // Remove shared client - create new clients per request to avoid connection issues
        private static readonly object _lock = new object();

        static ApiHelper()
        {
            // Configure ServicePointManager for better connection handling
            ServicePointManager.DefaultConnectionLimit = 100;
            ServicePointManager.Expect100Continue = false;
            ServicePointManager.MaxServicePointIdleTime = 5000; // 5 seconds idle timeout
            ServicePointManager.SetTcpKeepAlive(false, 0, 0); // Disable TCP keep-alive
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;
        }

        private static HttpClient CreateHttpClient()
        {
            var handler = new HttpClientHandler
            {
                SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                MaxConnectionsPerServer = 10,
                UseProxy = false,
                AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
            };

            var client = new HttpClient(handler)
            {
                Timeout = TimeSpan.FromMinutes(5) // Set reasonable timeout instead of Infinite
            };

            client.DefaultRequestHeaders.ExpectContinue = false;
            client.DefaultRequestHeaders.ConnectionClose = true; // Force connection close after request

            return client;
        }

        public static async Task<string> ServerAPI(string sendURL, string StrContentType, string PostData = "", string MethodType = "POST", CancellationToken cancellationToken = default)
        {
            using (new LogUtility.LogScope("ServerAPI"))
            using (var perfScope = PerformanceHelper.MeasureExecutionTime("API Call"))
            {
                // Use retry logic for transient failures
                return await ApiOperationHelper.ExecuteWithRetry(
                    async (token) =>
                    {
                        using (var client = CreateHttpClient())
                        {
                            return await ExecuteApiCall(client, sendURL, StrContentType, PostData, MethodType, token);
                        }
                    },
                    cancellationToken,
                    $"API call to {sendURL}"
                ).ConfigureAwait(false);
            }
        }

        /// <summary>
        /// Notifies the server that a report run was cancelled client-side, so it stops processing
        /// server-side. Best-effort and fire-and-forget: failures are logged, never thrown, since this
        /// runs from cancellation-handling paths that must not block the UI or surface further errors.
        /// </summary>
        public static async Task NotifyCancelRunAsync(string loginUrl, string runId)
        {
            if (string.IsNullOrWhiteSpace(loginUrl) || string.IsNullOrWhiteSpace(runId))
            {
                return;
            }

            string cancelUrl = $"{loginUrl.TrimEnd('/')}/rest/secure/report/cancel-run?runId={runId}";

            try
            {
                using (var client = CreateHttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(30);
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", XLEdgeAppState.Instance.LoginToken);

                    using (var request = new HttpRequestMessage(HttpMethod.Post, cancelUrl))
                    using (HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            string body = string.Empty;
                            try { body = await response.Content.ReadAsStringAsync().ConfigureAwait(false); }
                            catch (Exception readEx) { LogUtility.LogDebug($"NotifyCancelRunAsync|best-effort diagnostic only - failed to read error response body: {readEx.Message}"); }

                            LogUtility.LogWarn($"NotifyCancelRunAsync|Server returned {(int)response.StatusCode} for cancel-run (runId={runId}): {body}");
                        }
                        else
                        {
                            LogUtility.LogDebug($"NotifyCancelRunAsync|Cancel-run notification sent for runId={runId}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"NotifyCancelRunAsync failed for runId={runId}");
            }
        }

        /// <summary>
        /// Downloads an authenticated attachment to the user's Downloads folder, naming it from the
        /// response's Content-Disposition header. Returns the saved file's full path, or null on failure.
        /// </summary>
        public static async Task<string> DownloadFileAsync(string url, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return null;
            }

            try
            {
                using (var client = CreateHttpClient())
                {
                    client.Timeout = TimeSpan.FromMinutes(10);
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", XLEdgeAppState.Instance.LoginToken);

                    using (var request = new HttpRequestMessage(HttpMethod.Get, url))
                    using (HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            LogUtility.LogWarn($"DownloadFileAsync|Server returned {(int)response.StatusCode} for {url}");
                            return null;
                        }

                        string fileName = ExtractFileNameFromContentDisposition(response) ?? $"attachment_{DateTime.Now.Ticks}";
                        string downloadsFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                        Directory.CreateDirectory(downloadsFolder);
                        string destinationPath = Path.Combine(downloadsFolder, fileName);

                        using (Stream responseStream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false))
                        using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            await responseStream.CopyToAsync(fileStream).ConfigureAwait(false);
                        }

                        return destinationPath;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"DownloadFileAsync cancelled for {url}");
                throw;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"DownloadFileAsync failed for {url}");
                return null;
            }
        }

        private static string ExtractFileNameFromContentDisposition(HttpResponseMessage response)
        {
            try
            {
                string headerValue = response.Content?.Headers?.ContentDisposition?.FileNameStar
                    ?? response.Content?.Headers?.ContentDisposition?.FileName;

                if (string.IsNullOrWhiteSpace(headerValue) &&
                    response.Content?.Headers != null &&
                    response.Content.Headers.TryGetValues("Content-Disposition", out IEnumerable<string> values))
                {
                    // Fallback: manually parse the raw header value if HttpContentHeaders couldn't.
                    string raw = values.FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(raw) && raw.Contains("="))
                    {
                        headerValue = raw.Substring(raw.IndexOf('=') + 1);
                    }
                }

                if (string.IsNullOrWhiteSpace(headerValue))
                {
                    return null;
                }

                headerValue = headerValue.Trim('"');

                string fileName = headerValue;
                if (fileName.Contains("/"))
                {
                    string[] segments = fileName.Split('/');
                    fileName = segments[segments.Length - 1];
                }

                foreach (char invalidChar in Path.GetInvalidFileNameChars())
                {
                    fileName = fileName.Replace(invalidChar, '_');
                }

                return fileName;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ExtractFileNameFromContentDisposition));
                return null;
            }
        }

        private static async Task<string> ExecuteApiCall(HttpClient client, string sendURL, string StrContentType, string PostData, string MethodType, CancellationToken cancellationToken)
        {
            try
            {
                var content = CreateHttpContent(PostData, StrContentType);
                LogRequestDetails(sendURL, MethodType, StrContentType, PostData, client.Timeout);

                cancellationToken.ThrowIfCancellationRequested();

                // Set authorization header for this request
                lock (_lock)
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", XLEdgeAppState.Instance.LoginToken);
                }

                var response = await SendRequestAsync(client, sendURL, MethodType, content, cancellationToken);

                // Read response as stream in chunks to avoid large single-buffer reads
                var responseBody = await ReadResponseStreamAsStringAsync(response, cancellationToken).ConfigureAwait(false);

                LogResponseDetails(response, responseBody);

                var result = ProcessResponse(response, responseBody);

                // A non-success HTTP status is treated as an error immediately rather than passed to
                // EnsureValidApiResponse's text-pattern checks. The error is logged and thrown here;
                // callers are responsible for displaying it to the user.
                if (!response.IsSuccessStatusCode)
                {
                    LogUtility.LogError($"API call to {sendURL} failed with status {(int)response.StatusCode} {response.StatusCode}: {result}");
                    throw new ApiRequestException(result, response.StatusCode);
                }

                ApiOperationHelper.EnsureValidApiResponse(result, $"API call to {sendURL}");

                return result;
            }
            catch (OperationCanceledException ex) when (!cancellationToken.IsCancellationRequested)
            {
                // HttpClient's own Timeout expired (distinct from caller-driven cancellation, since the
                // supplied CancellationToken was never triggered). Thrown as a dedicated ApiTimeoutException
                // so callers can tell a timeout apart from a user-initiated cancel.
                LogUtility.LogWarn($"API call timed out for {sendURL}");
                throw new ApiTimeoutException("The request timed out", ex);
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn($"API call cancelled for {sendURL}");
                throw;
            }
            catch (HttpRequestException ex) when (ex.InnerException is WebException webEx &&
                   (webEx.Status == WebExceptionStatus.KeepAliveFailure ||
                    webEx.Status == WebExceptionStatus.ConnectionClosed))
            {
                // These are transient errors that should be retried
                LogUtility.LogWarn($"Connection closed by server for {sendURL}: {ex.Message}");
                throw new HttpRequestException("Connection was closed by server (transient)", ex);
            }
            catch (Exception ex)
            {
                // Single summary line only. The full exception (with stack trace) is logged once by
                // whichever caller ultimately handles/displays this failure (e.g. ReportGenerator,
                // XLEdgeCTP's broadcast fetch) - re-dumping the full trace here as well just duplicates
                // it, since ApiOperationHelper.ExecuteWithRetry and that caller both see this same
                // exception again as it propagates up.
                LogUtility.LogWarn($"API call failed for {sendURL}: {ex.Message}");
                throw;
            }
        }

        private static HttpContent CreateHttpContent(string postData, string contentType)
        {
            if (string.IsNullOrWhiteSpace(postData))
                return null;

            var mediaType = contentType == "JSON" ? "application/json" : "application/x-www-form-urlencoded";
            return new StringContent(postData, Encoding.UTF8, mediaType);
        }

        private static void LogRequestDetails(string url, string method, string contentType, string payload, TimeSpan timeout)
        {
            LogUtility.LogDebug($"API Request: {method} {url} (ContentType: {contentType}, Timeout: {timeout})");

            if (!string.IsNullOrWhiteSpace(payload))
            {
                LogPayloadChunks(payload);
            }
        }

        private static void LogPayloadChunks(string payload)
        {
            const int maxLogLength = 2000;

            if (payload.Length > maxLogLength)
            {
                LogUtility.LogDebug($"Payload (truncated {maxLogLength} chars of {payload.Length}): {payload.Substring(0, maxLogLength)}...");
            }
            else
            {
                LogUtility.LogDebug($"Payload: {payload}");
            }
        }

        private static async Task<HttpResponseMessage> SendRequestAsync(HttpClient client, string url, string method, HttpContent content, CancellationToken cancellationToken)
        {
            var httpMethod = new HttpMethod(method.ToUpperInvariant());
            using var request = new HttpRequestMessage(httpMethod, url)
            {
                Content = content
            };

            // Add Connection: close header explicitly
            request.Headers.ConnectionClose = true;

            return await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        }

        private static async Task<string> ReadResponseStreamAsStringAsync(HttpResponseMessage response, CancellationToken cancellationToken)
        {
            if (response?.Content == null)
                return string.Empty;

            // Try to determine encoding from response headers, default to UTF8
            Encoding encoding = Encoding.UTF8;
            try
            {
                var charset = response.Content.Headers.ContentType?.CharSet;
                if (!string.IsNullOrWhiteSpace(charset))
                {
                    charset = charset.Trim('"');
                    encoding = Encoding.GetEncoding(charset);
                }
            }
            catch (Exception ex)
            {
                // Safe to ignore: charset parsing is best-effort only; falling back to UTF8 is a
                // reasonable, expected default when the server's charset header is missing/malformed.
                LogUtility.LogDebug($"{nameof(ReadResponseStreamAsStringAsync)}: failed to parse response charset, defaulting to UTF8 - {ex.Message}");
                encoding = Encoding.UTF8;
            }

            // Read the response stream in chunks
            using var stream = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 65536, leaveOpen: false);
            var sb = new StringBuilder();
            var buffer = new char[65536];
            int read;
            while ((read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
            {
                cancellationToken.ThrowIfCancellationRequested();
                sb.Append(buffer, 0, read);
            }

            return sb.ToString();
        }

        private static void LogResponseDetails(HttpResponseMessage response, string responseBody)
        {
            LogUtility.LogDebug($"API Response: {(int)response.StatusCode} {response.StatusCode} (ContentType: {response.Content?.Headers?.ContentType?.MediaType ?? "N/A"}, ContentLength: {response.Content?.Headers?.ContentLength ?? 0})");

            if (!string.IsNullOrWhiteSpace(responseBody))
            {
                const int maxLogLength = 2000;
                if (responseBody.Length > maxLogLength)
                {
                    LogUtility.LogDebug($"Response (truncated {maxLogLength} chars of {responseBody.Length}): {responseBody.Substring(0, maxLogLength)}...");
                }
                else
                {
                    LogUtility.LogDebug($"Response: {responseBody}");
                }
            }
        }

        private static string ProcessResponse(HttpResponseMessage response, string responseBody)
        {
            if (response.IsSuccessStatusCode)
                return CleanResponse(responseBody);

            // Turns an HTML/JSON error body into a short, readable message instead of surfacing the
            // raw response body or just the bare HTTP status code name.
            string extractedMessage = string.IsNullOrWhiteSpace(responseBody)
                ? $"Server returned status code: {response.StatusCode} ({(int)response.StatusCode})"
                : ApiErrorMessageExtractor.ExtractErrorMessage(responseBody);

            LogUtility.LogWarn($"ProcessResponse|Non-success status {(int)response.StatusCode} {response.StatusCode}. Extracted error message: {extractedMessage}");

            return extractedMessage;
        }

        private static string CleanResponse(string response)
        {
            if (string.IsNullOrEmpty(response))
                return response;

            var cleaned = response.Replace("null", string.Empty);

            var sb = new StringBuilder(cleaned.Length + 8);
            bool inString = false;
            bool escape = false;

            for (int i = 0; i < cleaned.Length; i++)
            {
                char c = cleaned[i];

                if (inString)
                {
                    sb.Append(c);

                    if (escape)
                    {
                        escape = false;
                        continue;
                    }

                    if (c == '\\')
                    {
                        escape = true;
                    }
                    else if (c == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (c == '"')
                {
                    inString = true;
                    sb.Append(c);
                    continue;
                }

                if (c == ':')
                {
                    int j = i + 1;
                    while (j < cleaned.Length && char.IsWhiteSpace(cleaned[j]))
                    {
                        j++;
                    }

                    if (j < cleaned.Length && (cleaned[j] == ',' || cleaned[j] == '}' || cleaned[j] == ']'))
                    {
                        sb.Append(':');
                        sb.Append("\"\"");
                        i = j - 1;
                        continue;
                    }
                }

                sb.Append(c);
            }

            return sb.ToString();
        }
    }
}
