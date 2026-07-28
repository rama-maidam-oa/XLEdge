using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Threading;
using XLEdge.Utilities;
using XLEdge.Views;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Centralized helper for API operations with retry logic and better error handling
    /// </summary>
    public static class ApiOperationHelper
    {
        private const int MaxRetries = 3;
        private const int RetryDelayMs = 1000;

        /// <summary>
        /// Executes an API call with retry logic for transient failures
        /// </summary>
        public static async Task<string> ExecuteWithRetry(
            Func<CancellationToken, Task<string>> apiCall,
            CancellationToken cancellationToken,
            string operationName = "API call")
        {
            using (new LogUtility.LogScope($"ExecuteWithRetry: {operationName}"))
            {
                int retryCount = 0;
                Exception lastException = null;

                while (retryCount < MaxRetries)
                {
                    try
                    {
                        CancellationTokenHelper.ThrowIfCancelled(cancellationToken, operationName);

                        if (retryCount > 0)
                        {
                            LogUtility.LogDebug($"Retry attempt {retryCount} of {MaxRetries - 1}");
                        }

                        var result = await apiCall(cancellationToken);

                        if (retryCount > 0)
                        {
                            // Logged as a warning (always visible) since a successful retry is a
                            // reliability signal worth noticing, even though it isn't fatal.
                            LogUtility.LogWarn($"{operationName} succeeded after {retryCount} retries - can be ignored if intermittent, but investigate if this recurs often.");
                        }

                        return result;
                    }
                    catch (OperationCanceledException)
                    {
                        LogUtility.LogWarn($"{operationName} was cancelled");
                        throw;
                    }
                    catch (Exception ex) when (IsTransientError(ex))
                    {
                        lastException = ex;
                        retryCount++;

                        if (retryCount >= MaxRetries)
                        {
                            LogUtility.LogError($"{operationName} failed after {MaxRetries} attempts");
                            ExceptionHelper.LogDetailedException(ex, $"{operationName} - Max retries exceeded");
                            throw;
                        }

                        LogUtility.LogWarn($"{operationName} failed (transient error) - retrying in {RetryDelayMs}ms: {ex.Message}");

                        await CancellationTokenHelper.DelayWithLogging(
                            RetryDelayMs * retryCount, // Exponential backoff
                            cancellationToken,
                            $"Retry delay for {operationName}");
                    }
                    catch (Exception ex)
                    {
                        // Non-transient error - don't retry
                        LogUtility.LogError($"{operationName} failed with non-transient error");
                        ExceptionHelper.LogDetailedException(ex, operationName);
                        throw;
                    }
                }

                // Should never reach here, but for safety
                if (lastException != null)
                {
                    throw lastException;
                }

                return string.Empty;
            }
        }

        /// <summary>
        /// Determines if an error is transient and worth retrying
        /// </summary>
        private static bool IsTransientError(Exception ex)
        {
            return ex is System.Net.Http.HttpRequestException ||
                   ex is TimeoutException ||
                   ex is System.Net.WebException ||
                   (ex is System.Net.Sockets.SocketException socketEx &&
                    IsTransientSocketError(socketEx));
        }

        private static bool IsTransientSocketError(System.Net.Sockets.SocketException socketEx)
        {
            // Connection timeout, connection refused, network unreachable, etc.
            return socketEx.SocketErrorCode == System.Net.Sockets.SocketError.TimedOut ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.ConnectionRefused ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.NetworkUnreachable ||
                   socketEx.SocketErrorCode == System.Net.Sockets.SocketError.HostDown;
        }

        /// <summary>
        /// Validates API response for common error patterns
        /// </summary>
        public static bool ValidateApiResponse(string response, out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                if (string.IsNullOrWhiteSpace(response))
                {
                    errorMessage = "Empty response from server";
                    LogUtility.LogWarn(errorMessage);
                    return false;
                }

                if (response.IndexOf("(401) Unauthorized", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    errorMessage = "Session expired! Please re-login.";
                    LogUtility.LogError(errorMessage);
                    return false;
                }

                if (response.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
                {
                    errorMessage = response;

                    // Only log the full response body when Debug mode + "Include Output Data" are
                    // both on, since it can be very large; otherwise log a short length-only summary.
                    if (LogUtility.DebugMode && XLEdgeAppState.Instance.DebugOutputData)
                    {
                        LogUtility.LogError($"API returned error: {response}");
                    }
                    else
                    {
                        LogUtility.LogError($"API returned error ({response.Length} character(s) - enable Debug mode + 'Include Output Data' to log the full response).");
                    }

                    return false;
                }

                if (response.IndexOf("<!DOCTYPE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    errorMessage = "Received HTML response instead of expected data";
                    LogUtility.LogError(errorMessage);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHelper.LogDetailedException(ex, "ValidateApiResponse");
                errorMessage = "Error validating API response";
                return false;
            }
        }

        /// <summary>
        /// Parses JSON response with detailed error logging
        /// </summary>
        public static T ParseJsonResponse<T>(string jsonResponse, string operationName = "")
        {
            using (new LogUtility.LogScope($"ParseJsonResponse<{typeof(T).Name}>: {operationName}"))
            {
                try
                {
                    LogUtility.LogDebug($"Parsing JSON response (Length: {jsonResponse?.Length ?? 0})");

                    var result = JsonSerializer.Deserialize<T>(jsonResponse, JsonGlobals.Options);

                    return result;
                }
                catch (JsonException ex)
                {
                    LogUtility.LogRawJson($"ParseJsonResponse<{typeof(T).Name}>", jsonResponse ?? string.Empty);
                    ExceptionHelper.LogDetailedException(ex, $"ParseJsonResponse: {operationName}");
                    throw;
                }
            }
        }

        /// <summary>
        /// Displays an API error message either on the current overlay or via GLMessageWindow and returns control.
        /// </summary>
        public static void NotifyApiError(string errorMessage, Dispatcher dispatcher = null, AppOverlay overlay = null, string operationName = null)
        {
            var message = string.IsNullOrWhiteSpace(errorMessage)
                ? "No response received from the server."
                : errorMessage.Trim();

            if (!string.IsNullOrWhiteSpace(operationName))
            {
                message = $"{operationName}: {message}";
            }

            if (dispatcher != null && overlay != null)
            {
                try
                {
                    dispatcher.InvokeAsync(() => overlay.ShowError(message), DispatcherPriority.Background);
                    return;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to display API error on overlay");
                }
            }

            MessageFunctions.XLEdgeMessage(message, MessageBoxIcon.Error, MessageBoxButtons.OK);
        }

        /// <summary>
        /// Validates the API response and throws an <see cref="InvalidOperationException"/> carrying
        /// the validation error message when invalid, leaving display of the error to the caller.
        /// </summary>
        public static void EnsureValidApiResponse(string response, string operationName, Dispatcher dispatcher = null, AppOverlay overlay = null)
        {
            if (ValidateApiResponse(response, out var errorMessage))
                return;

            throw new InvalidOperationException(string.IsNullOrWhiteSpace(errorMessage)
                ? "Invalid API response."
                : errorMessage);
        }
    }
}
