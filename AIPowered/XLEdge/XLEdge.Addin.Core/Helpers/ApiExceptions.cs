using System;
using System.Net;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Ported from FormProcessBar.vb's ApiRequestException - a definitive, non-success HTTP response
    /// from the server (status code + the readable message extracted by ApiErrorMessageExtractor).
    /// Deliberately NOT one of ApiOperationHelper.IsTransientError's recognized types, so a call that
    /// fails this way is not blindly retried the way a genuinely transient network error is.
    /// </summary>
    public class ApiRequestException : Exception
    {
        public HttpStatusCode StatusCode { get; }

        public ApiRequestException(string message, HttpStatusCode statusCode)
            : base(message)
        {
            StatusCode = statusCode;
        }

        public ApiRequestException(string message, HttpStatusCode statusCode, Exception innerException)
            : base(message, innerException)
        {
            StatusCode = statusCode;
        }
    }

    /// <summary>
    /// Ported from FormProcessBar.vb's ApiTimeoutException - thrown when an API call's own request
    /// timeout expires (as opposed to a caller/user-driven cancellation, which still surfaces as a
    /// plain OperationCanceledException). Inherits TimeoutException so
    /// ApiOperationHelper.IsTransientError's existing "retry on TimeoutException" branch already
    /// covers it without any further change there.
    /// </summary>
    public class ApiTimeoutException : TimeoutException
    {
        public ApiTimeoutException(string message)
            : base(message)
        {
        }

        public ApiTimeoutException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
