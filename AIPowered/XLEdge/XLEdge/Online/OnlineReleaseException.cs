using System;

namespace XLEdge.Online
{
    /// <summary>A failed call to the server's XLEdge release endpoints, with the message the user is shown.</summary>
    [Serializable]
    internal sealed class OnlineReleaseException : Exception
    {
        public OnlineReleaseException(string message, int? statusCode) : base(message) { StatusCode = statusCode; }

        /// <summary>The HTTP status the server answered with; null when it could not be reached.</summary>
        public int? StatusCode { get; }
    }
}
