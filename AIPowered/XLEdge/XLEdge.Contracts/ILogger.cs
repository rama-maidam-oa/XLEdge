using System;
using System.Runtime.CompilerServices;

namespace XLEdge.Contracts
{
    /// <summary>
    /// Logger owned by the shell (one NLog instance, one log file). The core logs through a
    /// remoted proxy of this interface, so its implementation must be a MarshalByRefObject.
    /// Exceptions are passed as text (exception.ToString()): an Exception object would cross the
    /// AppDomain boundary by value, and core-only exception types (API, WebView2) can neither be
    /// serialized nor resolved in the shell's AppDomain.
    /// </summary>
    public interface ILogger
    {
        void LogInfo(string message);
        void LogWarn(string message);
        void LogError(string message, string exceptionDetails = null);
        void LogDebug(string message);
        void LogException(string exceptionDetails, string context = "");
        void LogRawJson(string context, string rawJson);
        void FlushDebugLogs(string section = "Buffered Logs");

        /// <summary>Marks the start of a loaded core release in the log (startup and every reload).</summary>
        void BeginReleaseSession(string version, string releaseDate);

        /// <summary>Marks the end of a loaded core release in the log (shutdown and before a reload).</summary>
        void EndReleaseSession(string version, string releaseDate);

        void LogMethodEntry([CallerMemberName] string methodName = "");
        void LogMethodExit([CallerMemberName] string methodName = "");

        /// <summary>
        /// Opens a buffered log scope for one user action; disposing it writes everything logged
        /// inside in one batch. Returns IDisposable so the core never needs the concrete scope type.
        /// </summary>
        IDisposable BeginLogScope(string scopeName);
    }
}
