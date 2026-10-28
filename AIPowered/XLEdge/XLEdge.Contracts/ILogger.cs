using System;
using System.Runtime.CompilerServices;

namespace XLEdge.Contracts
{
    /// <summary>
    /// Logger owned by the shell (one NLog instance, one log file). The core logs through a
    /// remoted proxy of this interface, so its implementation must be a MarshalByRefObject.
    /// </summary>
    public interface ILogger
    {
        void LogInfo(string message);
        void LogWarn(string message);
        void LogError(string message, Exception ex = null);
        void LogDebug(string message);
        void LogException(Exception ex, string context = "");
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
