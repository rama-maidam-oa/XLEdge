using System;
using System.Text;
using XLEdge.Addin.Core.Infrastructure;
using XLEdge.Contracts;

namespace XLEdge.Utilities
{
    /// <summary>
    /// Today's XLEdge logging facade (same members), for code ported into the core. The shell owns the log: its Logger
    /// (reached through the context) writes the file, buffers debug lines per action scope and flushes them, so this class
    /// only formats what today's LogUtility formatted itself (the exception body, the raw-JSON preview) and forwards. Before
    /// Initialize and after Shutdown there is no logger: calls are dropped, never thrown.
    /// </summary>
    public static class LogUtility
    {
        private const int RawJsonPreviewLength = 200;

        public static bool DebugMode => XLEdgeAppState.Instance.DebugLogs;

        /// <summary>Flushes every open action buffer (Debug toggled off, shutdown, unhandled exception).</summary>
        public static void FlushAllOpenBuffers(string reason) => Write(l => l.FlushDebugLogs(reason));

        public static void LogWarn(string message) => Write(l => l.LogWarn(message));

        public static void LogInfo(string message) => Write(l => l.LogInfo(message));

        public static void LogError(string message) => Write(l => l.LogError(message, null));

        public static void LogDebug(string message)
        {
            // Today's check, so a disabled debug line is not formatted or sent (DebugMode itself is one call to the shell).
            if (!DebugMode) return;
            Write(l => l.LogDebug(message));
        }

        /// <summary>Today's exception block; the shell's logger adds the frame, the context line and the indentation.</summary>
        public static void LogException(Exception ex, string context = "")
        {
            if (ex == null) return;
            var sb = new StringBuilder();
            sb.AppendLine($"Type: {ex.GetType().FullName}");
            sb.AppendLine($"Message: {ex.Message}");
            sb.AppendLine($"Source: {ex.Source}");
            sb.AppendLine($"TargetSite: {ex.TargetSite}");
            sb.AppendLine("StackTrace:");
            foreach (var line in ex.StackTrace?.Split(new[] { Environment.NewLine }, StringSplitOptions.None) ?? new string[0])
                sb.AppendLine(line);
            if (ex.InnerException != null)
            {
                sb.AppendLine("----- Inner Exception -----");
                sb.AppendLine($"Type: {ex.InnerException.GetType().FullName}");
                sb.AppendLine($"Message: {ex.InnerException.Message}");
                sb.Append(ex.InnerException.StackTrace);
            }
            string details = sb.ToString().TrimEnd();
            Write(l => l.LogException(details, context));
        }

        /// <summary>
        /// A raw JSON payload (e.g. on a parse failure). With Debug mode on, the full payload - in the log when small, in a
        /// XLEdge_Logs\Temp file (preview + path in the log) when over 100 KB (LargeLogText, ported from GLSense OISR-22830;
        /// replaces the "Include Output Data" option). With Debug off, its length and a short preview.
        /// </summary>
        public static void LogRawJson(string context, string rawJson)
        {
            string text;
            if (string.IsNullOrEmpty(rawJson))
            {
                text = "<empty>";
            }
            else if (DebugMode)
            {
                text = LargeLogText.ToLogText(rawJson, XLEdgeAppPaths.TempFolder, "raw-json_" + context);
            }
            else
            {
                string preview = rawJson.Substring(0, Math.Min(RawJsonPreviewLength, rawJson.Length));
                text = $"<{rawJson.Length} character(s) - enable Debug mode to log the full payload> Preview: {preview}";
            }
            Write(l => l.LogRawJson(context ?? string.Empty, text));
        }

        /// <summary>
        /// A response body for a log line: with Debug mode on, the full text when small or a preview + the XLEdge_Logs\Temp
        /// file holding it (see LargeLogText); with Debug off, only its length.
        /// </summary>
        public static string BodyForLog(string body, string nameHint)
        {
            if (string.IsNullOrEmpty(body)) return "<empty>";
            return DebugMode
                ? LargeLogText.ToLogText(body, XLEdgeAppPaths.TempFolder, nameHint)
                : $"<{body.Length} character(s) - enable Debug mode to log the full body>";
        }

        public static void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "") =>
            LogDebug($"Entering {methodName}");

        public static void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "") =>
            LogDebug($"Exiting {methodName}");

        /// <summary>One action's log scope (today's type): the shell buffers its debug lines and writes them as one block.</summary>
        public sealed class LogScope : IDisposable
        {
            private IDisposable _scope;

            public LogScope(string scopeName)
            {
                try { _scope = Logger?.BeginLogScope(scopeName); }
                catch (Exception) { _scope = null; }
            }

            public void Dispose()
            {
                IDisposable scope = _scope;
                _scope = null;
                try { scope?.Dispose(); }
                catch (Exception) { /* the shell's scope may be gone with a closing session */ }
            }
        }

        private static ILogger Logger
        {
            get
            {
                try { return CoreServices.Context?.Logger; }
                catch (Exception) { return null; }
            }
        }

        private static void Write(Action<ILogger> write)
        {
            try
            {
                ILogger logger = Logger;
                if (logger != null) write(logger);
            }
            catch (Exception)
            {
                // Logging must never break the caller (and nothing may escape the core).
            }
        }
    }
}
