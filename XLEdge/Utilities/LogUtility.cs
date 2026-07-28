using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Text;
using System.Threading.Tasks;

namespace XLEdge.Utilities
{
    public static class LogUtility
    {
        // Buffer for debug logs (used when DebugMode is true)
        private static readonly List<string> _debugBuffer = new List<string>();

        // Toggle debug buffering
        public static bool DebugMode => XLEdgeAppState.Instance.DebugLogs;

        // Thread-local scope depth for nested indentation
        [ThreadStatic]
        private static int _scopeDepth;

        internal static void IncrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth) + 1;
        }

        internal static void DecrementScope()
        {
            _scopeDepth = Math.Max(0, _scopeDepth - 1);
        }

        private static string Indent()
        {
            // Ensure _scopeDepth is never negative
            int safeDepth = Math.Max(0, _scopeDepth);
            return new string(' ', safeDepth * 2);
        }

        #region Logging Methods
        // Log levels: Warn and Error always write; Debug only writes when DebugMode is enabled.
        public static void LogWarn(string message)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {message}";
            AddinModule.Logger?.Warn(logMessage);
        }

        public static void LogError(string message)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {message}";
            AddinModule.Logger?.Error(logMessage);
        }

        public static void LogDebug(string message)
        {
            // ONLY log if debug mode is enabled
            if (!DebugMode)
                return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {message}";

            // Write to logger
            AddinModule.Logger?.Debug(logMessage);

            // Also buffer for flushing
            _debugBuffer.Add(logMessage);
        }

        public static void LogException(Exception ex, string context = "")
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}========== Exception ==========");
            if (!string.IsNullOrEmpty(context))
                sb.AppendLine($"{Indent()}Context: {context}");
            sb.AppendLine($"{Indent()}Type: {ex.GetType().FullName}");
            sb.AppendLine($"{Indent()}Message: {ex.Message}");
            sb.AppendLine($"{Indent()}Source: {ex.Source}");
            sb.AppendLine($"{Indent()}TargetSite: {ex.TargetSite}");
            sb.AppendLine($"{Indent()}StackTrace:");
            foreach (var line in ex.StackTrace?.Split(new[] { Environment.NewLine }, StringSplitOptions.None) ?? [])
                sb.AppendLine($"{Indent()}{line}");
            if (ex.InnerException != null)
            {
                sb.AppendLine($"{Indent()}----- Inner Exception -----");
                sb.AppendLine($"{Indent()}Type: {ex.InnerException.GetType().FullName}");
                sb.AppendLine($"{Indent()}Message: {ex.InnerException.Message}");
                sb.AppendLine($"{Indent()}{ex.InnerException.StackTrace}");
            }
            sb.AppendLine($"{Indent()}============================");

            var exceptionMessage = sb.ToString();
            AddinModule.Logger?.Error(exceptionMessage);
        }

        // Logs a raw JSON payload (e.g. on a parse failure) at Error level. The full payload is
        // only included when both DebugMode and "Include Output Data" are enabled; otherwise only
        // a short length + preview snippet is logged.
        public static void LogRawJson(string context, string rawJson)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"{Indent()}----- Raw JSON {(string.IsNullOrWhiteSpace(context) ? string.Empty : "(" + context + ")")} -----");

            if (string.IsNullOrEmpty(rawJson))
            {
                sb.AppendLine("<empty>");
            }
            else if (DebugMode && XLEdgeAppState.Instance.DebugOutputData)
            {
                sb.AppendLine(rawJson);
            }
            else
            {
                string preview = rawJson.Substring(0, Math.Min(200, rawJson.Length));
                sb.AppendLine($"<{rawJson.Length} character(s) - enable Debug mode + 'Include Output Data' to log the full payload> Preview: {preview}");
            }

            sb.AppendLine($"{Indent()}----- End Raw JSON -----");
            AddinModule.Logger?.Error(sb.ToString());
        }
        #endregion

        #region Flush
        public static void FlushDebugLogs(string section = "Buffered Logs")
        {
            if (_debugBuffer.Count == 0) return;

            var header = $"===== {section} | {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====";
            var underline = new string('-', header.Length);
            var sb = new StringBuilder();
            sb.AppendLine(header);
            sb.AppendLine(underline);
            foreach (var line in _debugBuffer)
                sb.AppendLine(line);
            sb.AppendLine(new string('-', underline.Length));
            sb.AppendLine();

            AddinModule.Logger?.Debug(sb.ToString());
            _debugBuffer.Clear();
        }
        #endregion

        #region Additional Helper Methods (Optional)
        public static void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
            IncrementScope();
        }

        public static void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            DecrementScope();
            LogDebug($"Exiting {methodName}");
        }

        // Helper class for scope-based logging
        public class LogScope : IDisposable
        {
            private readonly string _scopeName;
            private bool _disposed = false;
            public LogScope(string scopeName)
            {
                _scopeName = scopeName;
                LogUtility.LogDebug($"BEGIN: {_scopeName}");
                LogUtility.IncrementScope();
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!_disposed)
                {
                    if (disposing)
                    {
                        // Dispose managed resources here
                        LogUtility.DecrementScope();
                        LogUtility.LogDebug($"END: {_scopeName}");
                    }

                    // Dispose unmanaged resources here (none in this case)
                    _disposed = true;
                }
            }
        }
        #endregion
    }
}
