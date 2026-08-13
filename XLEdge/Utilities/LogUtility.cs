using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace XLEdge.Utilities
{
    // Per-action debug-log buffering, ported from GLSense's identical LogUtility.cs overhaul (see
    // that project's history for the full design rationale). Debug-mode log lines are buffered per
    // logical action (one top-level LogScope - one ribbon click, one API call, one window's
    // lifecycle) and flushed to disk as one batched write when the outermost scope closes, instead
    // of one file open+write+flush+close cycle per line. Previously, LogDebug wrote to the logger
    // immediately AND separately appended to a global, never-cleared-until-flush "_debugBuffer" -
    // that buffer was purely additive (every line was still written immediately either way) and
    // shared across the whole process rather than scoped to one action, so it never actually
    // avoided the per-line write cost it looked like it was trying to avoid.
    public static class LogUtility
    {
        public static bool DebugMode => XLEdgeAppState.Instance.DebugLogs;

        private static readonly object _lock = new object();

        private sealed class ActionBuffer
        {
            public readonly Guid Id = Guid.NewGuid();
            public readonly List<string> Lines = new List<string>();
            public readonly object Lock = new object();
            public string RootScopeName;
            public int Depth;
            public DateTime OldestUnflushedAtUtc = DateTime.UtcNow;
        }

        // AsyncLocal (not [ThreadStatic]) so the current buffer correctly follows an async method's
        // continuations after a ConfigureAwait(false) resumes on a different thread-pool thread, and
        // so two concurrent-but-unrelated actions never cross-contaminate each other's buffer -
        // forking an async flow copies the *pointer* to the current buffer, not the underlying
        // mutable object, so siblings never see each other's lines.
        private static readonly AsyncLocal<ActionBuffer> _currentBuffer = new AsyncLocal<ActionBuffer>();

        // Every action buffer currently open, keyed by its own id - lets the time-based safety net
        // and FlushAllOpenBuffers (called from shutdown/unhandled-exception hooks) reach buffers that
        // live on a different async flow than whichever thread happens to run them.
        private static readonly ConcurrentDictionary<Guid, ActionBuffer> _openBuffers = new ConcurrentDictionary<Guid, ActionBuffer>();
        private static readonly TimeSpan SafetyNetMaxAge = TimeSpan.FromSeconds(30);
        private static Timer _safetyNetTimer;
        private static readonly object _safetyNetInitLock = new object();

        // Anything logged before AddinModule.Logger is actually initialized (very early startup)
        // would otherwise be silently dropped - held here and flushed once the logger comes online.
        private static readonly List<string> _startupFallbackBuffer = new List<string>();

        private static void EnsureSafetyNetTimerStarted()
        {
            if (_safetyNetTimer != null) return;
            lock (_safetyNetInitLock)
            {
                if (_safetyNetTimer != null) return;
                _safetyNetTimer = new Timer(_ => RunSafetyNetSweep(), null, TimeSpan.FromSeconds(15), TimeSpan.FromSeconds(15));
            }
        }

        private static void RunSafetyNetSweep()
        {
            try
            {
                var now = DateTime.UtcNow;
                foreach (var kvp in _openBuffers)
                {
                    var buffer = kvp.Value;
                    bool stale;
                    lock (buffer.Lock)
                    {
                        stale = buffer.Lines.Count > 0 && (now - buffer.OldestUnflushedAtUtc) > SafetyNetMaxAge;
                    }

                    if (stale)
                    {
                        FlushBuffer(buffer, $"{buffer.RootScopeName} - safety-net flush (still open after {SafetyNetMaxAge.TotalSeconds:F0}s)");
                    }
                }
            }
            catch
            {
                // The safety net must never itself throw - a missed sweep just means the next one
                // (15s later) picks up whatever is still stale.
            }
        }

        internal static object BeginScope(string scopeName)
        {
            if (!DebugMode) return null;
            if (_currentBuffer.Value != null) return null; // nested - not the owner, don't create a second buffer

            var buffer = new ActionBuffer { RootScopeName = scopeName };
            _currentBuffer.Value = buffer;
            _openBuffers[buffer.Id] = buffer;
            EnsureSafetyNetTimerStarted();
            return buffer;
        }

        internal static void IncrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null) buffer.Depth++;
        }

        internal static void DecrementDepth()
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null && buffer.Depth > 0) buffer.Depth--;
        }

        internal static void EndScope(object owned)
        {
            if (owned is ActionBuffer buffer)
            {
                FlushBuffer(buffer, buffer.RootScopeName);
                _openBuffers.TryRemove(buffer.Id, out _);
                _currentBuffer.Value = null;
            }
        }

        private static string Indent()
        {
            int depth = _currentBuffer.Value?.Depth ?? 0;
            return new string(' ', Math.Max(0, depth) * 2);
        }

        private static void FlushBuffer(ActionBuffer buffer, string label)
        {
            string[] lines;
            lock (buffer.Lock)
            {
                if (buffer.Lines.Count == 0) return;
                lines = buffer.Lines.ToArray();
                buffer.Lines.Clear();
            }

            var header = $"===== {label} | {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====";
            var underline = new string('-', header.Length);
            var sb = new StringBuilder();
            sb.AppendLine(header);
            sb.AppendLine(underline);
            foreach (var line in lines)
                sb.AppendLine(line);
            sb.AppendLine(new string('-', underline.Length));

            AddinModule.Logger?.Debug(sb.ToString());
        }

        /// <summary>
        /// Flushes every action buffer currently open, whatever state it's in. Used by the ribbon
        /// Debug toggle (turning debug off mid-action must not silently drop what's already
        /// buffered), and by shutdown/unhandled-exception hooks so nothing buffered is ever lost.
        /// </summary>
        public static void FlushAllOpenBuffers(string reason)
        {
            foreach (var kvp in _openBuffers)
            {
                FlushBuffer(kvp.Value, $"{kvp.Value.RootScopeName} - {reason}");
                _openBuffers.TryRemove(kvp.Key, out _);
            }
        }

        #region Logging Methods
        // Log levels: Warn and Error always write immediately (and force-flush whatever's
        // currently buffered, since a warning/error is exactly the kind of thing that must never be
        // stuck behind a buffer that later gets lost); Debug only writes when DebugMode is enabled,
        // and buffers instead of writing immediately whenever a scope is open.
        public static void LogWarn(string message)
        {
            var logMessage = $"{Indent()}WARN  | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage, NLog.LogLevel.Warn);
            FlushCurrentBuffer("warning logged");
        }
        public static void LogInfo(string message)
        {
            var logMessage = $"{Indent()}INFO  | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage, NLog.LogLevel.Info);
            FlushCurrentBuffer("info logged");
        }
        public static void LogError(string message)
        {
            var logMessage = $"{Indent()}ERROR | {DateTime.Now:HH:mm:ss} | {message}";
            WriteImmediate(logMessage, NLog.LogLevel.Error);
            FlushCurrentBuffer("error logged");
        }

        public static void LogDebug(string message)
        {
            if (!DebugMode) return;

            var logMessage = $"{Indent()}DEBUG | {DateTime.Now:HH:mm:ss} | {message}";

            var buffer = _currentBuffer.Value;
            if (buffer == null)
            {
                WriteImmediate(logMessage, NLog.LogLevel.Debug);
                return;
            }

            lock (buffer.Lock)
            {
                if (buffer.Lines.Count == 0) buffer.OldestUnflushedAtUtc = DateTime.UtcNow;
                buffer.Lines.Add(logMessage);
            }
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

            WriteImmediate(sb.ToString(), NLog.LogLevel.Error);
            FlushCurrentBuffer("exception logged");
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
            WriteImmediate(sb.ToString(), NLog.LogLevel.Error);
            FlushCurrentBuffer("raw JSON logged");
        }
        #endregion

        private static void FlushCurrentBuffer(string reason)
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null)
            {
                FlushBuffer(buffer, $"{buffer.RootScopeName} - {reason}");
            }
        }

        // level controls only the NLog LogLevel the line is recorded under (so LogHelper's
        // "${level:uppercase=true}" layout column shows the right thing) - every level is routed to
        // the same file regardless (see LogHelper.InitializeLogger's AddRule calls), so this never
        // affects whether a line is written, only how it's labeled.
        private static void WriteImmediate(string logMessage, NLog.LogLevel level)
        {
            var logger = AddinModule.Logger;
            if (logger == null)
            {
                lock (_lock)
                {
                    _startupFallbackBuffer.Add(logMessage);
                }
                return;
            }

            if (_startupFallbackBuffer.Count > 0)
            {
                FlushStartupFallbackBuffer(logger);
            }

            logger.Log(level, logMessage);
        }

        private static void FlushStartupFallbackBuffer(NLog.Logger logger)
        {
            List<string> pending;
            lock (_lock)
            {
                if (_startupFallbackBuffer.Count == 0) return;
                pending = new List<string>(_startupFallbackBuffer);
                _startupFallbackBuffer.Clear();
            }

            // These lines already carry their own "WARN |"/"ERROR |"/"DEBUG |" text prefix from
            // whichever LogXxx call originally buffered them, but the level at the time they were
            // buffered (before the logger existed) isn't tracked - Debug is the safe default since
            // AddRule wires every level to the same file/target anyway.
            foreach (var line in pending)
            {
                logger.Debug(line);
            }
        }

        #region Additional Helper Methods (Optional)
        public static void LogMethodEntry([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            LogDebug($"Entering {methodName}");
            IncrementDepth();
        }

        public static void LogMethodExit([System.Runtime.CompilerServices.CallerMemberName] string methodName = "")
        {
            DecrementDepth();
            LogDebug($"Exiting {methodName}");
        }

        public sealed class LogScope : IDisposable
        {
            private readonly string _scopeName;
            private readonly object _ownedBuffer;
            private bool _disposed;

            public LogScope(string scopeName)
            {
                _scopeName = scopeName;
                _ownedBuffer = LogUtility.BeginScope(scopeName);
                LogUtility.LogDebug($"BEGIN: {_scopeName}");
                LogUtility.IncrementDepth();
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            private void Dispose(bool disposing)
            {
                if (_disposed) return;

                if (disposing)
                {
                    LogUtility.DecrementDepth();
                    LogUtility.LogDebug($"END: {_scopeName}");
                    LogUtility.EndScope(_ownedBuffer);
                }

                _disposed = true;
            }
        }
        #endregion
    }
}
