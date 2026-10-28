using System;
using System.Collections.Generic;
using System.Linq;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>Test-only ILogger that keeps every line; MarshalByRefObject so a second AppDomain can log to it.</summary>
    public sealed class RecordingLogger : MarshalByRefObject, ILogger
    {
        private readonly List<string> _lines = new List<string>();

        public string[] Lines { get { lock (_lines) return _lines.ToArray(); } }

        public bool Contains(string text) => Lines.Any(l => l.IndexOf(text, StringComparison.Ordinal) >= 0);

        public string Dump() => string.Join(Environment.NewLine, Lines);

        private readonly List<string> _events = new List<string>();

        /// <summary>Flushes and scopes, kept apart from the log lines so tests that compare lines are unaffected.</summary>
        public string[] Events { get { lock (_events) return _events.ToArray(); } }

        private void Event(string text) { lock (_events) _events.Add(text); }

        private void Add(string level, string message) { lock (_lines) _lines.Add(level + " | " + message); }

        public void LogInfo(string message) => Add("INFO", message);
        public void LogWarn(string message) => Add("WARN", message);
        public void LogError(string message, string exceptionDetails = null) => Add("ERROR", message + (exceptionDetails == null ? "" : "|" + exceptionDetails));
        public void LogDebug(string message) => Add("DEBUG", message);
        public void LogException(string exceptionDetails, string context = "") => Add("EXCEPTION", context + "|" + exceptionDetails);
        public void LogRawJson(string context, string rawJson) => Add("JSON", context + "|" + rawJson);
        public void FlushDebugLogs(string section = "Buffered Logs") => Event("flush:" + section);
        public void BeginReleaseSession(string version, string releaseDate) => Add("SESSION", "begin " + version);
        public void EndReleaseSession(string version, string releaseDate) => Add("SESSION", "end " + version);
        public IDisposable BeginLogScope(string scopeName)
        {
            Event("begin:" + scopeName);
            return new Scope(() => Event("end:" + scopeName));
        }

        public override object InitializeLifetimeService() => null;

        private sealed class Scope : MarshalByRefObject, IDisposable
        {
            private readonly Action _onDispose;
            public Scope(Action onDispose = null) { _onDispose = onDispose; }
            public void Dispose() => _onDispose?.Invoke();
            public override object InitializeLifetimeService() => null;
        }
    }
}
