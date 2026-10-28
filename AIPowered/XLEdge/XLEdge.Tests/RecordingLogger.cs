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

        private void Add(string level, string message) { lock (_lines) _lines.Add(level + " | " + message); }

        public void LogInfo(string message) => Add("INFO", message);
        public void LogWarn(string message) => Add("WARN", message);
        public void LogError(string message, string exceptionDetails = null) => Add("ERROR", message + (exceptionDetails == null ? "" : "|" + exceptionDetails));
        public void LogDebug(string message) => Add("DEBUG", message);
        public void LogException(string exceptionDetails, string context = "") => Add("EXCEPTION", context + "|" + exceptionDetails);
        public void LogRawJson(string context, string rawJson) => Add("JSON", context);
        public void FlushDebugLogs(string section = "Buffered Logs") { }
        public void BeginReleaseSession(string version, string releaseDate) => Add("SESSION", "begin " + version);
        public void EndReleaseSession(string version, string releaseDate) => Add("SESSION", "end " + version);
        public IDisposable BeginLogScope(string scopeName) => new Scope();

        public override object InitializeLifetimeService() => null;

        private sealed class Scope : MarshalByRefObject, IDisposable
        {
            public void Dispose() { }
            public override object InitializeLifetimeService() => null;
        }
    }
}
