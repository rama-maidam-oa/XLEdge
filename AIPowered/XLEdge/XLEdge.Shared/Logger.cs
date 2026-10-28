using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using NLog;
using NLog.Config;
using NLog.Targets;
using XLEdge.Contracts;
using ILogger = XLEdge.Contracts.ILogger;

namespace XLEdge.Shared
{
    /// <summary>
    /// The XLEdge log: one file per day, XLEdge_Logs_dd-MMM-yyyy.log under Paths.Logs, written by one NLog
    /// instance owned by the shell. The core logs through a remoting proxy of this object.
    /// <para>
    /// Debug lines are written only in debug mode. Inside a scope (BeginLogScope) they are buffered and written
    /// as one block when the scope ends, when a warning, error or exception is logged, or by the safety net
    /// when the oldest buffered line is more than 30 seconds old.
    /// </para>
    /// </summary>
    public sealed class Logger : MarshalByRefObject, ILogger, IDisposable
    {
        private const string LogFilePrefix = "XLEdge_Logs_";
        private const string StartMarkerPrefix = "===== Start of XLEdge version ";
        private const string EndMarkerPrefix = "===== End of XLEdge version ";
        private static readonly CultureInfo English = CultureInfo.GetCultureInfo("en-US");
        private static readonly TimeSpan SafetyNetMaxAge = TimeSpan.FromSeconds(30);
        private static readonly TimeSpan SafetyNetPeriod = TimeSpan.FromSeconds(15);

        private readonly IXLEdgeContext _context;
        private readonly string _logsFolder;
        private readonly LogFactory _factory;
        private readonly NLog.Logger _nlog;
        private readonly AsyncLocal<ActionBuffer> _currentBuffer = new AsyncLocal<ActionBuffer>();
        private readonly ConcurrentDictionary<Guid, ActionBuffer> _openBuffers = new ConcurrentDictionary<Guid, ActionBuffer>();
        private readonly object _timerLock = new object();
        private Timer _safetyNetTimer;
        private bool _disposed;

        private sealed class ActionBuffer
        {
            public readonly Guid Id = Guid.NewGuid();
            public readonly List<string> Lines = new List<string>();
            public readonly object Lock = new object();
            public string RootScopeName;
            public int Depth;
            public DateTime? OldestUnflushedAtUtc;
        }

        public Logger(IXLEdgeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logsFolder = context.Paths.Logs;
            try
            {
                Directory.CreateDirectory(_logsFolder);
                _factory = new LogFactory();
                var config = new LoggingConfiguration(_factory);
                var file = new FileTarget("logfile")
                {
                    FileName = Path.Combine(_logsFolder, LogFilePrefix + "${date:format=dd-MMM-yyyy:culture=en-US}.log"),
                    Header = EscapeLayout(BuildEnvironmentSnapshot()),
                    Layout = "${longdate}|${message:withException=true:exceptionSeparator=|}",
                    AutoFlush = true,
                    KeepFileOpen = true,
                    DeleteOldFileOnStartup = false,
                    ArchiveAboveSize = 20 * 1024 * 1024,
                    MaxArchiveFiles = 30,
                    // No ArchiveFileName: with a date-based FileName it switches NLog to the legacy archive
                    // handler, which fights KeepFileOpen. This gives XLEdge_Logs_28-Oct-2026(1).log, (2)...
                    ArchiveSuffixFormat = "({0})",
                };
                config.AddRule(LogLevel.Debug, LogLevel.Fatal, file);
                _factory.Configuration = config;
                _nlog = _factory.GetLogger("XLEdge");
            }
            catch (Exception ex)
            {
                _nlog = null;
                WriteInitFailure(ex);
            }
        }

        // ---- ILogger ----------------------------------------------------------------

        public void LogInfo(string message) => WriteImmediate(Format("INFO  ", message));

        public void LogWarn(string message)
        {
            WriteImmediate(Format("WARN  ", message));
            FlushCurrent("warning logged");
        }

        public void LogError(string message, string exceptionDetails = null)
        {
            string line = Format("ERROR ", message);
            if (!string.IsNullOrEmpty(exceptionDetails)) line += "|" + exceptionDetails;
            WriteImmediate(line);
            FlushCurrent("error logged");
        }

        public void LogDebug(string message)
        {
            if (!DebugMode) return;
            string line = Format("DEBUG ", message);
            var buffer = _currentBuffer.Value;
            if (buffer == null)
            {
                WriteImmediate(line);
                return;
            }
            lock (buffer.Lock)
            {
                buffer.Lines.Add(line);
                if (buffer.OldestUnflushedAtUtc == null) buffer.OldestUnflushedAtUtc = DateTime.UtcNow;
            }
        }

        public void LogException(string exceptionDetails, string context = "")
        {
            string indent = Indent();
            var block = new StringBuilder();
            block.Append(indent).Append("========== Exception ==========");
            if (!string.IsNullOrWhiteSpace(context)) block.AppendLine().Append(indent).Append("Context: ").Append(context);
            foreach (string line in SplitLines(string.IsNullOrEmpty(exceptionDetails) ? "<no details>" : exceptionDetails))
                block.AppendLine().Append(indent).Append(line);
            block.AppendLine().Append(indent).Append("============================");
            WriteImmediate(block.ToString());
            FlushCurrent("exception logged");
        }

        public void LogRawJson(string context, string rawJson)
        {
            string indent = Indent();
            WriteImmediate($"{indent}----- Raw JSON ({context}) -----{Environment.NewLine}{(string.IsNullOrEmpty(rawJson) ? "<empty>" : rawJson)}{Environment.NewLine}{indent}----- End Raw JSON -----");
        }

        public void FlushDebugLogs(string section = "Buffered Logs")
        {
            foreach (var buffer in _openBuffers.Values)
                FlushBuffer(buffer, $"{buffer.RootScopeName} - {section}");
        }

        /// <summary>
        /// Writes the start marker of a core release section, unless the latest section in today's log is already
        /// the start of this same release (e.g. Excel restarted on the same release).
        /// </summary>
        public void BeginReleaseSession(string version, string releaseDate)
        {
            string marker = StartMarkerPrefix + FormatRelease(version, releaseDate) + " logs";
            if (LatestSectionMarkerStartsWith(marker)) return;
            WriteImmediate($"{marker} as on {DateTime.Now.ToString("dddd, dd MMMM yyyy", English)} =====");
        }

        public void EndReleaseSession(string version, string releaseDate)
        {
            FlushDebugLogs("release session ending");
            WriteImmediate(EndMarkerPrefix + FormatRelease(version, releaseDate) + " logs =====" + Environment.NewLine);
        }

        public IDisposable BeginLogScope(string scopeName) => new LogScope(this, scopeName);

        public override object InitializeLifetimeService() => null;

        /// <summary>Flushes every open buffer and closes the log file. Called by the shell when the add-in shuts down.</summary>
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            lock (_timerLock)
            {
                _safetyNetTimer?.Dispose();
                _safetyNetTimer = null;
            }
            FlushDebugLogs("add-in shutting down");
            _factory?.Shutdown();
        }

        /// <summary>A per-action log scope. MarshalByRefObject so the core can dispose it through the proxy.</summary>
        public sealed class LogScope : MarshalByRefObject, IDisposable
        {
            private readonly Logger _logger;
            private readonly string _name;
            private readonly ActionBuffer _owned;
            private bool _disposed;

            internal LogScope(Logger logger, string name)
            {
                _logger = logger;
                _name = name;
                _owned = logger.BeginScope(name);
                logger.LogDebug("BEGIN: " + name);
                logger.ChangeDepth(+1);
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;
                _logger.ChangeDepth(-1);
                _logger.LogDebug("END: " + _name);
                _logger.EndScope(_owned);
            }

            public override object InitializeLifetimeService() => null;
        }

        // ---- buffering --------------------------------------------------------------

        private bool DebugMode
        {
            get
            {
                try { return _context.DebugMode; }
                catch (Exception) { return false; }
            }
        }

        private ActionBuffer BeginScope(string name)
        {
            if (!DebugMode || _currentBuffer.Value != null) return null;
            var buffer = new ActionBuffer { RootScopeName = name };
            _currentBuffer.Value = buffer;
            _openBuffers[buffer.Id] = buffer;
            EnsureSafetyNetTimer();
            return buffer;
        }

        private void EndScope(ActionBuffer owned)
        {
            if (owned == null) return;
            FlushBuffer(owned, owned.RootScopeName);
            _openBuffers.TryRemove(owned.Id, out _);
            if (_currentBuffer.Value == owned) _currentBuffer.Value = null;
        }

        private void ChangeDepth(int delta)
        {
            var buffer = _currentBuffer.Value;
            if (buffer == null) return;
            lock (buffer.Lock) buffer.Depth = Math.Max(0, buffer.Depth + delta);
        }

        private string Indent()
        {
            var buffer = _currentBuffer.Value;
            if (buffer == null) return string.Empty;
            lock (buffer.Lock) return new string(' ', 2 * buffer.Depth);
        }

        private void FlushCurrent(string reason)
        {
            var buffer = _currentBuffer.Value;
            if (buffer != null) FlushBuffer(buffer, $"{buffer.RootScopeName} - {reason}");
        }

        private void FlushBuffer(ActionBuffer buffer, string label)
        {
            string[] lines;
            lock (buffer.Lock)
            {
                if (buffer.Lines.Count == 0) return;
                lines = buffer.Lines.ToArray();
                buffer.Lines.Clear();
                buffer.OldestUnflushedAtUtc = null;
            }
            string header = $"===== {label} | {DateTime.Now:yyyy-MM-dd HH:mm:ss} =====";
            string dashes = new string('-', header.Length);
            string nl = Environment.NewLine;
            WriteImmediate(header + nl + dashes + nl + string.Join(nl, lines) + nl + dashes);
        }

        private void EnsureSafetyNetTimer()
        {
            if (_safetyNetTimer != null) return;
            lock (_timerLock)
            {
                if (_safetyNetTimer == null && !_disposed)
                    _safetyNetTimer = new Timer(_ => SweepSafetyNet(DateTime.UtcNow), null, SafetyNetPeriod, SafetyNetPeriod);
            }
        }

        /// <summary>Flushes buffers whose oldest line is older than 30 seconds (long-running actions keep buffering afterwards).</summary>
        internal void SweepSafetyNet(DateTime nowUtc)
        {
            try
            {
                foreach (var buffer in _openBuffers.Values)
                {
                    DateTime? oldest;
                    lock (buffer.Lock) oldest = buffer.OldestUnflushedAtUtc;
                    if (oldest != null && nowUtc - oldest.Value >= SafetyNetMaxAge)
                        FlushBuffer(buffer, $"{buffer.RootScopeName} - safety-net flush, action still running");
                }
            }
            catch (Exception)
            {
                // Never let the timer thread throw.
            }
        }

        // ---- writing ----------------------------------------------------------------

        private string Format(string level, string message) => $"{Indent()}{level}| {DateTime.Now:HH:mm:ss} | {message}";

        private void WriteImmediate(string text)
        {
            if (_nlog == null || _disposed) return;
            _nlog.Debug(text);
        }

        private static string FormatRelease(string version, string releaseDate) =>
            $"{(string.IsNullOrWhiteSpace(version) ? "unknown" : version)} (released {(string.IsNullOrWhiteSpace(releaseDate) ? "unknown release date" : releaseDate)})";

        private static IEnumerable<string> SplitLines(string text) => text.Replace("\r\n", "\n").Split('\n');

        // NLog parses "${" in a literal layout; the snapshot is plain text.
        private static string EscapeLayout(string text) => text.Replace("${", "$ {");

        /// <summary>True when the last release marker in today's newest log file starts with <paramref name="marker"/>.</summary>
        private bool LatestSectionMarkerStartsWith(string marker)
        {
            try
            {
                string pattern = LogFilePrefix + DateTime.Now.ToString("dd-MMM-yyyy", English) + "*.log";
                var newest = new DirectoryInfo(_logsFolder).GetFiles(pattern).OrderByDescending(f => f.LastWriteTimeUtc).FirstOrDefault();
                if (newest == null) return false;

                foreach (string line in ReadLinesBackwards(newest.FullName))
                {
                    int bar = line.IndexOf('|');
                    string message = bar >= 0 ? line.Substring(bar + 1) : line;
                    if (message.StartsWith(StartMarkerPrefix, StringComparison.Ordinal) || message.StartsWith(EndMarkerPrefix, StringComparison.Ordinal))
                        return message.StartsWith(marker, StringComparison.Ordinal);
                }
            }
            catch (Exception)
            {
            }
            return false;
        }

        /// <summary>Lines of a (possibly 20 MB) log file from the last to the first, read in 64 KB chunks.</summary>
        private static IEnumerable<string> ReadLinesBackwards(string path)
        {
            const int chunkSize = 64 * 1024;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                long position = stream.Length;
                var carry = new List<byte>();
                var chunk = new byte[chunkSize];
                while (position > 0)
                {
                    int count = (int)Math.Min(chunkSize, position);
                    position -= count;
                    stream.Position = position;
                    int read = 0;
                    while (read < count) read += stream.Read(chunk, read, count - read);

                    var bytes = new byte[count + carry.Count];
                    Array.Copy(chunk, bytes, count);
                    carry.CopyTo(bytes, count);

                    int end = bytes.Length;
                    for (int i = bytes.Length - 1; i >= 0; i--)
                    {
                        if (bytes[i] != (byte)'\n') continue;
                        yield return Encoding.UTF8.GetString(bytes, i + 1, end - i - 1).TrimEnd('\r');
                        end = i;
                    }
                    carry = bytes.Take(end).ToList();
                }
                if (carry.Count > 0) yield return Encoding.UTF8.GetString(carry.ToArray()).TrimEnd('\r');
            }
        }

        private string BuildEnvironmentSnapshot()
        {
            var nl = Environment.NewLine;
            var sb = new StringBuilder();
            sb.Append("===== Environment Snapshot =====").Append(nl);
            sb.Append("Product: Orbit XLEdge").Append(nl);
            sb.Append("Time zone: ").Append(TimeZoneInfo.Local.DisplayName).Append(nl);
            sb.Append("Excel version: ").Append(ExcelVersion()).Append(", process bitness: ").Append(Environment.Is64BitProcess ? "64-bit" : "32-bit").Append(nl);
            sb.Append("OS: ").Append(Environment.OSVersion.VersionString).Append(", ").Append(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit").Append(" OS").Append(nl);
            sb.Append(".NET runtime: ").Append(RuntimeInformation.FrameworkDescription).Append(nl);
            float dpi = ScreenDpi();
            sb.Append("Screen DPI: ").Append(dpi.ToString("F0", CultureInfo.InvariantCulture)).Append(" (").Append((dpi / 96f * 100f).ToString("F0", CultureInfo.InvariantCulture)).Append("% scale)").Append(nl);
            sb.Append("Culture: ").Append(CultureInfo.CurrentCulture.Name).Append(" (UI: ").Append(CultureInfo.CurrentUICulture.Name).Append(")").Append(nl);
            sb.Append("Machine: ").Append(Environment.MachineName).Append(", User: ").Append(Environment.UserName).Append(nl);
            sb.Append("=================================");
            return sb.ToString();
        }

        // XLEdge.Shared has no Excel interop reference; read Application.Version late-bound.
        private string ExcelVersion()
        {
            try
            {
                object app = _context.ExcelApp;
                if (app == null) return "unknown";
                return Convert.ToString(app.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, app, null), CultureInfo.InvariantCulture) ?? "unknown";
            }
            catch (Exception)
            {
                return "unknown";
            }
        }

        private static float ScreenDpi()
        {
            try
            {
                using (var g = Graphics.FromHwnd(IntPtr.Zero)) return g.DpiX;
            }
            catch (Exception)
            {
                return 96f;
            }
        }

        private void WriteInitFailure(Exception ex)
        {
            string text = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss}|ERROR|Logger initialization failed: {ex.Message}|{ex}";
            try
            {
                File.WriteAllText(Path.Combine(_logsFolder, "XLEdge_Error_" + DateTime.Now.ToString("dd-MMM-yyyy", English) + ".log"), text);
            }
            catch (Exception)
            {
                try { File.WriteAllText(Path.Combine(Path.GetTempPath(), $"XLEdge_Error_{DateTime.Now:yyyyMMdd_HHmmss}.log"), text); }
                catch (Exception) { }
            }
        }
    }
}
