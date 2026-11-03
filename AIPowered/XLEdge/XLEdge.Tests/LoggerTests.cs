using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class LoggerTests
    {
        private static readonly CultureInfo English = CultureInfo.GetCultureInfo("en-US");

        private sealed class Fixture : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();

            public Fixture(bool debugMode = true)
            {
                Paths = new PathProvider(_temp.Combine("inst"), _temp.Combine("data"));
                Context = new FakeContext(Paths) { DebugMode = debugMode };
                Logger = new Logger(Context);
            }

            public PathProvider Paths { get; }
            public FakeContext Context { get; }
            public Logger Logger { get; private set; }

            public string LogFile => Path.Combine(Paths.Logs, "XLEdge_Logs_" + DateTime.Now.ToString("dd-MMM-yyyy", English) + ".log");

            public string Text()
            {
                using (var s = new FileStream(LogFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var r = new StreamReader(s))
                    return r.ReadToEnd();
            }

            public string[] Lines() => Text().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            public void Restart()
            {
                Logger.Dispose();
                Logger = new Logger(Context);
            }

            public void Dispose()
            {
                Logger?.Dispose();
                _temp.Dispose();
            }
        }

        [TestMethod]
        public void FileName_EnglishMonthUnderFrFR()
        {
            var previous = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            try
            {
                using (var f = new Fixture())
                {
                    f.Logger.LogInfo("hello");
                    Assert.IsTrue(File.Exists(f.LogFile), "expected " + f.LogFile);
                    Assert.AreEqual(1, Directory.GetFiles(f.Paths.Logs).Length);
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [TestMethod]
        public void Header_HasProductAndSnapshot()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogInfo("hello");
                string text = f.Text();
                StringAssert.StartsWith(text, "===== Environment Snapshot =====");
                StringAssert.Contains(text, "Product: Orbit XLEdge");
                StringAssert.Contains(text, "Excel version: unknown");
                StringAssert.Contains(text, "=================================");
            }
        }

        [TestMethod]
        public void Info_LineFormat()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogInfo("hello");
                Assert.IsTrue(f.Lines().Any(l => Regex.IsMatch(l, @"^\S+ \S+\|INFO  \| \d\d:\d\d:\d\d \| hello$")), f.Text());
            }
        }

        [TestMethod]
        public void Debug_DroppedWhenDebugOff()
        {
            using (var f = new Fixture(debugMode: false))
            {
                f.Logger.LogDebug("invisible");
                using (f.Logger.BeginLogScope("Refresh")) f.Logger.LogDebug("also invisible");
                f.Logger.LogInfo("visible");
                StringAssert.DoesNotMatch(f.Text(), new Regex("invisible|Refresh"));
            }
        }

        [TestMethod]
        public void Debug_OutsideScope_WrittenImmediately()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogDebug("now");
                Assert.IsTrue(f.Lines().Any(l => Regex.IsMatch(l, @"\|DEBUG \| \d\d:\d\d:\d\d \| now$")));
            }
        }

        [TestMethod]
        public void Scope_BuffersDebugAndFlushesOneBlockOnDispose()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Refresh"))
                {
                    f.Logger.LogDebug("step 1");
                    Assert.IsFalse(File.Exists(f.LogFile) && f.Text().Contains("step 1"), "buffered until the scope ends");
                }
                string text = f.Text();
                Assert.IsTrue(Regex.IsMatch(text, @"\|===== Refresh \| \d{4}-\d\d-\d\d \d\d:\d\d:\d\d =====\r?\n-+\r?\nDEBUG \| \d\d:\d\d:\d\d \| BEGIN: Refresh\r?\n  DEBUG \| \d\d:\d\d:\d\d \| step 1\r?\nDEBUG \| \d\d:\d\d:\d\d \| END: Refresh\r?\n-+"), text);
            }
        }

        [TestMethod]
        public void Scope_FromSecondAppDomain_BuffersLikeInDomain()
        {
            using (var f = new Fixture())
            {
                var domain = AppDomain.CreateDomain("CoreLike", null, new AppDomainSetup
                {
                    ApplicationBase = Path.GetDirectoryName(typeof(LoggerTests).Assembly.Location),
                });
                try
                {
                    var caller = (RemoteScopeCaller)domain.CreateInstanceAndUnwrap(typeof(RemoteScopeCaller).Assembly.FullName, typeof(RemoteScopeCaller).FullName);
                    caller.Run(f.Logger, "SyncScope");
                    caller.RunWithAwait(f.Logger, "AsyncScope");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }

                string text = f.Text();
                foreach (var name in new[] { "SyncScope", "AsyncScope" })
                {
                    Assert.IsTrue(Regex.IsMatch(text,
                        @"\|===== " + name + @" \| [^\r\n]+=====\r?\n-+\r?\nDEBUG \| \d\d:\d\d:\d\d \| BEGIN: " + name + @"\r?\n  DEBUG \| \d\d:\d\d:\d\d \| inside\r?\nDEBUG \| \d\d:\d\d:\d\d \| END: " + name + @"\r?\n-+"),
                        name + " was not buffered as one block:\n" + text);
                }
            }
        }

        [TestMethod]
        public void Dispose_WritesTheDebugLinesStillBuffered()
        {
            // Excel closing inside a running action: what was buffered so far is the most useful part of the log.
            using (var f = new Fixture())
            {
                var scope = f.Logger.BeginLogScope("Refresh");
                f.Logger.LogDebug("buffered before shutdown");
                f.Logger.Dispose();
                StringAssert.Contains(f.Text(), "buffered before shutdown");
                f.Logger.LogInfo("after dispose");
                Assert.IsFalse(f.Text().Contains("after dispose"), "nothing is written once disposed");
                GC.KeepAlive(scope);
            }
        }

        [TestMethod]
        public void Writes_WaitForTheCrossProcessLogLock()
        {
            // Each write opens the file, seeks to the end and writes; two Excel processes doing that at once can land on the same
            // offset and overwrite each other's line. Writes therefore take a lock per logs folder that every process shares.
            using (var f = new Fixture())
            using (var held = new Mutex(false, Logger.LockName(f.Paths.Logs)))
            {
                Assert.IsTrue(held.WaitOne(0));
                var write = System.Threading.Tasks.Task.Run(() => f.Logger.LogInfo("waited for the lock"));
                Assert.IsFalse(write.Wait(300), "the write must wait while another process holds the lock");
                held.ReleaseMutex();
                Assert.IsTrue(write.Wait(5000));
                StringAssert.Contains(f.Text(), "| waited for the lock");
            }
        }

        [TestMethod]
        public void Writes_LockHeldTooLong_StillWrite()
        {
            // A hung process holding the lock must not freeze logging (and so Excel): after the wait the line is written anyway.
            using (var f = new Fixture())
            using (var held = new Mutex(false, Logger.LockName(f.Paths.Logs)))
            {
                Assert.IsTrue(held.WaitOne(0));
                var clock = System.Diagnostics.Stopwatch.StartNew();
                var write = System.Threading.Tasks.Task.Run(() => f.Logger.LogInfo("written without the lock"));
                Assert.IsTrue(write.Wait(Logger.LockTimeout + TimeSpan.FromSeconds(3)));
                Assert.IsTrue(clock.Elapsed >= Logger.LockTimeout - TimeSpan.FromMilliseconds(200), $"waited only {clock.Elapsed} for the lock");
                held.ReleaseMutex();
                StringAssert.Contains(f.Text(), "| written without the lock");
            }
        }

        [TestMethod]
        public void Writes_AfterATimeout_DoNotWaitAgainForAWhile()
        {
            // A stuck process holding the lock costs one wait, not one per line: an action logging 100 lines must not freeze
            // Excel for minutes.
            using (var f = new Fixture())
            using (var held = new Mutex(false, Logger.LockName(f.Paths.Logs)))
            {
                Assert.IsTrue(held.WaitOne(0));
                var clock = System.Diagnostics.Stopwatch.StartNew();
                var writes = System.Threading.Tasks.Task.Run(() => { for (int i = 0; i < 5; i++) f.Logger.LogInfo("line " + i); });
                Assert.IsTrue(writes.Wait(TimeSpan.FromSeconds(30)));
                Assert.IsTrue(clock.Elapsed < Logger.LockTimeout + Logger.LockTimeout, $"5 writes took {clock.Elapsed}");
                held.ReleaseMutex();
                Assert.AreEqual(5, f.Lines().Count(l => Regex.IsMatch(l, @"\| line \d$")));
            }
        }

        private sealed class BadLogsPaths : IPathProvider
        {
            public string Root => @"C:\bad|folder";
            public string Logs => @"C:\bad|folder\Logs";
            public string Temp => Root;
            public string BrowserDataPath => Root;
            public string UrlsFile => Root;
            public string VersionsPath => Root;
            public string SharedDependenciesPath => Root;
            public string ManifestFile => Root;
            public string ManifestDirectory => Root;
            public string ReleaseHistoryFile => Root;
            public string LatestVersion => null;
            public string LatestReleaseDate => null;
            public void Ensure() { }
            public void Refresh() { }
        }

        [TestMethod]
        public void Ctor_UnusableLogsFolder_DoesNotThrow()
        {
            // The add-in must start even when the log cannot be written: the failure is reported, never thrown.
            using (var logger = new Logger(new FakeContext(new BadLogsPaths())))
            {
                logger.LogInfo("goes nowhere");
            }
        }

        [TestMethod]
        public void LockName_IsPerLogsFolderAndCaseInsensitive()
        {
            Assert.AreEqual(Logger.LockName(@"C:\Data\XLEdge_Logs\Logs"), Logger.LockName(@"c:\data\xledge_logs\LOGS\"));
            Assert.AreNotEqual(Logger.LockName(@"C:\A\Logs"), Logger.LockName(@"C:\B\Logs"));
            StringAssert.StartsWith(Logger.LockName(@"C:\A\Logs"), @"Global\XLEdge_Log_");
        }

        [TestMethod]
        public void TwoLoggersSameFile_KeepEveryLine()
        {
            // Two Excel processes log to the same day's file; neither may overwrite the other's lines.
            using (var f = new Fixture())
            using (var second = new Logger(new FakeContext(f.Paths) { DebugMode = true }))
            {
                // Enough lines that two writers racing for the end of the file collide on every run, not now and then.
                const int Count = 2000;
                var a = new Thread(() => { for (int i = 0; i < Count; i++) f.Logger.LogInfo("A" + i); });
                var b = new Thread(() => { for (int i = 0; i < Count; i++) second.LogInfo("B" + i); });
                a.Start(); b.Start(); a.Join(); b.Join();
                var lines = f.Lines();
                Assert.AreEqual(Count, lines.Count(l => Regex.IsMatch(l, @"\| A\d+$")), "lines from the first logger");
                Assert.AreEqual(Count, lines.Count(l => Regex.IsMatch(l, @"\| B\d+$")), "lines from the second logger");
            }
        }

        [TestMethod]
        [Timeout(10000)]
        public void ReadLinesBackwards_StreamShorterThanReported_Terminates()
        {
            // A log file truncated while it is being read must not hang Excel's startup.
            var lines = Logger.ReadLinesBackwards(new ShortStream(reportedLength: 200_000, actual: "a\nb\n")).ToList();
            Assert.IsTrue(lines.Count < 10);
        }

        private sealed class ShortStream : MemoryStream
        {
            private readonly long _reported;
            public ShortStream(long reportedLength, string actual) : base(System.Text.Encoding.UTF8.GetBytes(actual)) { _reported = reportedLength; }
            public override long Length => _reported;
            public override int Read(byte[] buffer, int offset, int count) => Position >= base.Length ? 0 : base.Read(buffer, offset, count);
            public override long Position
            {
                get => base.Position;
                set => base.Position = Math.Min(value, base.Length);
            }
        }

        [TestMethod]
        public void Warn_InsideScope_EarlierDebugLinesComeFirst()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Run"))
                {
                    f.Logger.LogDebug("before");
                    f.Logger.LogWarn("careful");
                    f.Logger.LogError("broken");
                    f.Logger.LogDebug("between");
                    f.Logger.LogException("boom", "ctx");
                }
                string text = f.Text();
                int before = text.IndexOf("| before", StringComparison.Ordinal), careful = text.IndexOf("| careful", StringComparison.Ordinal);
                int between = text.IndexOf("| between", StringComparison.Ordinal), exception = text.IndexOf("========== Exception", StringComparison.Ordinal);
                Assert.IsTrue(before >= 0 && before < careful, "debug line logged before the warning must come first:\n" + text);
                Assert.IsTrue(between >= 0 && between < exception, "debug line logged before the exception must come first:\n" + text);
            }
        }

        [TestMethod]
        public void ConcurrentFlushes_KeepLineOrder()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Busy"))
                {
                    using (var stop = new ManualResetEventSlim())
                    {
                        ThreadStart flusher = () => { while (!stop.IsSet) f.Logger.FlushDebugLogs("tick"); };
                        var t1 = new Thread(flusher); var t2 = new Thread(flusher);
                        t1.Start(); t2.Start();
                        for (int i = 0; i < 2000; i++) f.Logger.LogDebug("n" + i.ToString("D5"));
                        stop.Set(); t1.Join(); t2.Join();
                    }
                }
                var numbers = f.Lines().Select(l => Regex.Match(l, @"\| n(\d{5})$")).Where(m => m.Success).Select(m => int.Parse(m.Groups[1].Value)).ToList();
                Assert.AreEqual(2000, numbers.Count);
                for (int i = 1; i < numbers.Count; i++)
                    Assert.IsTrue(numbers[i] > numbers[i - 1], $"line n{numbers[i]:D5} written after n{numbers[i - 1]:D5}");
            }
        }

        [TestMethod]
        public void Timestamps_AreInvariantUnderFiFIAndThai()
        {
            var previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                foreach (var culture in new[] { "fi-FI", "th-TH" })
                {
                    Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo(culture);
                    using (var f = new Fixture())
                    {
                        f.Logger.LogInfo("hello");
                        using (f.Logger.BeginLogScope("Scope")) f.Logger.LogDebug("x");
                        string text = f.Text();
                        Assert.IsTrue(Regex.IsMatch(text, @"\|INFO  \| \d\d:\d\d:\d\d \| hello"), culture + ":\n" + text);
                        Assert.IsTrue(Regex.IsMatch(text, @"===== Scope \| " + DateTime.Now.Year + @"-\d\d-\d\d \d\d:\d\d:\d\d ====="), culture + ":\n" + text);
                    }
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [TestMethod]
        public void Ctor_NullPaths_ThrowsArgumentException()
        {
            Assert.ThrowsExactly<ArgumentException>(() => new Logger(new FakeContext(null)));
        }

        [TestMethod]
        public void ReleaseSession_SameReleaseAfterEnd_StartsNewSection()
        {
            using (var f = new Fixture())
            {
                f.Logger.BeginReleaseSession("11.1.2.0", "d");
                f.Logger.EndReleaseSession("11.1.2.0", "d");
                f.Logger.BeginReleaseSession("11.1.2.0", "d");
                Assert.AreEqual(2, Regex.Matches(f.Text(), "Start of XLEdge version 11.1.2 ").Count);
            }
        }

        [TestMethod]
        public void ReleaseSession_FirstLineAfterTheSnapshot_IsTheMarker_DebugLinesOfTheStartFollowIt()
        {
            using (var f = new Fixture(debugMode: true))
            {
                using (f.Logger.BeginLogScope("XLEdge start"))
                {
                    f.Logger.LogDebug("resolved the release");
                    f.Logger.BeginReleaseSession("11.1.3.0", "d");
                    f.Logger.LogDebug("core loaded");
                }
                string[] lines = f.Lines();
                int snapshotEnd = Array.FindIndex(lines, l => l.StartsWith("=================================", StringComparison.Ordinal));
                Assert.IsTrue(snapshotEnd >= 0, f.Text());
                StringAssert.Contains(lines[snapshotEnd + 1], "===== Start of XLEdge version 11.1.3 (released d) logs as on ", f.Text());
                int resolved = Array.FindIndex(lines, l => l.Contains("resolved the release"));
                Assert.IsTrue(resolved > snapshotEnd + 1, "the start's debug lines come after the marker:\n" + f.Text());
            }
        }

        [TestMethod]
        public void ReleaseSession_DebugOff_TheSnapshotIsFollowedByTheMarkerOnly()
        {
            using (var f = new Fixture(debugMode: false))
            {
                using (f.Logger.BeginLogScope("XLEdge start"))
                {
                    f.Logger.LogDebug("resolved the release");
                    f.Logger.BeginReleaseSession("11.1.3.0", "d");
                    f.Logger.LogDebug("core loaded");
                }
                string[] lines = f.Lines().Where(l => l.Length > 0).ToArray();
                int snapshotEnd = Array.FindIndex(lines, l => l.StartsWith("=================================", StringComparison.Ordinal));
                Assert.AreEqual(snapshotEnd + 2, lines.Length, f.Text());
                StringAssert.Contains(lines[snapshotEnd + 1], "===== Start of XLEdge version 11.1.3 (released d) logs as on ");
                Assert.IsFalse(f.Text().Contains("INFO"), f.Text());
            }
        }

        [TestMethod]
        public void Scope_NestedReusesBuffer()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Outer"))
                {
                    using (f.Logger.BeginLogScope("Inner")) f.Logger.LogDebug("deep");
                    Assert.IsFalse(File.Exists(f.LogFile) && f.Text().Contains("deep"), "inner scope must not flush");
                }
                string text = f.Text();
                Assert.AreEqual(1, Regex.Matches(text, "=====  ?Outer|===== Outer").Count, text);
                StringAssert.Contains(text, "  DEBUG");
                StringAssert.Contains(text, "    DEBUG");
                StringAssert.DoesNotMatch(text, new Regex("===== Inner"));
            }
        }

        [TestMethod]
        public void Warn_InsideScope_FlushesWithReason()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Run"))
                {
                    f.Logger.LogDebug("before");
                    f.Logger.LogWarn("careful");
                    string text = f.Text();
                    StringAssert.Contains(text, "WARN  | ");
                    StringAssert.Contains(text, "===== Run - warning logged | ");
                    StringAssert.Contains(text, "before");
                }
            }
        }

        [TestMethod]
        public void Error_AppendsDetails()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogError("failed", "System.IO.IOException: disk full");
                f.Logger.LogError("plain");
                string text = f.Text();
                Assert.IsTrue(Regex.IsMatch(text, @"\|ERROR \| \d\d:\d\d:\d\d \| failed\|System\.IO\.IOException: disk full"), text);
                Assert.IsTrue(f.Lines().Any(l => l.EndsWith("| plain")), text);
            }
        }

        [TestMethod]
        public void LogException_WritesBlock_AndNullDetailsDoNotThrow()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogException("System.InvalidOperationException: boom\r\n   at X.Y()", "Refreshing sheet");
                f.Logger.LogException(null);
                f.Logger.LogError("x", null);
                string text = f.Text();
                StringAssert.Contains(text, "========== Exception ==========");
                StringAssert.Contains(text, "Context: Refreshing sheet");
                StringAssert.Contains(text, "System.InvalidOperationException: boom");
                StringAssert.Contains(text, "<no details>");
                StringAssert.Contains(text, "============================");
            }
        }

        [TestMethod]
        public void RawJson_Format()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogRawJson("Report definition", "{\"a\":1}");
                f.Logger.LogRawJson("Empty", "");
                string text = f.Text();
                StringAssert.Contains(text, "----- Raw JSON (Report definition) -----");
                StringAssert.Contains(text, "{\"a\":1}");
                StringAssert.Contains(text, "<empty>");
                StringAssert.Contains(text, "----- End Raw JSON -----");
            }
        }

        [TestMethod]
        public void ReleaseSession_BeginNotDuplicated_EndFlushes()
        {
            using (var f = new Fixture())
            {
                f.Logger.BeginReleaseSession("11.1.2.0", "2026-10-28T10:00:00");
                f.Restart();
                f.Logger.BeginReleaseSession("11.1.2.0", "2026-10-28T10:00:00");
                Assert.AreEqual(1, Regex.Matches(f.Text(), "Start of XLEdge version").Count, "same release must not start a second section");

                using (f.Logger.BeginLogScope("Pending")) { f.Logger.LogDebug("unflushed"); f.Logger.EndReleaseSession("11.1.2.0", "2026-10-28T10:00:00"); }
                string text = f.Text();
                StringAssert.Contains(text, "===== Start of XLEdge version 11.1.2 (released 2026-10-28T10:00:00) logs as on ");
                StringAssert.Contains(text, "===== Pending - release session ending | ");
                StringAssert.Contains(text, "===== End of XLEdge version 11.1.2 (released 2026-10-28T10:00:00) logs =====");

                f.Logger.BeginReleaseSession("11.1.3.0", null);
                StringAssert.Contains(f.Text(), "===== Start of XLEdge version 11.1.3 (released unknown release date) logs as on ");
            }
        }

        [TestMethod]
        public void ReleaseSession_DedupeWorksUnderFrFR()
        {
            // The "today's log file" lookup must use the same (English) month name as the file NLog writes,
            // otherwise on non-English Windows every Excel start adds a duplicate section. fr-FR month
            // abbreviations differ from English in every month (de-DE only in four).
            var previous = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
            try
            {
                using (var f = new Fixture())
                {
                    f.Logger.BeginReleaseSession("11.1.2.0", "2026-10-28T10:00:00");
                    f.Restart();
                    f.Logger.BeginReleaseSession("11.1.2.0", "2026-10-28T10:00:00");
                    Assert.AreEqual(1, Regex.Matches(f.Text(), "Start of XLEdge version").Count);
                }
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [TestMethod]
        public void SafetyNet_FlushesBufferOlderThan30Seconds()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Long"))
                {
                    f.Logger.LogDebug("still running");
                    f.Logger.SweepSafetyNet(DateTime.UtcNow.AddSeconds(10));
                    Assert.IsFalse(File.Exists(f.LogFile) && f.Text().Contains("still running"));
                    f.Logger.SweepSafetyNet(DateTime.UtcNow.AddSeconds(31));
                    StringAssert.Contains(f.Text(), "===== Long - safety-net flush, action still running | ");
                }
            }
        }

        [TestMethod]
        public void FlushDebugLogs_FlushesOpenScopes()
        {
            using (var f = new Fixture())
            {
                using (f.Logger.BeginLogScope("Work"))
                {
                    f.Logger.LogDebug("pending");
                    f.Logger.FlushDebugLogs("unhandled exception");
                    StringAssert.Contains(f.Text(), "===== Work - unhandled exception | ");
                }
            }
        }

        [TestMethod]
        public void Lease_IsInfinite()
        {
            using (var f = new Fixture())
            {
                Assert.IsNull(f.Logger.InitializeLifetimeService());
                var scope = f.Logger.BeginLogScope("s");
                Assert.IsInstanceOfType(scope, typeof(MarshalByRefObject));
                Assert.IsNull(((MarshalByRefObject)scope).InitializeLifetimeService());
                scope.Dispose();
            }
        }

        [TestMethod]
        public void Dispose_ReleasesFile()
        {
            using (var f = new Fixture())
            {
                f.Logger.LogInfo("hello");
                f.Logger.Dispose();
                File.Delete(f.LogFile);
                Assert.IsFalse(File.Exists(f.LogFile));
            }
        }

        [TestMethod]
        public void Ctor_NullContext_Throws()
        {
            Assert.ThrowsExactly<ArgumentNullException>(() => new Logger(null));
        }
    }
}
