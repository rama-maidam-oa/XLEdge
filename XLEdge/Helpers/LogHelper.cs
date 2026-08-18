using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge.Helpers
{
    public static class LogHelper
    {
        private static bool _isInitialized = false;
        private static readonly object _initLock = new();

        public static void InitializeLogger()
        {
            lock (_initLock)
            {
                if (_isInitialized) return;

                try
                {
                    // Create header layout with dynamic date evaluation
                    string HdrText = BuildLogHeader();

                    // IMPORTANT: Use NLog's date pattern, NOT a pre-evaluated date
                    string fileNamePattern = XLEdgeAppPaths.LogFolder + @"\XLEdge_Logs_${date:format=dd-MMM-yyyy}.log";
                    var fileNameLayout = NLog.Layouts.Layout.FromString(fileNamePattern);

                    var logfile = new FileTarget("logfile")
                    {
                        FileName = fileNameLayout,  // This will be evaluated at runtime
                        Header = HdrText,  // NLog will write this header when creating new files
                        AutoFlush = true,
                        // Includes callsite info (file/method) to make it easier to trace which method logged what.
                        Layout = "${longdate}|${level:uppercase=true}|${callsite:className=false:fileName=true:includeSourcePath=false:methodName=true}|${message:withException=true:exceptionSeparator=|}",
                        // Was false - every single log call opened, wrote, flushed, and closed the file
                        // handle. Combined with LogUtility's per-action buffering, most Debug lines now
                        // arrive as one batched write per action rather than one write per line, so
                        // keeping the handle open between writes is both safe and meaningfully faster.
                        KeepFileOpen = true,
                        DeleteOldFileOnStartup = false,
                        ArchiveAboveSize = XLEdgeAppConstants.LogMaxFileSizeBytes,  // 20MB archive size
                        MaxArchiveFiles = XLEdgeAppConstants.LogMaxArchiveFiles,
                        ArchiveFileName = XLEdgeAppPaths.LogFolder + @"\XLEdge_Logs_{#}.log"
                    };

                    var XLEdgeLoggerConfiguration = new LoggingConfiguration();

                    // Single rule covering the required level range to avoid duplicate routing.
                    // The 5 overlapping per-level rules this replaced (Info-Fatal, Debug-Fatal,
                    // Warn-Fatal, Error-Fatal, Fatal-Fatal) all targeted the same "logfile"
                    // target, so any Info/Warn/Error/Fatal event matched more than one rule and
                    // was written to the file multiple times per event.
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);

                    LogManager.Configuration = XLEdgeLoggerConfiguration;

                    // Store the logger instance
                    AddinModule.LoggerConfiguration = XLEdgeLoggerConfiguration;
                    AddinModule.Logger = LogManager.GetCurrentClassLogger();

                    _isInitialized = true;
                }
                catch (Exception ex)
                {
                    // Route through the shared XLEdgeMessageWindow used elsewhere in this app.
                    try
                    {
                        XLEdge.Utilities.MessageFunctions.XLEdgeMessage(
                            $"Error initializing logger: {ex.Message}",
                            MessageBoxIcon.Error,
                            MessageBoxButtons.OK);
                    }
                    catch
                    {
                        // The logger failing to initialize is early enough in startup that even the
                        // custom window's own WPF-thread infrastructure might not be available yet -
                        // fall back to a native message box only as an absolute last resort so this
                        // failure is never silently swallowed.
                        MessageBox.Show($"Error initializing logger: {ex.Message}", "Orbit XLEdge", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private static string BuildLogHeader()
        {

            var sb = new StringBuilder();

            string header = $"Orbit XLEdge logs generated on: {DateTime.Now:dddd, dd MMMM yyyy}. Time Zone: {TimeZoneInfo.Local.DisplayName}";
            sb.AppendLine(header);

            // Add underline that exactly matches the header length in characters
            sb.AppendLine(new string('-', header.Length));

            AppendEnvironmentSnapshot(sb);

            return sb.ToString();
        }

        // Moved here from AddinModule.LogEnvironmentSnapshot, which used to log this as a
        // regular LogInfo call from AddinModule_OnRibbonLoaded - that method fires once per
        // Excel session, but the log file itself is per-day (XLEdge_Logs_{date}.log), so
        // every subsequent Excel open on the same day re-appended an identical snapshot into
        // that day's file instead of writing it once. NLog's FileTarget.Header is only
        // written when the target actually creates a new file, so folding this into the
        // header instead makes it genuinely once-per-file (once-per-day) for free, with no
        // new file-existence tracking needed here.
        //
        // Runs BEFORE LogManager.Configuration is assigned (see InitializeLogger above), so
        // LogUtility/AddinModule.Logger calls are not usable yet here - failures are
        // swallowed silently with a safe fallback value instead of being logged, unlike the
        // original method's LogWarn calls.
        private static void AppendEnvironmentSnapshot(StringBuilder sb)
        {
            string excelVersion = "unknown";
            try
            {
                excelVersion = (AddinModule.CurrentInstance?.HostApplication as Excel.Application)?.Version ?? "unknown";
            }
            catch
            {
                // HostApplication may not be assigned yet depending on call order - "unknown"
                // is an acceptable fallback for a one-time header line.
            }

            double dpi = 96d;
            try
            {
                using (var g = System.Drawing.Graphics.FromHwnd(IntPtr.Zero))
                {
                    dpi = g.DpiX;
                }
            }
            catch
            {
                // Fall back to the 96 DPI (100% scale) default set above.
            }

            sb.AppendLine("===== Environment Snapshot =====");
            sb.AppendLine($"XLEdge version: {XLEdgeAppConstants.DefaultVersion} (released {XLEdgeAppConstants.DefaultCommitDate})");
            sb.AppendLine($"Excel version: {excelVersion}, process bitness: {(Environment.Is64BitProcess ? "64-bit" : "32-bit")}");
            sb.AppendLine($"OS: {Environment.OSVersion.VersionString}, {(Environment.Is64BitOperatingSystem ? "64-bit" : "32-bit")} OS");
            sb.AppendLine($".NET runtime: {System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}");
            sb.AppendLine($"Screen DPI: {dpi:F0} ({dpi / 96d * 100:F0}% scale)");
            sb.AppendLine($"Culture: {CultureInfo.CurrentCulture.Name} (UI: {CultureInfo.CurrentUICulture.Name})");
            sb.AppendLine($"Machine: {Environment.MachineName}, User: {Environment.UserName}");
            sb.AppendLine("=================================");
        }
    }
}
