using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                        KeepFileOpen = false,
                        DeleteOldFileOnStartup = false,
                        ArchiveAboveSize = XLEdgeAppConstants.LogMaxFileSizeBytes,  // 20MB archive size
                        MaxArchiveFiles = XLEdgeAppConstants.LogMaxArchiveFiles,
                        ArchiveFileName = XLEdgeAppPaths.LogFolder + @"\XLEdge_Logs_{#}.log"
                    };

                    var XLEdgeLoggerConfiguration = new LoggingConfiguration();
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Warn, LogLevel.Fatal, logfile);
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Error, LogLevel.Fatal, logfile);
                    XLEdgeLoggerConfiguration.AddRule(LogLevel.Fatal, LogLevel.Fatal, logfile);

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

            string header = $"Orbit XLEdge(version : {XLEdgeAppConstants.DefaultVersion} Released on : {XLEdgeAppConstants.DefaultCommitDate}). Logs As On {DateTime.Now:dddd, dd MMMM yyyy}. Time Zone: {TimeZoneInfo.Local.DisplayName}";
            sb.AppendLine(header);

            // Add underline that exactly matches the header length in characters
            sb.AppendLine(new string('-', header.Length));

            return sb.ToString();
        }
    }
}
