using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge
{
    /// <summary>
    /// What the shell gives the core: Excel, the logger and folders (owned here, in the shell's AppDomain), the ribbon, and the
    /// loaded release. Lives for the whole Excel session and survives core reloads.
    /// </summary>
    internal sealed class XLEdgeContext : MarshalByRefObject, IXLEdgeContext, Hosting.ISessionBag, IDisposable
    {
        private const string GLSenseProgId = "GLSense.AddinModule";
        private readonly Logger _logger;

        /// <param name="excelApp">The Excel Application COM object (ADX HostApplication).</param>
        /// <param name="installRoot">The folder XLEdge.dll runs from (Assembly.Location, not AppDomain.BaseDirectory: that can be Excel's folder).</param>
        public XLEdgeContext(object excelApp, string installRoot)
            : this(excelApp, new PathProvider(installRoot))
        {
        }

        internal XLEdgeContext(object excelApp, IPathProvider paths)
        {
            ExcelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp));
            Paths = paths ?? throw new ArgumentNullException(nameof(paths));
            Paths.Ensure();
            _logger = new Logger(this);
            ShellVersion = FileVersionInfo.GetVersionInfo(typeof(XLEdgeContext).Assembly.Location).FileVersion;
        }

        public object ExcelApp { get; }
        public ILogger Logger => _logger;
        public IPathProvider Paths { get; }
        public IRibbonController RibbonController { get; private set; }
        public IntPtr ExcelHandle { get; set; }
        public bool DebugMode { get; set; }
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string ActiveFolderName { get; set; }
        public string ShellVersion { get; }

        public void SetRibbonController(IRibbonController controller) => RibbonController = controller;

        /// <summary>
        /// The GLSense add-in's COM object (COMAddIns item "GLSense.AddinModule", when connected), or null. Looked up on every
        /// call: GLSense can be disconnected or reconnected during the session.
        /// </summary>
        public bool IsGLSenseConnected()
        {
            object glSense = GetGLSenseAddinInstance();
            ReleaseComObject(glSense);
            return glSense != null;
        }

        /// <summary>See <see cref="IXLEdgeContext.CallGLSense"/>. Only the method name is logged: the arguments carry the login token.</summary>
        public void CallGLSense(string method, object[] args)
        {
            object glSense = GetGLSenseAddinInstance();
            if (glSense == null)
            {
                Logger.LogDebug($"XLEdgeContext.CallGLSense: GLSense is not connected; {method} not called.");
                return;
            }
            try
            {
                glSense.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, glSense, args ?? new object[0]);
                Logger.LogDebug($"XLEdgeContext.CallGLSense: {method} called.");
            }
            catch (Exception ex)
            {
                Exception cause = ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex;
                Logger.LogWarn($"XLEdgeContext.CallGLSense: GLSense {method} failed: {cause.Message}");
                throw new InvalidOperationException($"GLSense {method} failed: {cause.Message}");
            }
            finally
            {
                ReleaseComObject(glSense);
            }
        }

        public object GetGLSenseAddinInstance()
        {
            object addins = null;
            try
            {
                addins = ExcelApp.GetType().InvokeMember("COMAddIns", BindingFlags.GetProperty, null, ExcelApp, null);
                int count = Convert.ToInt32(addins.GetType().InvokeMember("Count", BindingFlags.GetProperty, null, addins, null));
                for (int i = 1; i <= count; i++)
                {
                    object addin = addins.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, addins, new object[] { i });
                    try
                    {
                        string progId = Convert.ToString(addin.GetType().InvokeMember("ProgId", BindingFlags.GetProperty, null, addin, null));
                        if (!string.Equals(progId, GLSenseProgId, StringComparison.OrdinalIgnoreCase)) continue;
                        bool connected = Convert.ToBoolean(addin.GetType().InvokeMember("Connect", BindingFlags.GetProperty, null, addin, null));
                        if (!connected) return null;
                        object glSense = addin.GetType().InvokeMember("Object", BindingFlags.GetProperty, null, addin, null);
                        Logger.LogDebug("XLEdgeContext.GetGLSenseAddinInstance: GLSense add-in found.");
                        return glSense;
                    }
                    finally
                    {
                        ReleaseComObject(addin);
                    }
                }
                Logger.LogDebug("XLEdgeContext.GetGLSenseAddinInstance: GLSense add-in is not installed.");
                return null;
            }
            catch (Exception ex)
            {
                Logger.LogWarn($"XLEdgeContext.GetGLSenseAddinInstance: could not look for the GLSense add-in: {ex.Message}");
                return null;
            }
            finally
            {
                ReleaseComObject(addins);
            }
        }

        private readonly System.Collections.Concurrent.ConcurrentDictionary<string, string> _session =
            new System.Collections.Concurrent.ConcurrentDictionary<string, string>(StringComparer.Ordinal);

        /// <summary>See <see cref="IXLEdgeContext.GetSessionValue"/>: lives as long as this context (the Excel session).</summary>
        public string GetSessionValue(string key) => key != null && _session.TryGetValue(key, out string value) ? value : null;

        public void SetSessionValue(string key, string value)
        {
            if (key == null) return;
            if (value == null) _session.TryRemove(key, out _);
            else _session[key] = value;
        }

        public void ClearSession() => _session.Clear();

        /// <summary>See <see cref="IXLEdgeContext.GetExcelDispatch"/>: AddRef'ed, the core releases it.</summary>
        public IntPtr GetExcelDispatch()
        {
            try
            {
                return ExcelApp != null && Marshal.IsComObject(ExcelApp) ? Marshal.GetIDispatchForObject(ExcelApp) : IntPtr.Zero;
            }
            catch (Exception ex)
            {
                Logger.LogWarn("XLEdgeContext.GetExcelDispatch: could not hand Excel to the core: " + ex.Message);
                return IntPtr.Zero;
            }
        }

        public override object InitializeLifetimeService() => null;

        /// <summary>Flushes and closes the log. Called once, when the add-in shuts down.</summary>
        public void Dispose() => _logger.Dispose();

        private static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }
}
