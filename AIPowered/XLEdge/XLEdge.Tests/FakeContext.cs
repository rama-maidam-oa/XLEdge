using System;
using System.Linq;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>
    /// Test-only IXLEdgeContext: real paths, settable DebugMode, no Excel. MarshalByRefObject like the shell's real
    /// context, so the core's AppDomain receives a proxy (its release folder has no copy of this test assembly).
    /// </summary>
    internal sealed class FakeContext : MarshalByRefObject, IXLEdgeContext, XLEdge.Hosting.ISessionBag
    {
        public override object InitializeLifetimeService() => null;

        public FakeContext(IPathProvider paths) { Paths = paths; }

        public object ExcelApp { get; set; }
        public ILogger Logger { get; set; }
        public IPathProvider Paths { get; }
        private IRibbonController _ribbonController;
        /// <summary>Makes the RibbonController getter throw - a core start that fails in its own code.</summary>
        public string RibbonControllerThrows { get; set; }
        public IRibbonController RibbonController
        {
            get => RibbonControllerThrows != null ? throw new InvalidOperationException(RibbonControllerThrows) : _ribbonController;
            private set => _ribbonController = value;
        }
        public IntPtr ExcelHandle { get; set; }
        public bool DebugMode { get; set; }
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string ActiveFolderName { get; set; }
        public string ShellVersion { get; set; } = "11.1.2.0";
        public void SetRibbonController(IRibbonController controller) => RibbonController = controller;
        public object GetGLSenseAddinInstance() => null;
        /// <summary>Contract 5: GLSense as the shell sees it - connected or not, the calls made, or a failing GLSense.</summary>
        public bool GLSenseConnected { get; set; }
        public bool GLSenseThrows { get; set; }
        private readonly System.Collections.Generic.List<string> _glSenseCalls = new System.Collections.Generic.List<string>();
        public string[] GLSenseCalls { get { lock (_glSenseCalls) return _glSenseCalls.ToArray(); } }
        public bool IsGLSenseConnected() => GLSenseConnected;
        public void CallGLSense(string method, object[] args)
        {
            if (!GLSenseConnected) return;
            if (GLSenseThrows) throw new InvalidOperationException("GLSense " + method + " failed: boom");
            lock (_glSenseCalls) _glSenseCalls.Add(string.Join("|", new[] { method }.Concat((args ?? new object[0]).Select(a => Convert.ToString(a)))));
        }
        /// <summary>As the shell hands Excel over (an AddRef'ed IDispatch) when ExcelApp is a COM object - a real Excel in the Excel tests.</summary>
        public IntPtr GetExcelDispatch() =>
            ExcelApp != null && System.Runtime.InteropServices.Marshal.IsComObject(ExcelApp) ? System.Runtime.InteropServices.Marshal.GetIDispatchForObject(ExcelApp) : IntPtr.Zero;
        // As the shell's XLEdgeContext: null keys are ignored, a null value removes.
        private readonly System.Collections.Generic.Dictionary<string, string> _session = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
        public string GetSessionValue(string key) { if (key == null) return null; lock (_session) return _session.TryGetValue(key, out string v) ? v : null; }
        public void SetSessionValue(string key, string value) { if (key == null) return; lock (_session) { if (value == null) _session.Remove(key); else _session[key] = value; } }
        public void ClearSession() { lock (_session) _session.Clear(); }
    }
}
