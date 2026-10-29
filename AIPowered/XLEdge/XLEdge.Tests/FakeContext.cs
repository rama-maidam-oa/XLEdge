using System;
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
        public IRibbonController RibbonController { get; private set; }
        public IntPtr ExcelHandle { get; set; }
        public bool DebugMode { get; set; }
        public string Version { get; set; }
        public string ReleaseDate { get; set; }
        public string ActiveFolderName { get; set; }
        public string ShellVersion { get; set; } = "11.1.2.0";
        public void SetRibbonController(IRibbonController controller) => RibbonController = controller;
        public object GetGLSenseAddinInstance() => null;
        public IntPtr GetExcelDispatch() => IntPtr.Zero;
        // As the shell's XLEdgeContext: null keys are ignored, a null value removes.
        private readonly System.Collections.Generic.Dictionary<string, string> _session = new System.Collections.Generic.Dictionary<string, string>(StringComparer.Ordinal);
        public string GetSessionValue(string key) { if (key == null) return null; lock (_session) return _session.TryGetValue(key, out string v) ? v : null; }
        public void SetSessionValue(string key, string value) { if (key == null) return; lock (_session) { if (value == null) _session.Remove(key); else _session[key] = value; } }
        public void ClearSession() { lock (_session) _session.Clear(); }
    }
}
