using System;
using XLEdge.Contracts;

namespace XLEdge.Tests
{
    /// <summary>Test-only IXLEdgeContext: real paths, settable DebugMode, no Excel.</summary>
    internal sealed class FakeContext : IXLEdgeContext
    {
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
    }
}
