using System;

namespace XLEdge.Contracts
{
    /// <summary>
    /// What the shell gives the core in IXLEdgeAddin.Initialize. Lives in the shell's AppDomain
    /// for the whole Excel session (it survives core reloads), so its implementation must be a
    /// MarshalByRefObject with an infinite lease.
    /// </summary>
    public interface IXLEdgeContext
    {
        /// <summary>The Excel Application COM object (Microsoft.Office.Interop.Excel.Application), typed as object.</summary>
        object ExcelApp { get; }

        ILogger Logger { get; }
        IPathProvider Paths { get; }
        IRibbonController RibbonController { get; }
        IntPtr ExcelHandle { get; set; }
        bool DebugMode { get; set; }

        /// <summary>Version of the loaded core release.</summary>
        string Version { get; set; }

        string ReleaseDate { get; set; }

        /// <summary>Folder under AddinCore\Versions\ the loaded core release runs from.</summary>
        string ActiveFolderName { get; set; }

        /// <summary>Version of XLEdge.dll, the shell installed by the MSI (shown in About next to the core version).</summary>
        string ShellVersion { get; }

        void SetRibbonController(IRibbonController controller);

        /// <summary>
        /// The GLSense add-in's COM object (COMAddIns item "GLSense.AddinModule", .Object), or null
        /// when GLSense is not loaded. Used for the XLEdge/GLSense login sync.
        /// </summary>
        object GetGLSenseAddinInstance();
    }
}
