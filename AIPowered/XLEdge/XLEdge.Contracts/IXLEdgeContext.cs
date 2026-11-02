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

        /// <summary>
        /// Contract 5: true when the GLSense add-in is loaded and connected (looked up now).
        /// </summary>
        bool IsGLSenseConnected();

        /// <summary>
        /// Contract 5: calls <paramref name="method"/> of the GLSense add-in's COM object (looked up now) with
        /// <paramref name="args"/> (strings), in the shell's AppDomain, so no COM object crosses to the core. Does nothing when
        /// GLSense is not connected; a failed call comes back as an InvalidOperationException naming the method.
        /// </summary>
        void CallGLSense(string method, object[] args);

        /// <summary>
        /// An AddRef'ed IDispatch pointer of <see cref="ExcelApp"/> (IntPtr.Zero when there is none); the caller releases it.
        /// The core builds its own RCW from it (same process, so the pointer is valid in every AppDomain): calls then go
        /// straight to Excel instead of through a remoting proxy of the shell's RCW.
        /// </summary>
        IntPtr GetExcelDispatch();

        /// <summary>
        /// A value the shell keeps for this Excel session across core reloads (null when there is none); gone when Excel closes.
        /// For what a reload must not lose, e.g. options applied for the session only.
        /// </summary>
        string GetSessionValue(string key);

        /// <summary>Stores (or, with null, removes) a session value; see <see cref="GetSessionValue"/>.</summary>
        void SetSessionValue(string key, string value);
    }
}
