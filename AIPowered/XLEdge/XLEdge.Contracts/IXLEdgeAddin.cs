using System;

namespace XLEdge.Contracts
{
    /// <summary>
    /// Entry point of the XLEdge core, called by the shell across the AppDomain boundary.
    /// <para>
    /// Arguments are primitives, strings, string/object arrays of those, or IntPtr only: never an
    /// Add-in Express event-args object and never a live Range/Worksheet (cache what you need as
    /// strings in the shell first). Every call returns quickly; the core starts its own async work
    /// and catches every exception itself, because a Task cannot cross .NET Remoting and an
    /// exception escaping into Add-in Express reaches the user as a COM error.
    /// </para>
    /// <para>
    /// Members added after the first release must be called by the shell inside try/catch that
    /// tolerates MissingMethodException and RemotingException: Release History can load an older
    /// core that does not have them.
    /// </para>
    /// </summary>
    public interface IXLEdgeAddin
    {
        void Initialize(IXLEdgeContext context);

        /// <summary>A ribbon control was clicked or toggled; parameter is the new pressed state for check boxes, otherwise null.</summary>
        void OnRibbonAction(string action, object parameter);

        /// <summary>
        /// Forwards an Excel application event: SheetSelectionChange, SheetActivate, WorkbookActivate,
        /// SheetFollowHyperlink or SheetBeforeDelete. The return value is reserved for events that can be
        /// cancelled (false = cancel); none of these five can, so the shell currently ignores it.
        /// </summary>
        bool OnExcelEvent(string eventName, object[] args);

        /// <summary>Closes the core's windows, stops its timers and work and releases its COM references, before the AppDomain is unloaded.</summary>
        void Shutdown();

        /// <summary>
        /// Returns the window handle of a borderless core window holding the task pane content for the
        /// Excel window excelWindowHandle; the shell re-parents it into that window's ADX task pane.
        /// Idempotent: returns the existing handle if the content already exists for that window.
        /// </summary>
        IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle);

        /// <summary>Destroys the task pane content created for that Excel window (window closed, logoff, or before a reload).</summary>
        void CloseTaskPaneContent(IntPtr excelWindowHandle);

        /// <summary>The task pane of that Excel window has just been shown (ADX AfterTaskPaneShow), so the core can refresh its content.</summary>
        void OnTaskPaneShown(IntPtr excelWindowHandle);

        /// <summary>GLSense logged in and passed its session to XLEdge (the shell's public InvokedFromGLSense COM method forwards here).</summary>
        void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission);

        /// <summary>GLSense logged out and asks XLEdge to log out too (the shell's public LogoffFromAddin COM method forwards here).</summary>
        void OnLogoutFromGLSense();

        LoginInfo GetLoginInfo();

        /// <summary>Titles of the core's visible windows, so the shell can warn before a reload.</summary>
        string[] GetOpenWindowTitles();

        /// <summary>True while a report run or refresh is in progress; the shell blocks a reload while it is.</summary>
        bool IsBusy();
    }
}
