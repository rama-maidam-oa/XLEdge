using System;
using XLEdge.Addin.Core.Pane;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core
{
    /// <summary>
    /// Entry point of the XLEdge core, created by the loader inside the core's AppDomain and called by the shell through a
    /// remoting proxy. Minimal for now: the XLEdge logic is moved in from the existing add-in in later stages; the task pane
    /// shows the Stage 1 spike page (a WebView2 per Excel window, see <see cref="TaskPaneContentHost"/>).
    /// Every member catches its own exceptions — nothing may escape back across the AppDomain boundary.
    /// </summary>
    public sealed class AddinEntry : MarshalByRefObject, IXLEdgeAddin
    {
        private IXLEdgeContext _context;
        private TaskPaneContentHost _panes;

        private ILogger Logger => _context?.Logger;

        public void Initialize(IXLEdgeContext context)
        {
            try
            {
                _context = context;
                _panes = new TaskPaneContentHost(context);
                Logger?.LogInfo($"AddinEntry.Initialize: XLEdge core initialized in AppDomain '{AppDomain.CurrentDomain.FriendlyName}' from '{AppDomain.CurrentDomain.BaseDirectory}'.");
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.Initialize");
            }
        }

        public void OnRibbonAction(string action, object parameter)
        {
            try
            {
                Logger?.LogDebug($"AddinEntry.OnRibbonAction: '{action}'.");
                // Stage 1 spike wiring (Stage 3 brings today's login and pane logic): Login opens the pane, Show/Hide toggles it.
                if (action == "RibEdgeLogin") _context?.RibbonController?.ShowTaskPane();
                else if (action == "RibEdgeShowHide") _context?.RibbonController?.ToggleTaskPane();
            }
            catch (Exception ex) { Log(ex, "AddinEntry.OnRibbonAction"); }
        }

        public bool OnExcelEvent(string eventName, object[] args)
        {
            try { Logger?.LogDebug($"AddinEntry.OnExcelEvent: '{eventName}'."); }
            catch (Exception ex) { Log(ex, "AddinEntry.OnExcelEvent"); }
            return true;
        }

        public void Shutdown()
        {
            try
            {
                // Every window of this AppDomain goes before the domain is unloaded.
                _panes?.CloseAll();
                Logger?.LogInfo("AddinEntry.Shutdown: XLEdge core shut down.");
            }
            catch (Exception ex) { Log(ex, "AddinEntry.Shutdown"); }
        }

        public IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle)
        {
            try { return _panes?.Create(excelWindowHandle) ?? IntPtr.Zero; }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.CreateTaskPaneContent");
                return IntPtr.Zero;
            }
        }

        public void CloseTaskPaneContent(IntPtr excelWindowHandle)
        {
            try { _panes?.Close(excelWindowHandle); }
            catch (Exception ex) { Log(ex, "AddinEntry.CloseTaskPaneContent"); }
        }

        public void OnTaskPaneShown(IntPtr excelWindowHandle)
        {
            try { _panes?.Shown(excelWindowHandle); }
            catch (Exception ex) { Log(ex, "AddinEntry.OnTaskPaneShown"); }
        }

        public void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission) { }

        public void OnLogoutFromGLSense() { }

        public LoginInfo GetLoginInfo() => new LoginInfo();

        public string[] GetOpenWindowTitles() => new string[0];

        public bool IsBusy() => false;

        public override object InitializeLifetimeService() => null;

        private void Log(Exception ex, string context)
        {
            try { Logger?.LogException(ex.ToString(), context); }
            catch (Exception) { }
        }
    }
}
