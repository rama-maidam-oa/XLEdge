using System;
using System.Linq;
using XLEdge.Addin.Core.Infrastructure;
using XLEdge.Addin.Core.Pane;
using XLEdge.Contracts;
using XLEdge.Helpers;
using XLEdge.Utilities;

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
                CoreServices.Initialize(context);
                // XLEdge 11.1.2 code reaches the ribbon through XLEdgeRibbonHelper.Current.
                _ = new XLEdgeRibbonHelper(context.RibbonController);
                LoadPreferences();
                StartWpfUi();
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
                KeepSessionOptions();
                // Every window of this AppDomain goes before the domain is unloaded.
                _panes?.CloseAll();
                _panes = null;
                ShutDownWpf();
                // Last: everything above may still use Excel; the core's Excel reference must not outlive it.
                CoreServices.Shutdown();
                Logger?.LogInfo("AddinEntry.Shutdown: XLEdge core shut down.");
            }
            catch (Exception ex) { Log(ex, "AddinEntry.Shutdown"); }
        }

        /// <summary>
        /// The user's preferences into XLEdgeAppState, as XLEdge 11.1.2 did when its ribbon loaded. Every new core (start-up and
        /// each hot reload) starts from defaults, so this is also what brings the preferences back after a reload. A failure
        /// leaves today's defaults (the manager itself already falls back on a damaged file).
        /// </summary>
        private void LoadPreferences()
        {
            try { XLEdgePreferencesManager.Instance.Initialize(); }
            catch (Exception ex) { Log(ex, "AddinEntry.Initialize (preferences)"); }
            RestoreSessionOptions();
        }

        private const string SessionOptionsKey = "XLEdge.SessionOptions";

        /// <summary>
        /// At Shutdown (a reload, or Excel closing): options applied for this session only (Options > Apply) differ from the saved
        /// ones; the shell keeps them so the next core of this Excel session applies them again. Saved options leave nothing.
        /// </summary>
        private void KeepSessionOptions()
        {
            try
            {
                string inForce = SerializationHelper.SerializeToJson(XLEdgePreferencesManager.Instance.GetFromAppState());
                string saved = SerializationHelper.SerializeToJson(XLEdgePreferencesManager.Instance.Current);
                _context?.SetSessionValue(SessionOptionsKey, inForce == saved ? null : inForce);
            }
            catch (Exception ex) { Log(ex, "AddinEntry.Shutdown (session options)"); }
        }

        /// <summary>After the saved preferences load: this session's Apply-only options, if a previous core left any (never saved to the file).</summary>
        private void RestoreSessionOptions()
        {
            try
            {
                string json = _context?.GetSessionValue(SessionOptionsKey);
                if (string.IsNullOrEmpty(json)) return;
                var options = System.Text.Json.JsonSerializer.Deserialize<XLEdge.Models.XLEdgeUserPreferences>(json,
                    new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (options != null) XLEdgePreferencesManager.Instance.ApplyRuntime(options);
            }
            catch (Exception ex) { Log(ex, "AddinEntry.Initialize (session options)"); }
        }

        /// <summary>
        /// Closes WPF in this AppDomain (the Application the ported windows run on, and its dispatcher's message window, whose
        /// window procedure lives here) - on its own thread, which is Excel's: the shell calls Shutdown there.
        /// </summary>
        private void ShutDownWpf()
        {
            try
            {
                System.Windows.Application app = System.Windows.Application.Current;
                if (app == null) return;
                System.Windows.Threading.Dispatcher dispatcher = app.Dispatcher;
                Action close = () =>
                {
                    // Every core window first (they would otherwise be torn down mid-operation by the dispatcher shutdown).
                    foreach (System.Windows.Window window in app.Windows.Cast<System.Windows.Window>().ToList())
                    {
                        try { window.Close(); }
                        catch (Exception ex) { Log(ex, "AddinEntry.Shutdown (closing a window)"); }
                    }
                    dispatcher.InvokeShutdown();
                };
                if (dispatcher.CheckAccess()) close();
                else dispatcher.Invoke(close);
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.Shutdown (WPF)");
            }
        }

        /// <summary>WPF-UI with XLEdge's accent and theme, and its resources preloaded - what XLEdge 11.1.2 did on ribbon load.</summary>
        private void StartWpfUi()
        {
            try
            {
                WpfUiBootstrapper.Init(XLEdgeAppConstants.GLAccentHex, XLEdgeAppConstants.GLTheme);
                WpfUiBootstrapper.PreloadResources();
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.Initialize (WPF-UI)");
            }
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

        /// <summary>The core's open windows (the shell asks before a reload, which would close them).</summary>
        public string[] GetOpenWindowTitles()
        {
            try
            {
                System.Windows.Application app = System.Windows.Application.Current;
                if (app == null) return new string[0];
                Func<string[]> titles = () => app.Windows.Cast<System.Windows.Window>().Where(w => w.IsVisible)
                    .Select(w => string.IsNullOrWhiteSpace(w.Title) ? w.GetType().Name : w.Title).ToArray();
                return app.Dispatcher.CheckAccess() ? titles() : app.Dispatcher.Invoke(titles);
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.GetOpenWindowTitles");
                return new string[0];
            }
        }

        public bool IsBusy() => false;

        public override object InitializeLifetimeService() => null;

        private void Log(Exception ex, string context)
        {
            try { Logger?.LogException(ex.ToString(), context); }
            catch (Exception) { }
        }
    }
}
