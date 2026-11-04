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
        private AddinModule _module;

        private ILogger Logger => _context?.Logger;

        /// <summary>
        /// Starts the core. A failure of its own parts (services, ribbon helper, Excel's thread context, panes, the module) is
        /// thrown back - a core without them would ignore every click - so the loader unloads it and the shell says so; the
        /// best-effort steps after them only log. The exception carries text only: no core type may cross to the shell.
        /// </summary>
        public void Initialize(IXLEdgeContext context)
        {
            try
            {
                _context = context;
                CoreServices.Initialize(context);
                // XLEdge 11.1.2 code reaches the ribbon through XLEdgeRibbonHelper.Current.
                _ = new XLEdgeRibbonHelper(context.RibbonController);
                ExcelThreadContext.Initialize();
                _panes = new TaskPaneContentHost(context);
                _module = new AddinModule(XLEdgeRibbonHelper.Current, _panes);
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.Initialize (core start)");
                // The loader unloads this domain without a Shutdown: release the Excel reference already taken.
                try { _module?.Shutdown(); _module = null; CoreServices.Shutdown(); }
                catch (Exception cleanup) { Log(cleanup, "AddinEntry.Initialize (cleanup after a failed start)"); }
                throw new InvalidOperationException("XLEdge core start failed: " + ex.Message);
            }
            try
            {
                RestoreSessionLogin();
                _module.ApplyStartRibbonState();
                LoadPreferences();
                DeleteTempFiles();
                StartWpfUi();
                Logger?.LogDebug($"AddinEntry.Initialize: XLEdge core initialized in AppDomain '{AppDomain.CurrentDomain.FriendlyName}' from '{AppDomain.CurrentDomain.BaseDirectory}'.");
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
                // One line per click; the ribbon's property-changing callbacks run at every invalidation - not logged.
                if (action == null || !action.EndsWith("_PropertyChanging", StringComparison.Ordinal))
                    Logger?.LogDebug($"AddinEntry.OnRibbonAction: '{action}'.");
                ExcelThreadContext.Run(() =>
                {
                    _module?.OnRibbonAction(action, parameter);
                });
            }
            catch (Exception ex) { Log(ex, "AddinEntry.OnRibbonAction"); }
        }

        public bool OnExcelEvent(string eventName, object[] args)
        {
            try
            {
                ExcelThreadContext.Run(() => _module?.OnExcelEvent(eventName, args));
            }
            catch (Exception ex) { Log(ex, "AddinEntry.OnExcelEvent"); }
            return true;
        }

        public void Shutdown()
        {
            try
            {
                // Each step in its own try (ShutdownSteps): a failing one must not keep the later ones from running.
                ShutdownSteps.RunAll("AddinEntry.Shutdown",
                    ("session options", KeepSessionOptions),
                    ("session login", KeepSessionLogin),
                    // No deferred sheet delete may fire into the unloaded AppDomain.
                    ("module", () => { AddinModule module = _module; _module = null; module?.Shutdown(); }),
                    // Every window of this AppDomain goes before the domain is unloaded.
                    ("panes", () => { TaskPaneContentHost panes = _panes; _panes = null; panes?.CloseAll(); }),
                    ("wpf", ShutDownWpf),
                    // Last: everything above may still use Excel; the core's Excel reference must not outlive it.
                    ("core services", CoreServices.Shutdown));
                Logger?.LogDebug("AddinEntry.Shutdown: XLEdge core shut down.");
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

        /// <summary>
        /// Temp report files left in XLEdgeAppPaths.TempFolder - by a crash, or by the previous core of this session - go, as
        /// 11.1.2 did when its ribbon loaded (AddinModule.vb's EEDeleteAllFiles). No report runs now: a reload waits for IsBusy.
        /// </summary>
        private void DeleteTempFiles()
        {
            try { XLEdgeTempFileCleaner.DeleteAllTempFiles(); }
            catch (Exception ex) { Log(ex, "AddinEntry.Initialize (temp files)"); }

            // Debug output files (full API responses saved by LargeLogText instead of logged - ported from GLSense
            // OISR-22830) older than 7 days go. Best effort: CleanUp never throws for I/O.
            try { LargeLogText.CleanUp(XLEdgeAppPaths.TempFolder); }
            catch (Exception ex) { Log(ex, "AddinEntry.Initialize (debug output files)"); }
        }

        private void KeepSessionLogin()
        {
            try { SessionLogin.Keep(_context); }
            catch (Exception ex) { Log(ex, "AddinEntry.Shutdown (session login)"); }
        }

        /// <summary>Before the start ribbon state: a login the previous core left makes this one start LoggedIn.</summary>
        private void RestoreSessionLogin()
        {
            try
            {
                if (SessionLogin.Restore(_context))
                {
                    Logger?.LogDebug("AddinEntry.Initialize: the login of the previous core is back.");
                    _module?.ReattachGLSense();
                }
            }
            catch (Exception ex) { Log(ex, "AddinEntry.Initialize (session login)"); }
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
                // Starting from the saved options, so one an older core did not know keeps its saved value instead of false.
                XLEdge.Models.XLEdgeUserPreferences options = XLEdgePreferencesManager.Instance.Current;
                using (var document = System.Text.Json.JsonDocument.Parse(json))
                {
                    System.Text.Json.JsonElement root = document.RootElement;
                    bool Read(string name, bool saved) =>
                        root.ValueKind == System.Text.Json.JsonValueKind.Object && root.TryGetProperty(name, out System.Text.Json.JsonElement value) &&
                        (value.ValueKind == System.Text.Json.JsonValueKind.True || value.ValueKind == System.Text.Json.JsonValueKind.False)
                            ? value.GetBoolean() : saved;
                    options.ParameterValues = Read("parameterValues", options.ParameterValues);
                    options.ScheduledOutputs = Read("scheduledOutputs", options.ScheduledOutputs);
                    options.RefreshSync = Read("refreshSync", options.RefreshSync);
                    options.ChangeSheetName = Read("changeSheetName", options.ChangeSheetName);
                    options.CalendarCtrlDisplay = Read("calendarCtrlDisplay", options.CalendarCtrlDisplay);
                    options.SegmentSelectionWindowDisplay = Read("segmentSelectionWindowDisplay", options.SegmentSelectionWindowDisplay);
                    options.PreserveFormats = Read("preserveFormats", options.PreserveFormats);
                }
                XLEdgePreferencesManager.Instance.ApplyRuntime(options);
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
                        try
                        {
                            // The wait window refuses a plain Close (users cannot close it); RequestClose is its own way out,
                            // which also cancels the operation it is waiting for.
                            if (window is XLEdge.Views.XLEdgeWaitWindow wait) wait.RequestClose();
                            else window.Close();
                        }
                        catch (Exception ex) { Log(ex, "AddinEntry.Shutdown (closing a window)"); }
                    }
                    var stillOpen = app.Windows.Cast<System.Windows.Window>().Select(w => w.GetType().Name).ToList();
                    if (stillOpen.Count > 0) Logger?.LogWarn("AddinEntry.Shutdown: windows still open, closed by the dispatcher shutdown: " + string.Join(", ", stillOpen));
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

        // The pane calls run under Excel's thread context too (GLSense AIPowered "Excel stays running after close"): whatever they
        // start - the pane's WebView2 start-up, its page's reports - resumes on Excel's thread, never on a pool thread.

        public IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle)
        {
            try
            {
                IntPtr content = IntPtr.Zero;
                ExcelThreadContext.Run(() => content = _panes?.Create(excelWindowHandle) ?? IntPtr.Zero);
                return content;
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.CreateTaskPaneContent");
                return IntPtr.Zero;
            }
        }

        public void CloseTaskPaneContent(IntPtr excelWindowHandle)
        {
            try { ExcelThreadContext.Run(() => _panes?.Close(excelWindowHandle)); }
            catch (Exception ex) { Log(ex, "AddinEntry.CloseTaskPaneContent"); }
        }

        public void OnTaskPaneShown(IntPtr excelWindowHandle)
        {
            try { ExcelThreadContext.Run(() => _panes?.Shown(excelWindowHandle)); }
            catch (Exception ex) { Log(ex, "AddinEntry.OnTaskPaneShown"); }
        }

        /// <summary>GLSense logged in and hands its session over: 11.1.2's InvokedFromGLSense.</summary>
        public void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission)
        {
            try
            {
                Logger?.LogDebug("AddinEntry.OnLoginFromGLSense.");
                ExcelThreadContext.Run(() => _module?.InvokedFromGLSense(instanceName, loginUrl, token, userName, hasXLEdgePermission));
            }
            catch (Exception ex) { Log(ex, "AddinEntry.OnLoginFromGLSense"); }
        }

        /// <summary>GLSense logged out: 11.1.2's LogoffFromAddin (XLEdge's own logout sequence), counted as running until it ends.</summary>
        public void OnLogoutFromGLSense()
        {
            try
            {
                Logger?.LogDebug("AddinEntry.OnLogoutFromGLSense.");
                ExcelThreadContext.Run(() => _module?.LogoffFromGLSense());
            }
            catch (Exception ex) { Log(ex, "AddinEntry.OnLogoutFromGLSense"); }
        }

        /// <summary>
        /// Stage 6: the login the shell's Online reload calls the ORBIT server with. Logged in only when the login completed and a
        /// token and URL are there (GLSense CLAUDE.md 44.1: a flag never set kept its Online disabled).
        /// </summary>
        public LoginInfo GetLoginInfo()
        {
            try
            {
                XLEdgeAppState state = XLEdgeAppState.Instance;
                return new LoginInfo
                {
                    LoginUrl = state.LoginUrl,
                    LoginToken = state.LoginToken,
                    LoginUserName = state.LoginUserName,
                    IsLoggedIn = state.IsLoginCompleted && !string.IsNullOrWhiteSpace(state.LoginToken) && !string.IsNullOrWhiteSpace(state.LoginUrl),
                };
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.GetLoginInfo");
                return new LoginInfo();
            }
        }

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

        /// <summary>
        /// Busy - a reload now would unload the AppDomain under running code - while a core dialog's modal loop runs on Excel's
        /// thread, while one of 11.1.2's async handlers has not ended, and while a report runs or refreshes (ProcessRunning).
        /// </summary>
        public bool IsBusy()
        {
            try
            {
                bool dialog = System.Windows.Interop.ComponentDispatcher.IsThreadModal, handlers = ExcelThreadContext.Pending > 0,
                     background = ExcelThreadContext.Background > 0, report = XLEdgeAppState.Instance.ProcessRunning;
                if (dialog || handlers || background || report)
                    Logger?.LogDebug($"AddinEntry.IsBusy: busy (dialog open={dialog}, handlers running={ExcelThreadContext.Pending}, background={ExcelThreadContext.Background}, report running={report}).");
                return dialog || handlers || background || report;
            }
            catch (Exception ex)
            {
                Log(ex, "AddinEntry.IsBusy");
                return false;
            }
        }

        public override object InitializeLifetimeService() => null;

        private void Log(Exception ex, string context)
        {
            try { Logger?.LogException(ex.ToString(), context); }
            catch (Exception) { }
        }
    }
}
