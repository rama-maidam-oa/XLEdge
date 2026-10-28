using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Windows.Threading;
using AddinExpress.MSO;
using XLEdge.Hosting;
using XLEdge.Loader.Core;
using XLEdge.Views;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLEdge
{
    /// <summary>
    /// The XLEdge add-in as Excel sees it: ribbon, Excel events and task panes. It holds no XLEdge logic: every click and event
    /// is forwarded to the core, which runs in its own AppDomain and can be reloaded without restarting Excel (Reload Add-in /
    /// Version History). Only primitives cross to the core, and nothing thrown here or by the core reaches Add-in Express.
    /// </summary>
    [Guid("80B0FB76-A5F4-41E0-B283-C175B7374B60"), ProgId("XLEdge.AddinModule")]
    public partial class AddinModule : ADXAddinModule
    {
        private const int TaskPaneWidth = 500;

        private readonly UpdateBootstrapper _bootstrapper = new UpdateBootstrapper();
        private readonly WpfHostUi _hostUi;
        private XLEdgeContext _context;
        private RibbonController _ribbonController;
        private bool _hostWindowOpen;
        private bool _shutDown;

        public AddinModule()
        {
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            System.Windows.Forms.Application.EnableVisualStyles();
            _hostUi = new WpfHostUi(ReadExcelHandle);
            InitializeComponent();
        }

        #region Add-in Express automatic code

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance => ADXAddinModule.CurrentInstance as AddinModule;

        /// <summary>Starts, reloads and shuts down the core; null before the ribbon has loaded.</summary>
        internal CoreHost Host { get; private set; }

        // ---- calls from the GLSense add-in (COMAddIns("XLEdge.AddinModule").Object) -----------

        /// <summary>GLSense logged in and shares its session with XLEdge.</summary>
        public void InvokedFromGLSense(string eeName, string eeUrl, string eeToken, string eeUser, bool hasXLEdgePermission = true)
        {
            try
            {
                if (Host?.Addin == null) { _context?.Logger.LogWarn("AddinModule.InvokedFromGLSense: no XLEdge core is loaded; ignored."); return; }
                Host.Addin.OnLoginFromGLSense(eeName, eeUrl, eeToken, eeUser, hasXLEdgePermission);
            }
            catch (Exception ex)
            {
                LogError("AddinModule.InvokedFromGLSense failed.", ex);
            }
        }

        /// <summary>GLSense logged out and asks XLEdge to log out too.</summary>
        public void LogoffFromAddin()
        {
            try
            {
                if (Host?.Addin == null) { _context?.Logger.LogWarn("AddinModule.LogoffFromAddin: no XLEdge core is loaded; ignored."); return; }
                Host.Addin.OnLogoutFromGLSense();
            }
            catch (Exception ex)
            {
                LogError("AddinModule.LogoffFromAddin failed.", ex);
            }
        }

        // ---- lifecycle ----------------------------------------------------------------------

        private void AddinModule_OnRibbonLoaded(object sender, IRibbonUI ribbon)
        {
            // Office raises this once per ribbon load; the core is started only once per Excel session.
            if (Host != null) return;
            try
            {
                string installRoot = Path.GetDirectoryName(typeof(AddinModule).Assembly.Location);
                _context = new XLEdgeContext(HostApplication, installRoot);
                ReadExcelHandle();
                _context.Logger.LogInfo($"AddinModule: XLEdge {_context.ShellVersion} started (install folder '{installRoot}').");

                _ribbonController = new RibbonController(this, new ShellCommands(this, ribbon, Dispatcher.CurrentDispatcher), _context.Logger);
                _context.SetRibbonController(_ribbonController);
                _ribbonController.SetState("LoggedOut");

                Host = new CoreHost(_context, new DomainCoreLoader(new AddinDomainLoader()), new BootstrapperReleaseCatalog(_bootstrapper), _hostUi, new PaneHost(this));
                Host.Start();
            }
            catch (Exception ex)
            {
                LogError("AddinModule.OnRibbonLoaded failed.", ex);
                ShowMessage($"XLEdge could not start: {ex.Message}", "XLEdge");
            }
        }

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            _context?.Logger.LogDebug("AddinModule: Excel finished starting the add-in.");
        }

        private void AddinModule_AddinBeginShutdown(object sender, EventArgs e)
        {
            if (_shutDown) return;
            _shutDown = true;
            try
            {
                _context?.Logger.LogDebug("AddinModule: Excel is closing - shutting the XLEdge core down.");
                _context?.Logger.FlushDebugLogs("add-in shutting down");
                UnsubscribeExcelEvents();
                Host?.Shutdown();
                _context?.Logger.LogDebug("AddinModule: teardown complete.");
            }
            catch (Exception ex)
            {
                LogError("AddinModule.AddinBeginShutdown failed.", ex);
            }
            finally
            {
                try { _context?.Dispose(); } catch (Exception) { }
            }
        }

        private void AddinModule_AddinFinalize(object sender, EventArgs e)
        {
            try { UnsubscribeExcelEvents(); }
            catch (Exception) { }
            AppDomain.CurrentDomain.UnhandledException -= OnUnhandledException;
        }

        private void AddinModule_OnError(ADXErrorEventArgs e)
        {
            e.Handled = true;
            LogError("Add-in Express error: " + e.ADXError, null);
            ShowMessage("XLEdge error: " + e.ADXError, "XLEdge");
        }

        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            try
            {
                _context?.Logger.LogException(Convert.ToString(e.ExceptionObject), "AppDomain.UnhandledException (shell)");
                _context?.Logger.FlushDebugLogs("unhandled exception");
            }
            catch (Exception)
            {
            }
        }

        // ---- ribbon -------------------------------------------------------------------------

        private void RibLogin_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeLogin);
        private void RibEdgeLogout_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeLogout);
        private void RibEdgeLogout_PropertyChanging(object sender, ADXRibbonPropertyChangingEventArgs e)
        {
            // Today's add-in updates the sheet label only when Logout is shown or hidden (the login state changed).
            if (e.PropertyType == ADXRibbonControlPropertyType.Visible) Forward("RibEdgeLogout_PropertyChanging", e.PropertyType.ToString());
        }
        private void RibEdgeDialogBoxLauncher_OnAction(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeDialogBoxLauncher);
        private void RibEdgeRefresh_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeRefresh);
        private void RibEdgeRefreshAll_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeRefreshAll);
        private void RibEdgeParamRefresh_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeParamRefresh);
        private void RibEdgeParamRefreshBook_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeParamRefreshBook);
        private void RibEdgeShowHide_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeShowHide);
        private void RibEdgeIncludeOutputData_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeIncludeOutputData, pressed);
        private void RibEdgeOptions_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeOptions);
        private void RibControlSheet_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibControlSheet);
        private void RibEdgeAbout_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeAbout);
        private void RibEdgeHelp_OnClick(object sender, IRibbonControl control, bool pressed) => Forward(RibbonControlIds.RibEdgeHelp);

        /// <summary>The Debug toggle's state as the shell's logger sees it (also set when the core resets the ribbon).</summary>
        internal void SetDebugMode(bool enabled)
        {
            if (_context != null) _context.DebugMode = enabled;
        }

        private void RibEdgeDebug_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            SetDebugMode(pressed);
            Forward(RibbonControlIds.RibEdgeDebug, pressed);
        }

        private void RibReload_OnClick(object sender, IRibbonControl control, bool pressed) => ShowReloadWindow();
        private void RibVersionHistory_OnClick(object sender, IRibbonControl control, bool pressed) => ShowVersionHistory();

        /// <summary>IRibbonController.ExecuteAction: run a ribbon control's action as if it was clicked.</summary>
        internal void RunRibbonAction(string controlId)
        {
            switch (controlId)
            {
                case RibbonControlIds.RibReload: ShowReloadWindow(); break;
                case RibbonControlIds.RibVersionHistory: ShowVersionHistory(); break;
                default: Forward(controlId); break;
            }
        }

        private void Forward(string action, object parameter = null)
        {
            try
            {
                _context?.Logger.LogDebug($"AddinModule: ribbon '{action}'.");
                Host?.Addin?.OnRibbonAction(action, parameter);
            }
            catch (Exception ex)
            {
                LogError($"AddinModule: ribbon action '{action}' failed.", ex);
            }
        }

        // ---- hot reload windows -------------------------------------------------------------

        private void ShowReloadWindow()
        {
            if (_hostWindowOpen || Host == null) return;
            _hostWindowOpen = true;
            try
            {
                XLEdgeReloadSourcePicker picker;
                try
                {
                    picker = new XLEdgeReloadSourcePicker(_context);
                    new WindowInteropHelper(picker).Owner = ReadExcelHandle();
                    if (picker.ShowDialog() != true || picker.ScanResult == null) return;
                }
                catch (Exception ex)
                {
                    LogError("AddinModule: the reload window could not be shown.", ex);
                    ShowMessage($"Could not open the reload window: {ex.Message}", "Reload XLEdge Add-in");
                    return;
                }
                string source = picker.SelectedSource ?? "Offline";
                OfflineScanResult release = picker.ScanResult;
                // Staged only now that the reload goes ahead (not refused, not cancelled) and the old core is unloaded.
                Host.Reload(() =>
                {
                    OfflineReleaseScanner.Stage(release, _context.Paths);
                    return _bootstrapper.ResolveVersionToLoad(_context, source);
                });
            }
            finally
            {
                _hostWindowOpen = false;
            }
        }

        private void ShowVersionHistory()
        {
            if (_hostWindowOpen || Host == null) return;
            _hostWindowOpen = true;
            try
            {
                ResolvedRelease chosen;
                try
                {
                    var browser = new XLEdgeReleaseHistoryBrowser(_context);
                    new WindowInteropHelper(browser).Owner = ReadExcelHandle();
                    if (browser.ShowDialog() != true || browser.Chosen == null) return;
                    chosen = browser.Chosen;
                }
                catch (Exception ex)
                {
                    LogError("AddinModule: the version history window could not be shown.", ex);
                    ShowMessage($"Could not open the version history: {ex.Message}", "XLEdge Version History");
                    return;
                }
                Host.Reload(() => chosen);
            }
            finally
            {
                _hostWindowOpen = false;
            }
        }

        // ---- task panes (used by RibbonController and CoreHost) -------------------------------

        internal XLEdgeTaskPane ActivePane
        {
            get
            {
                try { return adxExcelTaskPanesCollectionItem1.TaskPaneInstance as XLEdgeTaskPane; }
                catch (Exception) { return null; }
            }
        }

        internal IEnumerable<XLEdgeTaskPane> AllPanes()
        {
            var panes = new List<XLEdgeTaskPane>();
            try
            {
                foreach (object pane in adxExcelTaskPanesCollectionItem1.TaskPaneInstances)
                    if (pane is XLEdgeTaskPane p) panes.Add(p);
            }
            catch (Exception ex)
            {
                LogError("AddinModule: could not list the task panes.", ex);
            }
            return panes;
        }

        internal void ShowTaskPane()
        {
            XLEdgeTaskPane.ShowRequested = true;
            try
            {
                XLEdgeTaskPane pane = ActivePane;
                if (pane == null)
                {
                    adxExcelTaskPanesCollectionItem1.Position = AddinExpress.XL.ADXExcelTaskPanePosition.Right;
                    pane = adxExcelTaskPanesCollectionItem1.CreateTaskPaneInstance() as XLEdgeTaskPane;
                    if (pane == null) return;
                    pane.Width = TaskPaneWidth;
                }
                pane.Show();
                pane.Visible = true;
            }
            finally
            {
                // A pane that did not show must not let Excel restore it later on its own.
                XLEdgeTaskPane.ShowRequested = false;
            }
        }

        // ---- Excel events: primitives only ----------------------------------------------------

        private void adxExcelAppEvents1_SheetSelectionChange(object sender, object sheet, object range)
        {
            // A click on a cell while the keyboard is in a pane's WebView2 must give the keyboard back to the worksheet.
            try
            {
                if (PaneFocus.ReturnToGridAfterClick(AllPanes().Where(p => p.IsHandleCreated).Select(p => p.Handle)))
                    _context?.Logger.LogDebug("AddinModule: keyboard focus returned from the task pane to the worksheet.");
            }
            catch (Exception ex)
            {
                LogError("AddinModule: returning the keyboard focus to the worksheet failed.", ex);
            }
            var target = range as Excel.Range;
            ForwardEvent("SheetSelectionChange", () => new object[] { SheetName(sheet), Address(target), CellCount(target) });
        }

        private void adxExcelAppEvents1_SheetActivate(object sender, object hostObj) =>
            ForwardEvent("SheetActivate", () => new object[] { SheetName(hostObj) });

        private void adxExcelAppEvents1_WorkbookActivate(object sender, object hostObj) =>
            ForwardEvent("WorkbookActivate", () => new object[] { (hostObj as Excel.Workbook)?.Name });

        private void adxExcelAppEvents1_SheetFollowHyperlink(object sender, object sheet, object hyperlink)
        {
            ForwardEvent("SheetFollowHyperlink", () =>
            {
                var link = hyperlink as Excel.Hyperlink;
                string cell = null;
                Excel.Range range = null;
                try
                {
                    range = link?.Range;
                    cell = Address(range);
                }
                catch (COMException) { }
                finally
                {
                    if (range != null) Marshal.ReleaseComObject(range);
                }
                return new object[] { SheetName(sheet), cell, Safe(() => link?.ScreenTip), Safe(() => link?.SubAddress), Safe(() => link?.TextToDisplay) };
            });
        }

        private void AdxExcelAppEvents1_SheetBeforeDelete(object sender, object sheet) =>
            ForwardEvent("SheetBeforeDelete", () => new object[] { SheetName(sheet) });

        private void ForwardEvent(string eventName, Func<object[]> args)
        {
            try
            {
                var addin = Host?.Addin;
                if (addin == null) return;
                addin.OnExcelEvent(eventName, args());
            }
            catch (Exception ex)
            {
                LogError($"AddinModule: Excel event '{eventName}' failed.", ex);
            }
        }

        private void UnsubscribeExcelEvents()
        {
            if (adxExcelAppEvents1 == null) return;
            adxExcelAppEvents1.SheetSelectionChange -= adxExcelAppEvents1_SheetSelectionChange;
            adxExcelAppEvents1.SheetActivate -= adxExcelAppEvents1_SheetActivate;
            adxExcelAppEvents1.WorkbookActivate -= adxExcelAppEvents1_WorkbookActivate;
            adxExcelAppEvents1.SheetFollowHyperlink -= adxExcelAppEvents1_SheetFollowHyperlink;
            adxExcelAppEvents1.SheetBeforeDelete -= AdxExcelAppEvents1_SheetBeforeDelete;
        }

        private static string SheetName(object sheet) => Safe(() => (sheet as Excel.Worksheet)?.Name);
        private static string Address(Excel.Range range) => Safe(() => range?.Address[false, false]);
        private static long CellCount(Excel.Range range) { try { return range == null ? 0 : Convert.ToInt64(range.CountLarge); } catch (COMException) { return 0; } }

        private static string Safe(Func<string> read)
        {
            try { return read(); }
            catch (COMException) { return null; }
        }

        // ---- helpers ------------------------------------------------------------------------

        internal void LogError(string message, Exception ex)
        {
            try { _context?.Logger.LogError(message, ex?.ToString()); }
            catch (Exception) { }
        }

        /// <summary>An error box owned by the Excel window (never throws).</summary>
        private void ShowMessage(string message, string title) => _hostUi.ShowError(message, title);

        /// <summary>Excel's main window, read fresh each time (Excel can recreate it); IntPtr.Zero when it cannot be read.</summary>
        private IntPtr ReadExcelHandle()
        {
            try
            {
                var handle = new IntPtr((HostApplication as Excel._Application).Hwnd);
                if (_context != null) _context.ExcelHandle = handle;
                return handle;
            }
            catch (Exception ex)
            {
                _context?.Logger.LogWarn("AddinModule: could not read Excel's window handle: " + ex.Message);
                return IntPtr.Zero;
            }
        }
    }
}
