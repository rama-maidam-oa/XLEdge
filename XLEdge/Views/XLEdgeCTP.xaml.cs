using AddinExpress.XL;
using Microsoft.Web.WebView2.Core;
using System.Linq;
using System.Text.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using XLEdge.Helpers;
using XLEdge.Models;
using XLEdge.Utilities;
using System.Web;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeCTP.xaml
    /// </summary>
    public partial class XLEdgeCTP : UserControl
    {
        private const double MinimumConfiguratorWidth = 600;

        private readonly XLEdgeRibbonHelper _ribbonHelper;
        private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);

        private Task _webViewInitTask;
        private readonly ADXExcelTaskPane1 _parentPane;
        private bool _isInitialized;
        private bool _webViewEventsHooked;
        private bool _isDisposed;
        private bool _isHeightCheckScheduled = false;

        private static XLEdgeAppState appState => XLEdgeAppState.Instance;

        public XLEdgeCTP(ADXExcelTaskPane1 parentPane = null)
        {
            InitializeComponent();

            _parentPane = parentPane;
            _ribbonHelper = XLEdgeRibbonHelper.Current;

            if (MainScrollViewer != null)
            {
                MainScrollViewer.MinWidth = MinimumConfiguratorWidth;
            }

            Loaded += OnLoaded;
            Unloaded += OnUnloaded;
            SizeChanged += OnSizeChanged;
            IsVisibleChanged += OnIsVisibleChanged;

            if (WebCtrl != null)
            {
                WebCtrl.Loaded += WebCtrl_Loaded;
            }

            if (_parentPane != null)
            {
                _parentPane.Resize += OnParentPaneResize;
            }
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_parentPane != null)
                {
                    _parentPane.Resize -= OnParentPaneResize;
                }

                if (WebCtrl != null)
                {
                    WebCtrl.Loaded -= WebCtrl_Loaded;
                    WebCtrl.CoreWebView2InitializationCompleted -= WebView_CoreWebView2InitializationCompleted;
                    if (WebCtrl.CoreWebView2 != null)
                    {
                        WebCtrl.CoreWebView2.PermissionRequested -= CoreWebView2_PermissionRequested;
                        WebCtrl.CoreWebView2.ProcessFailed -= CoreWebView2_ProcessFailed;
                        WebCtrl.CoreWebView2.SourceChanged -= WebCtrl_SourceChanged;
                        WebCtrl.CoreWebView2.DocumentTitleChanged -= WebView_DocumentTitleChanged;
                        WebCtrl.CoreWebView2.WebResourceRequested -= WebView_WebResourceRequested;
                    }

                    try
                    {
                        WebCtrl.Dispose();
                    }
                    catch (Exception disposeEx)
                    {
                        LogUtility.LogWarn($"Failed to dispose WebCtrl during unload - {disposeEx.Message}");
                    }
                }

                _isDisposed = true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error during XLEdgeCTP unload");
            }
        }

        private Task RunOnUIAsync(Action action, DispatcherPriority priority = DispatcherPriority.Normal)
        {
            if (_isDisposed)
            {
                return Task.CompletedTask;
            }

            if (Dispatcher.CheckAccess())
            {
                action();
                return Task.CompletedTask;
            }

            return Dispatcher.InvokeAsync(action, priority).Task;
        }

        private Task RunOnUIAsync(Func<Task> func, DispatcherPriority priority = DispatcherPriority.Normal)
        {
            if (_isDisposed)
            {
                return Task.CompletedTask;
            }

            if (Dispatcher.CheckAccess())
            {
                return func();
            }

            return Dispatcher.InvokeAsync(func, priority).Task.Unwrap();
        }

        private Task<T> RunOnUIAsync<T>(Func<T> func, DispatcherPriority priority = DispatcherPriority.Normal)
        {
            if (_isDisposed)
                return Task.FromResult(default(T));

            if (Dispatcher.CheckAccess())
                return Task.FromResult(func());

            return Dispatcher.InvokeAsync(func, priority).Task;
        }

        private Task<T> RunOnUIAsync<T>(Func<Task<T>> func, DispatcherPriority priority = DispatcherPriority.Normal)
        {
            if (_isDisposed)
            {
                return Task.FromResult(default(T));
            }

            if (Dispatcher.CheckAccess())
            {
                return func();
            }

            return Dispatcher.InvokeAsync(func, priority).Task.Unwrap();
        }

        private static void SafeFireAndForget(Func<Task> taskFactory, string context)
        {
            _ = Task.Run(async () =>
            {
                try
                {
                    await taskFactory();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, context);
                }
            });
        }

        private async void WebCtrl_Loaded(object sender, RoutedEventArgs e)
        {
            if (WebCtrl != null)
            {
                WebCtrl.Loaded -= WebCtrl_Loaded;
            }

            using (new LogUtility.LogScope("WebView2 Initialization"))
            {
                try
                {
                    await EnsureWebViewInitializedAsync();
                    await NavigateToLoginUrlAsync();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Failed to initialize WebView2");
                }
            }
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            SafeFireAndForget(async () =>
            {
                await RunOnUIAsync(() =>
                {
                    EnsureMinimumWidth();
                    UpdateLayout();
                    MainScrollViewer?.UpdateLayout();

                    if (WebCtrl != null)
                    {
                        WebCtrl.SizeChanged -= WebCtrl_SizeChanged;
                        WebCtrl.SizeChanged += WebCtrl_SizeChanged;
                    }

                    if (MainScrollViewer != null)
                    {
                        MainScrollViewer.ScrollChanged -= MainScrollViewer_ScrollChanged;
                        MainScrollViewer.ScrollChanged += MainScrollViewer_ScrollChanged;
                    }

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ScheduleHeightCheck();
                    }), DispatcherPriority.Loaded);

                }, DispatcherPriority.Loaded);
            }, "Error in OnLoaded");
        }

        public async Task ReLoadConfigurator()
        {
            try
            {
                await EnsureWebViewInitializedAsync();

                string loginUrl = appState.LoginUrl;
                if (string.IsNullOrWhiteSpace(loginUrl))
                {
                    loginUrl = "about:blank";
                }


                await RunOnUIAsync(() =>
                {
                    if (WebCtrl?.CoreWebView2 != null)
                        WebCtrl.CoreWebView2.Navigate(loginUrl);
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to navigate in ReLoadConfigurator");
            }
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (!IsVisible)
                return;

            SafeFireAndForget(async () =>
            {
                await RunOnUIAsync(() =>
                {
                    EnsureMinimumWidth();
                    UpdateLayout();
                    MainScrollViewer?.UpdateLayout();

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ScheduleHeightCheck();
                    }), DispatcherPriority.Background);
                }, DispatcherPriority.Loaded);
            }, "Error in OnIsVisibleChanged");
        }

        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            SafeFireAndForget(() => RunOnUIAsync(() =>
            {
                EnsureMinimumWidth();
                UpdateLayout();
                MainScrollViewer?.UpdateLayout();
                ScheduleHeightCheck();
            }), "Error in OnSizeChanged");
        }

        private void OnParentPaneResize(object sender, EventArgs e)
        {
            SafeFireAndForget(async () =>
            {
                await RunOnUIAsync(() =>
                {
                    EnsureMinimumWidth();
                    UpdateLayout();
                    MainScrollViewer?.UpdateLayout();

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        ScheduleHeightCheck();
                    }), DispatcherPriority.Background);

                }, DispatcherPriority.Loaded);
            }, "Error in OnParentPaneResize");
        }

        private void EnsureMinimumWidth()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => EnsureMinimumWidth());
                return;
            }

            if (MainScrollViewer != null)
            {
                MainScrollViewer.MinWidth = MinimumConfiguratorWidth;
            }

            if (_parentPane != null && _parentPane.Width < MinimumConfiguratorWidth)
            {
                _parentPane.Width = (int)MinimumConfiguratorWidth;
            }
        }

        // ========== WebView Height Management ==========

        private void WebCtrl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (MainScrollViewer != null && WebCtrl != null && !_isHeightCheckScheduled)
            {
                _isHeightCheckScheduled = true;
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        _isHeightCheckScheduled = false;
                        MainScrollViewer.UpdateLayout();
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"WebCtrl_SizeChanged error: {ex.Message}");
                    }
                }), DispatcherPriority.Background);
            }
        }

        private void MainScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            // Check if viewport height changed significantly
            if (e.ViewportHeight > 0)
            {
                ScheduleHeightCheck();
            }
        }

        private void EnsureWebViewFillsAvailableSpace()
        {
            if (WebCtrl == null || MainScrollViewer == null)
                return;

            try
            {
                double viewportHeight = MainScrollViewer.ViewportHeight;

                if (viewportHeight > 0)
                {
                    double desiredHeight = viewportHeight - 20;

                    if (desiredHeight < 400)
                        desiredHeight = 400;

                    if (Math.Abs(WebCtrl.Height - desiredHeight) > 5)
                    {
                        WebCtrl.Height = desiredHeight;
                        WebCtrl.UpdateLayout();

                        if (WebCtrl.Parent is FrameworkElement parent)
                        {
                            parent.UpdateLayout();
                        }

                        LogUtility.LogDebug($"Adjusted WebCtrl height to {desiredHeight} (viewport: {viewportHeight})");
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogDebug($"EnsureWebViewFillsAvailableSpace error: {ex.Message}");
            }
        }

        private void ScheduleHeightCheck()
        {
            if (_isDisposed || _isHeightCheckScheduled)
                return;

            _isHeightCheckScheduled = true;
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    _isHeightCheckScheduled = false;
                    UpdateLayout();
                    MainScrollViewer?.UpdateLayout();
                    EnsureWebViewFillsAvailableSpace();
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ScheduleHeightCheck error: {ex.Message}");
                }
            }), DispatcherPriority.Loaded);
        }

        // ========== END WebView Height Management ==========

        private async Task EnsureWebViewInitializedAsync()
        {
            if (_isDisposed)
            {
                throw new ObjectDisposedException(nameof(XLEdgeCTP));
            }

            if (_isInitialized && await RunOnUIAsync(() => WebCtrl?.CoreWebView2 != null))
            {
                return;
            }

            await _initLock.WaitAsync();
            try
            {
                if (_isInitialized && await RunOnUIAsync(() => WebCtrl?.CoreWebView2 != null))
                {
                    return;
                }

                if (_webViewInitTask == null)
                {
                    _webViewInitTask = InitializeWebViewInternalAsync();
                }
            }
            finally
            {
                _initLock.Release();
            }

            await _webViewInitTask;
        }

        private async Task InitializeWebViewInternalAsync()
        {
            await RunOnUIAsync(async () =>
            {
                string logDir = XLEdgeAppPaths.BrowserLogsFolder;
                DirectoryInfo di = new DirectoryInfo(logDir);
                if (!di.Exists)
                    di.Create();

                string webViewLogsPath = di.FullName;

                var envOptions = new CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = true
                };

                var env = await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: webViewLogsPath,
                    options: envOptions);

                if (WebCtrl == null)
                    throw new InvalidOperationException("WebCtrl is null.");

                WebCtrl.CoreWebView2InitializationCompleted -= WebView_CoreWebView2InitializationCompleted;
                WebCtrl.CoreWebView2InitializationCompleted += WebView_CoreWebView2InitializationCompleted;

                await WebCtrl.EnsureCoreWebView2Async(env);

                HookWebViewEvents();

                WebCtrl.CoreWebView2.Settings.AreDevToolsEnabled = true;

                var version = WebCtrl.CoreWebView2.Environment.BrowserVersionString;
                LogUtility.LogDebug($"WebView2 BrowserVersion={version}");

                _isInitialized = true;
            });
        }

        private void HookWebViewEvents()
        {
            if (_webViewEventsHooked || WebCtrl?.CoreWebView2 == null)
                return;

            WebCtrl.CoreWebView2.PermissionRequested -= CoreWebView2_PermissionRequested;
            WebCtrl.CoreWebView2.ProcessFailed -= CoreWebView2_ProcessFailed;
            WebCtrl.CoreWebView2.SourceChanged -= WebCtrl_SourceChanged;
            WebCtrl.CoreWebView2.DocumentTitleChanged -= WebView_DocumentTitleChanged;
            WebCtrl.CoreWebView2.WebResourceRequested -= WebView_WebResourceRequested;

            WebCtrl.CoreWebView2.PermissionRequested += CoreWebView2_PermissionRequested;
            WebCtrl.CoreWebView2.ProcessFailed += CoreWebView2_ProcessFailed;
            WebCtrl.CoreWebView2.SourceChanged += WebCtrl_SourceChanged;
            WebCtrl.CoreWebView2.DocumentTitleChanged += WebView_DocumentTitleChanged;
            WebCtrl.CoreWebView2.WebResourceRequested += WebView_WebResourceRequested;

            WebCtrl.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);

            _webViewEventsHooked = true;
        }

        private async Task<bool> IsCoreWebViewReadyAsync()
        {
            return await RunOnUIAsync(() => WebCtrl?.CoreWebView2 != null);
        }

        private async Task NavigateToLoginUrlAsync()
        {
            try
            {
                await EnsureWebViewInitializedAsync();

                string loginUrl = appState.LoginUrl;
                if (string.IsNullOrWhiteSpace(loginUrl))
                {
                    LogUtility.LogWarn("Login URL is empty, cannot navigate");
                    return;
                }

                string urlNavigate = appState.LoginFromGLSense
                    ? $"{loginUrl}/web/public/excel-auth-redirect"
                    : $"{loginUrl}?excel=Y";

                await RunOnUIAsync(() =>
                {
                    if (WebCtrl == null)
                        return;

                    WebCtrl.Visibility = Visibility.Visible;
                    WebCtrl.Source = new Uri(urlNavigate);

                    SetPaneCaption(appState.LoginUrl);
                });

                LogUtility.LogDebug($"Navigating to: {urlNavigate}");
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error navigating to login URL");
            }
        }

        public async Task<bool> LogoutSessionAsync(string loginUrl, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(loginUrl))
            {
                LogUtility.LogWarn("LogoutSessionAsync skipped because loginUrl is empty.");
                return false;
            }

            return await RunOnUIAsync(async () =>
            {
                if (WebCtrl == null || WebCtrl.CoreWebView2 == null)
                {
                    LogUtility.LogWarn("WebCtrl/CoreWebView2 is not ready for logout - skipping (nothing to log out of).");
                    return false;
                }

                string navUrl = $"{loginUrl.TrimEnd('/')}/web/secure/applogout";
                var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                EventHandler<CoreWebView2NavigationCompletedEventArgs> handler = null;

                handler = (s, e) =>
                {
                    try
                    {
                        WebCtrl.CoreWebView2.NavigationCompleted -= handler;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogDebug($"{nameof(LogoutSessionAsync)}: failed to unsubscribe NavigationCompleted handler - {ex.Message}");
                    }

                    tcs.TrySetResult(e.IsSuccess);
                };

                try
                {
                    token.ThrowIfCancellationRequested();

                    WebCtrl.CoreWebView2.NavigationCompleted += handler;
                    WebCtrl.Source = new Uri(navUrl);

                    Task completed = await Task.WhenAny(tcs.Task, Task.Delay(TimeSpan.FromSeconds(10), token));
                    if (completed != tcs.Task)
                    {
                        try
                        {
                            WebCtrl.CoreWebView2.NavigationCompleted -= handler;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(LogoutSessionAsync)}: failed to unsubscribe NavigationCompleted handler after timeout - {ex.Message}");
                        }

                        LogUtility.LogWarn("Logout navigation timeout or cancelled.");
                        token.ThrowIfCancellationRequested();
                        return false;
                    }

                    return await tcs.Task;
                }
                catch (OperationCanceledException)
                {
                    try
                    {
                        WebCtrl.CoreWebView2.NavigationCompleted -= handler;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"Logout navigation cancelled; also failed to unsubscribe NavigationCompleted handler - {ex.Message}");
                    }

                    throw;
                }
                catch (Exception ex)
                {
                    try
                    {
                        WebCtrl.CoreWebView2.NavigationCompleted -= handler;
                    }
                    catch (Exception unsubEx)
                    {
                        LogUtility.LogDebug($"{nameof(LogoutSessionAsync)}: failed to unsubscribe NavigationCompleted handler after exception - {unsubEx.Message}");
                    }

                    LogUtility.LogException(ex, "Exception during logout navigation.");
                    return false;
                }
            });
        }

        public async Task ExecuteScriptAsync(string script)
        {
            if (string.IsNullOrWhiteSpace(script))
            {
                return;
            }

            await EnsureWebViewInitializedAsync();

            await RunOnUIAsync(async () =>
            {
                if (WebCtrl?.CoreWebView2 == null)
                {
                    LogUtility.LogWarn("ExecuteScriptAsync skipped - WebCtrl/CoreWebView2 is not ready.");
                    return;
                }

                try
                {
                    await WebCtrl.CoreWebView2.ExecuteScriptAsync(script);
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "ExecuteScriptAsync failed.");
                }
            });
        }

        private void CoreWebView2_PermissionRequested(object sender, CoreWebView2PermissionRequestedEventArgs e)
        {
            using (new LogUtility.LogScope("CoreWebView2_PermissionRequested"))
            {
                try
                {
                    LogUtility.LogDebug($"Permission requested: Kind={e.PermissionKind}, Uri={e.Uri}");

                    switch (e.PermissionKind)
                    {
                        case CoreWebView2PermissionKind.Microphone:
                        case CoreWebView2PermissionKind.Camera:
                        case CoreWebView2PermissionKind.Geolocation:
                        case CoreWebView2PermissionKind.MidiSystemExclusiveMessages:
                        case CoreWebView2PermissionKind.ClipboardRead:
                            e.State = CoreWebView2PermissionState.Allow;
                            e.Handled = true;
                            break;

                        default:
                            e.State = CoreWebView2PermissionState.Deny;
                            e.Handled = true;
                            LogUtility.LogWarn($"Permission denied: {e.PermissionKind}");
                            break;
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "PermissionRequested handler error");
                    e.State = CoreWebView2PermissionState.Deny;
                    e.Handled = true;
                }
            }
        }

        private void CoreWebView2_ProcessFailed(object sender, CoreWebView2ProcessFailedEventArgs e)
        {
            LogUtility.LogWarn($"WebView2 process failed. Kind={e.ProcessFailedKind}");
        }

        private async void WebCtrl_SourceChanged(object sender, CoreWebView2SourceChangedEventArgs e)
        {
            try
            {
                await RunOnUIAsync(async () =>
                {
                    using (new LogUtility.LogScope("WebCtrl_SourceChanged"))
                    {
                        string sourceUrl = WebCtrl?.Source?.ToString() ?? string.Empty;

                        ScheduleHeightCheck();

                        if (sourceUrl.Contains("excel=Y#Home"))
                        {
                            appState.IsLoginCompleted = true;

                            HandleExcelHomeSource();
                            await ProcessCookiesAsync();
                            UpdateDialogLauncherState();

                            _ribbonHelper.SetControlCaption("RibEdgeLogout", appState.LoginUrlName);

                            try
                            {
                                await ProcessBroadcastMessagesAsync();
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogException(ex, "Error in ProcessBroadcastMessagesAsync");
                            }

                            SyncLoginToGLSense();
                            UpdateExcelTabLabel();

                            await Task.Delay(300);
                            ScheduleHeightCheck();

                            await Task.Delay(700);
                            ScheduleHeightCheck();

                            try
                            {
                                await ReportGenerator.ReleaseKeyboardFocusFromTaskPaneAsync();
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogException(ex, "Failed to release focus to Excel after login");
                            }

                            return;
                        }

                        if (sourceUrl.Contains("loggedout=true") || sourceUrl.Contains("applogout"))
                        {
                            SetPaneCaption(string.Empty);

                            appState.IsLoginCompleted = false;

                            _ribbonHelper.SetControlCaption("RibEdgeLogin", "Login");

                            if (WebCtrl != null)
                                WebCtrl.Source = new Uri("about:blank");

                            ScheduleHeightCheck();
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in WebCtrl_SourceChanged");
            }
        }

        private void HandleExcelHomeSource()
        {
            _ribbonHelper.ApplyState("LoggedIn");
        }

        private async Task ProcessCookiesAsync()
        {
            if (!string.IsNullOrEmpty(appState.LoginToken) || appState.LoginFromGLSense)
                return;

            try
            {
                await EnsureWebViewInitializedAsync();

                string currentUrl = await RunOnUIAsync(() => WebCtrl?.Source?.ToString() ?? string.Empty);
                if (string.IsNullOrWhiteSpace(currentUrl))
                    return;

                List<CoreWebView2Cookie> cookies = await RunOnUIAsync(async () =>
                    await WebCtrl.CoreWebView2.CookieManager.GetCookiesAsync(currentUrl));

                for (int i = 0; i < cookies.Count; i++)
                {
                    if (string.IsNullOrEmpty(cookies[i].Name))
                        continue;

                    string upperCookieName = cookies[i].Name.ToUpperInvariant();

                    if (upperCookieName == "XL-AUTH-TOKEN" || upperCookieName == "ORB-AUTH-TOKEN")
                    {
                        appState.LoginToken = cookies[i].Value;
                    }
                    else if (upperCookieName == "X-ORB-USERNAME")
                    {
                        appState.LoginUserName = HttpUtility.UrlDecode(cookies[i].Value) ?? string.Empty;
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in ProcessCookiesAsync");
            }
        }

        private void UpdateDialogLauncherState()
        {
            bool enableLauncher = !string.IsNullOrEmpty(appState.LoginUserName);
            _ribbonHelper.SetControlEnabled("RibEdgeDialogBoxLauncher", enableLauncher);
        }

        private async Task ProcessBroadcastMessagesAsync()
        {
            if (string.IsNullOrEmpty(appState.LoginUrl) ||
                _ribbonHelper.GetControlCaption("RibEdgeLogin") == appState.LoginUrl.Trim() ||
                string.IsNullOrWhiteSpace(appState.LoginToken) ||
                appState.LoginFromGLSense)
            {
                return;
            }

            using var cts = new CancellationHelper();

            string apiUrl = appState.LoginUrl.Trim() + "/web/secure/get-broadcast-msg";
            var broadcastMsg = await BroadcastMessageFromApi(apiUrl, cts.GetToken());

            if (AppOverlayControl != null)
            {
                await RunOnUIAsync(async () =>
                {
                    await AppOverlayControl.HideBusyAsync();

                    if (!string.IsNullOrWhiteSpace(broadcastMsg))
                        await AppOverlayControl.ShowInfoAsync(broadcastMsg);
                });
            }
        }

        private static async Task<string> BroadcastMessageFromApi(string apiUrl, CancellationToken ct)
        {
            try
            {
                string rawResponse = await FetchApiResponseAsync(apiUrl, ct);

                if (string.IsNullOrWhiteSpace(rawResponse))
                    return string.Empty;

                return FormatBroadcastMessages(rawResponse);
            }
            catch (OperationCanceledException ex)
            {
                LogUtility.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                return string.Empty;
            }
        }

        private static async Task<string> FetchApiResponseAsync(string apiUrl, CancellationToken ct)
        {
            try
            {
                return await ApiHelper.ServerAPI(apiUrl, "Form", "", "POST", ct);
            }
            catch (OperationCanceledException ex)
            {
                LogUtility.LogWarn(ex.Message);
                return string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"Error fetching broadcast messages from {apiUrl}: {ex.Message}");
                return string.Empty;
            }
        }

        private static string FormatBroadcastMessages(string rawResponse)
        {
            try
            {
                var result = ApiResponseHelper.Parse<List<BroadcastMessage>>(rawResponse, JsonGlobals.Options);

                if (!result.IsSuccess)
                {
                    LogUtility.LogWarn($"Broadcast parsing failed: {result.ErrorMessage}");
                    return string.Empty;
                }

                if (result.Value == null || result.Value.Count == 0)
                    return string.Empty;

                return BuildMessageString(result.Value);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error processing broadcast messages.");
                return string.Empty;
            }
        }

        private static string BuildMessageString(List<BroadcastMessage> messages)
        {
            if (messages == null || messages.Count == 0)
                return string.Empty;

            var sb = new StringBuilder();

            for (int i = 0; i < messages.Count; i++)
            {
                var msg = messages[i];

                sb.Append(i + 1)
                  .Append(".) ")
                  .Append(msg.MsgType ?? "Info")
                  .Append(" : ")
                  .Append(msg.Message ?? string.Empty);

                if (i < messages.Count - 1)
                    sb.AppendLine();
            }

            return sb.ToString();
        }

        private void SyncLoginToGLSense()
        {
            try
            {
                XLEdge.AddinModule.CurrentInstance?.NotifyGLSenseOfLogin(appState.LoginToken, appState.LoginUrl, appState.LoginUserName);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in SyncLoginToGLSense");
            }
        }

        /// <summary>
        /// Updates the Excel tab label when login state changes
        /// </summary>
        private static void UpdateExcelTabLabel()
        {
            // Intentionally left as-is - placeholder for future implementation
            // This method is called from WebCtrl_SourceChanged when login completes
            // Currently no action needed, but kept for extensibility
            try
            {
              ProgressCoordinator.UpdateRibbonLoginStatus();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(UpdateExcelTabLabel));
            }
        }

        private async void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            try
            {
                if (!e.IsSuccess)
                    return;

                await RunOnUIAsync(() =>
                {
                    if (WebCtrl?.CoreWebView2 == null)
                        return;

                    HookWebViewEvents();

                    ADXExcelTaskPane1 edgeExcelPane = XLEdge.AddinModule.CurrentInstance.GetPaneInstance();
                    string loginUrl = appState.LoginUrl;

                    if (edgeExcelPane != null && edgeExcelPane.Visible)
                    {
                        if (appState.LoginFromGLSense)
                        {
                            WebCtrl.Source = new Uri(loginUrl + "/web/public/excel-auth-redirect");
                        }
                        else if (!string.IsNullOrEmpty(loginUrl))
                        {
                            WebCtrl.Source = new Uri(loginUrl + "?excel=Y");
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in WebView_CoreWebView2InitializationCompleted");
            }
        }

        private void WebView_WebResourceRequested(object sender, CoreWebView2WebResourceRequestedEventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(appState.LoginToken))
                    return;

                e.Request.Headers.SetHeader("Authorization", "Bearer " + appState.LoginToken);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in WebResourceRequested handler");
                LogUtility.LogDebug($"Web resource requested: {e.Request.Uri}");
            }
        }

        private async void WebView_DocumentTitleChanged(object sender, object e)
        {
            try
            {
                string title = WebCtrl?.CoreWebView2?.DocumentTitle ?? string.Empty;
                LogUtility.LogDebug($"Document title changed: {title}");

                ScheduleHeightCheck();

                if (string.IsNullOrWhiteSpace(title) || !title.Contains("|"))
                {
                    return;
                }

                string[] parts = title.Split('|');
                string command = parts.Length > 0 ? parts[0] : string.Empty;

                switch (command)
                {
                    case "EdgeWorkbook":
                        {
                            string runIds = await FetchWorkbookRerunIdsAsync();
                            if (!string.IsNullOrWhiteSpace(runIds))
                            {
                                SafeFireAndForget(
                                    () => XLEdge.Helpers.ReportGenerator.CreateMultiDataReportsAsync(runIds, AppOverlayControl),
                                    "Handle EdgeWorkbook (MultiData) document title change");
                            }
                            break;
                        }

                    case "Process":
                    case "Edge":
                        if (parts.Length >= 5 && parts[4] == "XLSX")
                        {
                            string processId = parts.Length > 1 ? parts[1] : string.Empty;
                            string downloadUrl = $"{XLEdgeAppState.Instance.LoginUrl?.TrimEnd('/')}/rest/secure/process/finance-report-output?processId={processId}";
                            SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.DownloadFile1Async(downloadUrl), "Handle Process+XLSX file download document title change");
                        }
                        else
                        {
                            SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.CreateReportFromTitleAsync(title, AppOverlayControl), "Handle document title change");
                        }
                        break;

                    case "Logs":
                        SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.CreateLogsReportAsync(title, AppOverlayControl), "Handle Logs document title change");
                        break;

                    case "Excel":
                        {
                            string reportId = parts.Length > 1 ? parts[1] : string.Empty;
                            string downloadUrl = $"{XLEdgeAppState.Instance.LoginUrl?.TrimEnd('/')}/web/secure/financeTemplateFileDownload?reportId={reportId}";
                            SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.DownloadFile1Async(downloadUrl), "Handle Excel file download document title change");
                            break;
                        }

                    default:
                        LogUtility.LogDebug($"WebView_DocumentTitleChanged: unrecognized command '{command}' - ignored. Title: {title}");
                        break;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in WebView_DocumentTitleChanged");
            }
        }

        private async Task<string> FetchWorkbookRerunIdsAsync()
        {
            const string script = @"(() => {
                let element = document.querySelector('[reruntype=xledgeworkbookrerun]');
                return element ? element.getAttribute('newrunids') : null;
                })()";

            try
            {
                await EnsureWebViewInitializedAsync();

                string result = await RunOnUIAsync(async () =>
                {
                    if (WebCtrl?.CoreWebView2 == null)
                    {
                        return null;
                    }

                    return await WebCtrl.CoreWebView2.ExecuteScriptAsync(script);
                });

                LogUtility.LogDebug($"FetchWorkbookRerunIdsAsync result: {result}");

                if (string.IsNullOrWhiteSpace(result) || result == "null")
                {
                    return null;
                }

                return result.Trim('"');
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in FetchWorkbookRerunIdsAsync");
                return null;
            }
        }

        private void ADXExcelTaskPane1_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            SafeFireAndForget(async () =>
            {
                try
                {
                    await RunOnUIAsync(() =>
                    {
                        if (Width != 600)
                            Width = 600;

                        appState.EdgePaneShown = false;

                        if (!string.IsNullOrWhiteSpace(appState.LoginUrl))
                            SetPaneCaption(appState.LoginUrl);
                    });

                    var timeout = TimeSpan.FromMinutes(1);
                    var start = DateTime.UtcNow;

                    while (DateTime.UtcNow - start < timeout)
                    {
                        bool ready = await IsCoreWebViewReadyAsync();
                        if (ready)
                        {
                            await NavigateToLoginUrlSafeAsync();
                            return;
                        }

                        await Task.Delay(500);
                    }

                    LogUtility.LogWarn("WebView2 initialization timed out after 1 minute.");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Error in AfterTaskPaneShow");
                }
            }, "Unhandled error in ADXAfterTaskPaneShow");
        }

        internal async Task NavigateToLoginUrlSafeAsync()
        {
            try
            {
                await EnsureWebViewInitializedAsync();

                string urlNavigate = string.Empty;

                if (appState.LoginFromGLSense)
                {
                    urlNavigate = appState.LoginUrl + "/web/public/excel-auth-redirect";
                }
                else
                {
                    bool showLogin = XLEdge.AddinModule.CurrentInstance.loginButtonVisibility();
                    string currentSource = await RunOnUIAsync(() => WebCtrl?.Source?.ToString() ?? string.Empty);

                    if (showLogin)
                    {
                        urlNavigate = appState.LoginUrl + "?excel=Y";
                    }
                    else if (!currentSource.Contains("excel=Y#Home") && !string.IsNullOrEmpty(appState.LoginToken))
                    {
                        urlNavigate = appState.LoginUrl + "/web/public/excel-auth-redirect";
                    }
                }

                if (!string.IsNullOrEmpty(urlNavigate))
                {
                    await RunOnUIAsync(() =>
                    {
                        WebCtrl.Visibility = Visibility.Visible;
                        WebCtrl.Source = new Uri(urlNavigate);
                    });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error navigating to login URL");
            }
        }

        internal async Task NavigateBlankAsync()
        {
            try
            {
                await RunOnUIAsync(() =>
                {
                    if (WebCtrl != null)
                    {
                        WebCtrl.Source = new Uri("about:blank");
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error navigating WebView to about:blank");
            }
        }

        internal async Task RefreshLoginNavigationAsync()
        {
            await RunOnUIAsync(() =>
            {
                if (!string.IsNullOrWhiteSpace(appState.LoginUrl))
                    SetPaneCaption(appState.LoginUrl);
            });

            await NavigateToLoginUrlSafeAsync();
        }

        private void SetPaneCaption(string caption)
        {
            try
            {
                if (_parentPane != null)
                {
                    _parentPane.Text = caption;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"SetPaneCaption: failed to set task pane caption - {ex.Message}");
            }
        }

        public void ReleaseFocusToExcel()
        {
            try
            {
                Dispatcher.Invoke(() =>
                {
                    if (WebCtrl != null && WebCtrl.CoreWebView2 != null)
                    {
                        try
                        {
                            WebCtrl.CoreWebView2.ExecuteScriptAsync("document.activeElement?.blur();");
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogError($"ReleaseFocusToExcel: WebView2 blur failed - {ex.Message}");
                        }

                        if (_parentPane != null)
                        {
                            _parentPane.Focus();
                        }

                        try
                        {
                            var excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                            if (excelApp != null)
                            {
                                ExcelWindowHelper.ActivateExcelMainWindow(excelApp);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogError($"ReleaseFocusToExcel: COM focus failed - {ex.Message}");
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ReleaseFocusToExcel failed");
            }
        }

        /// <summary>
        /// Public method to manually refresh WebView height
        /// </summary>
        public void RefreshWebViewHeight()
        {
            if (_isDisposed)
                return;

            Dispatcher.BeginInvoke(new Action(() =>
            {
                UpdateLayout();
                MainScrollViewer?.UpdateLayout();
                ScheduleHeightCheck();
            }), DispatcherPriority.Loaded);
        }
    }
}