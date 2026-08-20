using AddinExpress.XL;
using Microsoft.Web.WebView2.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using XLEdge.Helpers;
using XLEdge.Models;
using XLEdge.Utilities;

namespace XLEdge
{
#nullable enable
    // Ported from ADXExcelTaskPane1.vb: hosts WebView2 (WebCtrl, added via the Designer - see
    // ADXExcelTaskPane1.Designer.cs) directly as a native WinForms control, the same way VB.NET does
    // it - no WPF, no ElementHost, no DPI-awareness or resize-handling code of any kind, matching
    // VB.NET's actual pane exactly (confirmed by reading ADXExcelTaskPane1.vb directly: it has no
    // Resize/ResizeBegin/SizeChanged handler and no MinimumSize enforcement at all - it relies
    // entirely on WinForms' own native Dock=Fill cascade). Busy/toast/confirm feedback uses the same
    // two already-existing, already-separate-window mechanisms every other part of this app uses:
    // XLEdgeWaitWindow (ReportGenerator's useWaitWindow:true path) for progress, and
    // MessageFunctions.XLEdgeMessage for one-off messages - matching VB.NET's own FormProcessBar/
    // MessageBox pattern, where busy/message UI is never drawn inside the task pane's own control tree.
    public partial class ADXExcelTaskPane1 : AddinExpress.XL.ADXExcelTaskPane
    {
        private readonly XLEdgeRibbonHelper _ribbonHelper;
        private readonly SemaphoreSlim _initLock = new SemaphoreSlim(1, 1);
        private Task? _webViewInitTask;
        private bool _isInitialized;
        private bool _webViewEventsHooked;

        private static readonly object _envLock = new object();
        private static Task<CoreWebView2Environment>? _sharedEnvironmentTask;

        private static XLEdgeAppState appState => XLEdgeAppState.Instance;

        public ADXExcelTaskPane1()
        {
            InitializeComponent();

            this.Text = string.Empty;

            _ribbonHelper = XLEdgeRibbonHelper.Current;

            this.HandleCreated += ADXExcelTaskPane1_HandleCreated;
        }

        private void ADXExcelTaskPane1_HandleCreated(object sender, EventArgs e)
        {
            SafeFireAndForget(async () =>
            {
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
            }, "Error initializing WebView2 on HandleCreated");
        }

        private void ADXExcelTaskPane1_ADXCloseButtonClick(object sender, ADXCloseButtonClickEventArgs e)
        {
            e.CloseForm = false;
            this.Visible = false;
        }

        private void ADXExcelTaskPane1_ADXBeforeTaskPaneShow(object sender, ADXBeforeTaskPaneShowEventArgs e)
        {
            try
            {
                if (!XLEdgeAppState.Instance.EdgePaneShown)
                {
                    this.Visible = false;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in BeforeTaskPaneShow");
            }
        }

        private void ADXExcelTaskPane1_ADXAfterTaskPaneShow(object sender, ADXAfterTaskPaneShowEventArgs e)
        {
            using (new LogUtility.LogScope("ADXExcelTaskPane1.AfterTaskPaneShow"))
            {
                try
                {
                    // Matches VB.NET's own AfterTaskPaneShow handler, which just forces Me.Width = 500
                    // once on first show and does nothing DPI/resize-aware beyond that - ported the
                    // same way, at 530px instead of VB's 500 per Rama's request.
                    const int TargetWidthPx = 530;
                    LogUtility.LogDebug($"AfterTaskPaneShow: pane size before: {this.Width}x{this.Height}");

                    if (this.Width < TargetWidthPx)
                    {
                        this.Width = TargetWidthPx;
                    }

                    XLEdgeAppState.Instance.EdgePaneShown = false;
                    _ = RefreshLoginNavigationAsync();
                    LogUtility.LogDebug($"AfterTaskPaneShow: pane size after: {this.Width}x{this.Height}");
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Error in AfterTaskPaneShow");
                }
            }
        }

        public void HidePaneSafe()
        {
            try
            {
                if (this.IsDisposed)
                {
                    return;
                }

                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(HidePaneSafe));
                    return;
                }

                this.Visible = false;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to hide ADXExcelTaskPane1 safely.");
            }
        }

        public void RefreshWebViewHeight()
        {
            // No-op: WebCtrl is a native Dock=Fill child, so it always fills the pane's client area
            // exactly - there is no separate height calculation to redo. Kept as a public no-op
            // rather than removed so every existing call site (RefreshListObjectAsync, etc.) keeps
            // compiling unchanged.
        }

        // ================= Ported from XLEdgeCTP.xaml.cs / ADXExcelTaskPane1.vb - operating directly
        // ================= on the Designer-created WebCtrl control.

        private Task RunOnUIAsync(Action action)
        {
            if (this.IsDisposed)
            {
                return Task.CompletedTask;
            }

            if (!this.InvokeRequired)
            {
                action();
                return Task.CompletedTask;
            }

            var tcs = new TaskCompletionSource<bool>();
            try
            {
                this.BeginInvoke(new Action(() =>
                {
                    try { action(); tcs.TrySetResult(true); }
                    catch (Exception ex) { tcs.TrySetException(ex); }
                }));
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
        }

        private Task RunOnUIAsync(Func<Task> func)
        {
            if (this.IsDisposed)
            {
                return Task.CompletedTask;
            }

            if (!this.InvokeRequired)
            {
                return func();
            }

            var tcs = new TaskCompletionSource<bool>();
            try
            {
                this.BeginInvoke(new Action(async () =>
                {
                    try { await func(); tcs.TrySetResult(true); }
                    catch (Exception ex) { tcs.TrySetException(ex); }
                }));
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
        }

        private Task<T> RunOnUIAsync<T>(Func<T> func)
        {
            if (this.IsDisposed)
            {
                // Disposed short-circuit only: the returned default(T) is never actually observed
                // by any caller (they all check IsDisposed/return early themselves too), so the
                // null-forgiving operator here just silences the generic-T nullable warning rather
                // than changing this helper's nullability contract for every call site.
                return Task.FromResult(default(T)!);
            }

            if (!this.InvokeRequired)
            {
                return Task.FromResult(func());
            }

            var tcs = new TaskCompletionSource<T>();
            try
            {
                this.BeginInvoke(new Action(() =>
                {
                    try { tcs.TrySetResult(func()); }
                    catch (Exception ex) { tcs.TrySetException(ex); }
                }));
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
        }

        private Task<T> RunOnUIAsync<T>(Func<Task<T>> func)
        {
            if (this.IsDisposed)
            {
                // Same reasoning as the Func<T> overload above.
                return Task.FromResult(default(T)!);
            }

            if (!this.InvokeRequired)
            {
                return func();
            }

            var tcs = new TaskCompletionSource<T>();
            try
            {
                this.BeginInvoke(new Action(async () =>
                {
                    try { tcs.TrySetResult(await func()); }
                    catch (Exception ex) { tcs.TrySetException(ex); }
                }));
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }

            return tcs.Task;
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

        private async Task EnsureWebViewInitializedAsync()
        {
            if (this.IsDisposed)
            {
                throw new ObjectDisposedException(nameof(ADXExcelTaskPane1));
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

            try
            {
                await _webViewInitTask;
            }
            catch
            {
                _webViewInitTask = null;
                throw;
            }
        }

        private static Task<CoreWebView2Environment> GetOrCreateSharedEnvironmentAsync()
        {
            lock (_envLock)
            {
                if (_sharedEnvironmentTask == null)
                {
                    _sharedEnvironmentTask = CreateSharedEnvironmentAsync();
                }
                return _sharedEnvironmentTask;
            }
        }

        private static async Task<CoreWebView2Environment> CreateSharedEnvironmentAsync()
        {
            try
            {
                string logDir = XLEdgeAppPaths.BrowserLogsFolder;
                DirectoryInfo di = new DirectoryInfo(logDir);
                if (!di.Exists)
                    di.Create();

                var envOptions = new CoreWebView2EnvironmentOptions
                {
                    AllowSingleSignOnUsingOSPrimaryAccount = true
                };

                return await CoreWebView2Environment.CreateAsync(
                    browserExecutableFolder: null,
                    userDataFolder: di.FullName,
                    options: envOptions);
            }
            catch
            {
                lock (_envLock)
                {
                    _sharedEnvironmentTask = null;
                }
                throw;
            }
        }

        private async Task InitializeWebViewInternalAsync()
        {
            await RunOnUIAsync(async () =>
            {
                var env = await GetOrCreateSharedEnvironmentAsync();

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

                    WebCtrl.Visible = true;
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

        public async Task<bool> LogoutAsync(string loginUrl, CancellationToken token)
        {
            if (string.IsNullOrWhiteSpace(loginUrl))
            {
                LogUtility.LogWarn("LogoutAsync skipped because loginUrl is empty.");
                return false;
            }

            try
            {
                return await RunOnUIAsync(async () =>
                {
                    if (WebCtrl == null || WebCtrl.CoreWebView2 == null)
                    {
                        LogUtility.LogWarn("WebCtrl/CoreWebView2 is not ready for logout - skipping (nothing to log out of).");
                        return false;
                    }

                    string navUrl = $"{loginUrl.TrimEnd('/')}/web/secure/applogout";
                    var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
                    EventHandler<CoreWebView2NavigationCompletedEventArgs>? handler = null;

                    handler = (s, e) =>
                    {
                        try
                        {
                            WebCtrl.CoreWebView2.NavigationCompleted -= handler;
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"{nameof(LogoutAsync)}: failed to unsubscribe NavigationCompleted handler - {ex.Message}");
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
                                LogUtility.LogDebug($"{nameof(LogoutAsync)}: failed to unsubscribe NavigationCompleted handler after timeout - {ex.Message}");
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
                            LogUtility.LogDebug($"{nameof(LogoutAsync)}: failed to unsubscribe NavigationCompleted handler after exception - {unsubEx.Message}");
                        }

                        LogUtility.LogException(ex, "Exception during logout navigation.");
                        return false;
                    }
                });
            }
            catch (OperationCanceledException)
            {
                LogUtility.LogWarn("LogoutAsync cancelled.");
                throw;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to logout in ADXExcelTaskPane1.LogoutAsync");
                return false;
            }
        }

        public async Task ExecuteScriptAsync(string script)
        {
            if (string.IsNullOrWhiteSpace(script))
            {
                return;
            }

            try
            {
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
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to execute script in ADXExcelTaskPane1.ExecuteScriptAsync");
            }
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

        // Previously showed the broadcast message via the pane-embedded AppOverlay (ShowInfoAsync).
        // Uses MessageFunctions.XLEdgeMessage - the app's existing standalone one-off-message window
        // (the C# port of VB's XLEdgeMsgDisplay) - matching VB.NET's own pattern instead of drawing
        // anything inside the pane's own control tree.
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

            if (!string.IsNullOrWhiteSpace(broadcastMsg))
            {
                MessageFunctions.XLEdgeMessage(broadcastMsg, System.Windows.Forms.MessageBoxIcon.Information);
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

        private static void UpdateExcelTabLabel()
        {
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

                    string loginUrl = appState.LoginUrl;

                    if (this.Visible)
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

        // Report-generation calls pass useWaitWindow: true instead of an AppOverlay reference, so
        // ReportGenerator shows its own already-existing standalone XLEdgeWaitWindow - the same
        // convention AddinModule.cs's own drilldown call site already used
        // (CreateReportFromTitleAsync(childTitle, useWaitWindow: true, ...)).
        private async void WebView_DocumentTitleChanged(object sender, object e)
        {
            try
            {
                string title = WebCtrl?.CoreWebView2?.DocumentTitle ?? string.Empty;
                LogUtility.LogDebug($"Document title changed: {title}");

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
                            string? runIds = await FetchWorkbookRerunIdsAsync();
                            if (!string.IsNullOrWhiteSpace(runIds))
                            {
                                SafeFireAndForget(
                                    () => XLEdge.Helpers.ReportGenerator.CreateMultiDataReportsAsync(runIds, useWaitWindow: true),
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
                            SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.CreateReportFromTitleAsync(title, useWaitWindow: true), "Handle document title change");
                        }
                        break;

                    case "Logs":
                        SafeFireAndForget(() => XLEdge.Helpers.ReportGenerator.CreateLogsReportAsync(title, useWaitWindow: true), "Handle Logs document title change");
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

        private async Task<string?> FetchWorkbookRerunIdsAsync()
        {
            const string script = @"(() => {
                let element = document.querySelector('[reruntype=xledgeworkbookrerun]');
                return element ? element.getAttribute('newrunids') : null;
                })()";

            try
            {
                await EnsureWebViewInitializedAsync();

                string? result = await RunOnUIAsync(async () =>
                {
                    if (WebCtrl?.CoreWebView2 == null)
                    {
                        return (string?)null;
                    }

                    return await WebCtrl.CoreWebView2.ExecuteScriptAsync(script);
                });

                LogUtility.LogDebug($"FetchWorkbookRerunIdsAsync result: {result}");

                if (string.IsNullOrWhiteSpace(result) || result == "null")
                {
                    return null;
                }

                // Null-forgiving: the guard above already proves result is neither null/whitespace
                // nor the literal string "null" - Roslyn's flow analysis just doesn't narrow through
                // the compound "||" condition here.
                return result!.Trim('"');
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error in FetchWorkbookRerunIdsAsync");
                return null;
            }
        }

        private async Task NavigateToLoginUrlSafeAsync()
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
                        WebCtrl.Visible = true;
                        WebCtrl.Source = new Uri(urlNavigate);
                    });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Error navigating to login URL");
            }
        }

        public async Task NavigateBlankAsync()
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

        public async Task RefreshLoginNavigationAsync()
        {
            try
            {
                await RunOnUIAsync(() =>
                {
                    if (!string.IsNullOrWhiteSpace(appState.LoginUrl))
                        SetPaneCaption(appState.LoginUrl);
                });

                await NavigateToLoginUrlSafeAsync();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Failed to refresh login navigation in ADXExcelTaskPane1.RefreshLoginNavigationAsync");
            }
        }

        private void SetPaneCaption(string caption)
        {
            try
            {
                this.Text = caption;
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
                if (this.InvokeRequired)
                {
                    this.Invoke(new Action(ReleaseFocusToExcel));
                    return;
                }

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
                }

                try
                {
                    this.Focus();

                    var excelApp = ExcelApplicationHelper.GetActiveExcelApplication();
                    if (excelApp != null)
                    {
                        ExcelWindowHelper.ActivateExcelMainWindow(excelApp);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogDebug($"ADXExcelTaskPane1.ReleaseFocusToExcel: focus failed - {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "ADXExcelTaskPane1.ReleaseFocusToExcel failed");
            }
        }
    }
}
