using MahApps.Metro.IconPacks;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Effects;
using System.Windows.Threading;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for AppOverlay.xaml
    /// </summary>
    public partial class AppOverlay : UserControl
    {
        // Attached property to mark elements blurred by this overlay (across instances)
        public static readonly DependencyProperty BlurAppliedProperty =
            DependencyProperty.RegisterAttached("BlurApplied", typeof(bool), typeof(AppOverlay), new PropertyMetadata(false));

        public static void SetBlurApplied(UIElement element, bool value)
        {
            try
            {
                element?.SetValue(BlurAppliedProperty, value);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"SetBlurApplied failed: {ex.Message}");
            }
        }

        public static bool GetBlurApplied(UIElement element)
        {
            try
            {
                return (bool)(element?.GetValue(BlurAppliedProperty) ?? false);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"GetBlurApplied failed: {ex.Message}");
                return false;
            }
        }
        // Keep track of elements that were blurred so we can restore them
        private readonly System.Collections.Generic.List<(UIElement Element, Effect OriginalEffect, bool OriginalHitTest)> _blurredElements = new();

        // Tracks WebView2 elements temporarily hidden so overlays can render above them
        // (WebView2 renders outside WPF's composition layer regardless of Panel.ZIndex).
        private readonly System.Collections.Generic.List<UIElement> _hiddenWebViewElements = new();

        // Tracks which overlay types (Toast/Busy/Confirm) currently require WebView2 siblings to be
        // hidden, so WebView2 is hidden/restored only when the combined active state changes and
        // multiple overlays can be shown concurrently without interfering with each other.
        private bool _toastHidesWebView2;
        private bool _busyHidesWebView2;
        private bool _confirmHidesWebView2;

        private DispatcherTimer _busyTimer;
        private DateTime? _busyStart;
        private DispatcherTimer _toastTimer;
        private TaskCompletionSource<bool> _activeToastTcs;
        private RoutedEventHandler YesHandler, NoHandler, CancelHandler, BusyCancelHandler;
        private EventHandler _hideBusyHandler;

        public AppOverlay()
        {
            InitializeComponent();
        }

        public bool IsBusyVisible => BusyOverlay.Visibility == Visibility.Visible;
        public bool IsConfirmVisible => ConfirmOverlay.Visibility == Visibility.Visible;

        /// <summary>
        /// Sets the toast's MaxHeight based on the parent container's current height, so the toast
        /// sizes correctly against available space and can grow with its content.
        /// </summary>
        private void UpdateToastMaxHeight()
        {
            if (Toast == null)
            {
                return;
            }

            try
            {
                double availableHeight = 0;

                if (this.Parent is FrameworkElement parentElement && parentElement.ActualHeight > 0)
                {
                    availableHeight = parentElement.ActualHeight;
                }
                else if (this.ActualHeight > 0)
                {
                    availableHeight = this.ActualHeight;
                }
                else
                {
                    // Last-resort fallback (e.g. the very first call, before anything has ever been
                    // measured) - better than leaving the toast pinned to a tiny/zero height.
                    availableHeight = SystemParameters.WorkArea.Height;
                }

                Toast.MaxHeight = availableHeight * 0.9;
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"Failed to update toast max height: {ex.Message}");
            }
        }

        // === Toast ===
        public void ShowToast(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60)
        {
            if (Toast == null) return;

            CollapseBusyOverlayForToast();

            this.Visibility = Visibility.Visible;
            UpdateToastMaxHeight();
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            // Block input to underlying UI while toast is visible
            if (ToastOverlay != null)
            {
                ToastOverlay.Visibility = Visibility.Visible;
                try
                {
                    ToastOverlay.Focus();
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Could not focus ToastOverlay: {ex.Message}");
                }
            }

            // Apply blur to underlying sibling elements (mirror/blur effect)
            ApplyBlurToSiblings();

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += OnToastTimerTick;
            _toastTimer.Start();
        }

        public void ShowSuccess(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleCheckSolid, Brushes.LimeGreen, 60);
        public void ShowError(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleXmarkSolid, Brushes.Red, 60);
        public void ShowWarning(string message) => ShowToast(message, PackIconFontAwesomeKind.TriangleExclamationSolid, Brushes.Orange, 60);
        public void ShowInfo(string message) => ShowToast(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, 60);

        public Task ShowToastAsync(string message, PackIconFontAwesomeKind icon, Brush color, int durationSeconds = 60)
        {
            _activeToastTcs?.TrySetResult(true);
            var tcs = new TaskCompletionSource<bool>();
            _activeToastTcs = tcs;

            CollapseBusyOverlayForToast();

            this.Visibility = Visibility.Visible;
            UpdateToastMaxHeight();
            Toast.Visibility = Visibility.Visible;
            Toast.Opacity = 1;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Panel.SetZIndex(Toast, 10001);

            // Block input to underlying UI while toast is visible
            if (ToastOverlay != null)
            {
                ToastOverlay.Visibility = Visibility.Visible;
                try
                {
                    ToastOverlay.Focus();
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Could not focus ToastOverlay: {ex.Message}");
                }
            }

            // Apply blur to underlying sibling elements (mirror/blur effect)
            ApplyBlurToSiblings();

            ToastMessage.Text = message;
            ToastIcon.Kind = icon;
            ToastIcon.Foreground = color;

            Panel.SetZIndex(this, 9999);

            _toastTimer?.Stop();
            _toastTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(durationSeconds) };
            _toastTimer.Tick += (s, e) =>
            {
                _toastTimer.Stop();
                var fade = new DoubleAnimation(1, 0, TimeSpan.FromSeconds(0.5));
                fade.Completed += (s2, e2) =>
                {
                    Toast.Opacity = 0;
                    // Hide the input blocker when toast fades out
                    if (ToastOverlay != null)
                        ToastOverlay.Visibility = Visibility.Collapsed;

                    // Remove blur from siblings
                    RemoveBlurFromSiblings();
                    if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;
                    tcs.TrySetResult(true);
                    if (_activeToastTcs == tcs)
                        _activeToastTcs = null;
                };
                Toast.BeginAnimation(Border.OpacityProperty, fade);
            };
            _toastTimer.Start();

            return tcs.Task;
        }

        public async Task ShowSuccessAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleCheckSolid, Brushes.LimeGreen, 60);
        }

        public async Task ShowErrorAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleXmarkSolid, Brushes.Red, 60);
        }

        public async Task ShowWarningAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.TriangleExclamationSolid, Brushes.Orange, 60);
        }

        public async Task ShowInfoAsync(string message)
        {
            await ShowToastAsync(message, PackIconFontAwesomeKind.CircleInfoSolid, Brushes.DodgerBlue, 60);
        }

        public void DismissToast()
        {
            if (Toast == null || Toast.Visibility != Visibility.Visible)
                return;

            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            Toast.BeginAnimation(Border.OpacityProperty, null);
            Toast.Opacity = 0;
            Toast.Visibility = Visibility.Collapsed;

            if (ToastOverlay != null)
                ToastOverlay.Visibility = Visibility.Collapsed;

            // Remove blur from siblings
            RemoveBlurFromSiblings();

            if (BusyOverlay.Visibility != Visibility.Visible && ConfirmOverlay.Visibility != Visibility.Visible)
                this.Visibility = Visibility.Collapsed;
        }
        private void OnToastTimerTick(object sender, EventArgs e)
        {
            _toastTimer.Stop();
            // Immediately dismiss toast and remove blur when timer elapses
            DismissToast();
        }

        private void BtnCloseToast_Click(object sender, RoutedEventArgs e)
        {
            // Immediately dismiss toast and remove blur when user clicks close
            _toastTimer?.Stop();
            _activeToastTcs?.TrySetResult(true);
            _activeToastTcs = null;
            DismissToast();
        }

        // === Busy Overlay ===
        public async Task ShowBusyasynTask(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            ShowBusyasyn(message, cancelAction);
            await Task.CompletedTask;
        }
        public void ShowBusyasyn(string message = "Please wait...", Func<Task> cancelAction = null)
        {
            BusyMessage.Text = message ?? "Please wait...";
            this.Visibility = Visibility.Visible;
            this.Opacity = 1.0;
            Panel.SetZIndex(this, 9999);

            // ?? Bring to front of parent container
            if (this.Parent is UIElement parent)
                Panel.SetZIndex(parent, 0);

            BusyOverlay.Visibility = Visibility.Visible;
            BusyOverlay.Opacity = 1;
            BusyOverlay.IsHitTestVisible = true;

            // Hide WebView2 siblings while the busy overlay is visible
            SetWebView2HiddenState(ref _busyHidesWebView2, true);

            if (BusyCancelHandler != null)
            {
                BtnCancelBusy.Click -= BusyCancelHandler;
                BusyCancelHandler = null;
            }

            BusyCancelHandler = async (sender, e) =>
            {
                BtnCancelBusy.IsEnabled = false;
                try
                {
                    // ?? Run async cancel if provided
                    if (cancelAction != null)
                        await cancelAction();
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, "Busy cancel handler");
                }
                finally
                {
                    await HideBusyAsync();
                }
            };

            BtnCancelBusy.Click += BusyCancelHandler;
            BtnCancelBusy.Visibility = Visibility.Visible;
            BtnCancelBusy.IsEnabled = true;

            // Start elapsed timer
            try
            {
                _busyStart = DateTime.UtcNow;
                BusyElapsed.Text = "Time Elapsed: 00:00:00";

                _busyTimer?.Stop();
                _busyTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
                _busyTimer.Tick += (s, e) =>
                {
                    if (_busyStart == null) return;
                    var span = DateTime.UtcNow - _busyStart.Value;
                    BusyElapsed.Text = $"Time Elapsed: {FormatTimeSpan(span)}";
                };
                _busyTimer.Start();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Start busy timer");
            }

            if (this.Resources["ShowBusy"] is Storyboard sb)
                sb.Begin(this);
        }
        private async void BtnHideBusy_Click(object sender, RoutedEventArgs e)
        {
            await HideBusyAsync();
        }
        public async Task HideBusyAsync()
        {
            var tcs = new TaskCompletionSource<bool>();

            await Dispatcher.InvokeAsync(() =>
            {
                if (this.Resources["HideBusy"] is not Storyboard sb)
                {
                    // Fallback: if storyboard missing, hide immediately
                    StopBusyTimer();
                    BusyOverlay.Visibility = Visibility.Collapsed;
                    SetWebView2HiddenState(ref _busyHidesWebView2, false);
                    this.Visibility = Visibility.Collapsed;
                    tcs.TrySetResult(true);
                    return;
                }

                // Remove previous handler to prevent multiple triggers
                if (_hideBusyHandler != null)
                    sb.Completed -= _hideBusyHandler;

                _hideBusyHandler = (s, e) =>
                {
                    StopBusyTimer();

                    BusyOverlay.Visibility = Visibility.Collapsed;
                    SetWebView2HiddenState(ref _busyHidesWebView2, false);

                    // Hide root only if no other overlay is visible
                    if (Math.Abs(Toast.Opacity - 0) < 0.0001 && ConfirmOverlay.Visibility != Visibility.Visible)
                        this.Visibility = Visibility.Collapsed;

                    sb.Completed -= _hideBusyHandler; // Clean up handler
                    _hideBusyHandler = null;

                    tcs.TrySetResult(true);
                };

                sb.Completed += _hideBusyHandler;

                // Ensure it stays visible during fade
                BusyOverlay.Visibility = Visibility.Visible;
                BusyOverlay.IsHitTestVisible = false; // allow clicks to pass during fade
                sb.Begin(this, true); // true = controllable animation
            });

            await tcs.Task;
        }

        /// <summary>
        /// Immediately collapses the busy overlay if it's visible, so a toast about to be shown
        /// isn't hidden behind it. BusyOverlay and ToastOverlay share the same Panel.ZIndex (10000)
        /// in AppOverlay.xaml, and BusyOverlay is declared later, so without this a toast fired
        /// while a busy spinner is up would render underneath the busy overlay's dark scrim -
        /// the taskpane just looks stuck/blurred with no visible message. Skips the normal
        /// HideBusyAsync fade since the toast is about to cover the same area anyway.
        /// </summary>
        private void CollapseBusyOverlayForToast()
        {
            if (BusyOverlay.Visibility != Visibility.Visible) return;

            StopBusyTimer();
            BusyOverlay.Visibility = Visibility.Collapsed;
            SetWebView2HiddenState(ref _busyHidesWebView2, false);
        }

        private void StopBusyTimer()
        {
            try
            {
                _busyTimer?.Stop();
                _busyTimer = null;
                _busyStart = null;
                if (BusyElapsed != null)
                    BusyElapsed.Text = string.Empty;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "Stop busy timer");
            }
        }

        private static string FormatTimeSpan(TimeSpan span)
        {
            // Format as HH:MM:SS
            return string.Format("{0:D2}:{1:D2}:{2:D2}", (int)span.TotalHours, span.Minutes, span.Seconds);
        }


        // === Confirmation ===
        public void ShowConfirm(string message, Action yesAction, Action noAction = null, Action cancelAction = null)
        {
            ConfirmText.Text = message;
            ConfirmOverlay.Visibility = Visibility.Visible;
            this.Visibility = Visibility.Visible;
            Panel.SetZIndex(this, 9999);
            SetWebView2HiddenState(ref _confirmHidesWebView2, true);

            ConfirmPopup.RenderTransform = new ScaleTransform(0.8, 0.8);

            if (YesHandler != null) BtnYes.Click -= YesHandler;
            if (NoHandler != null) BtnNo.Click -= NoHandler;
            if (CancelHandler != null) BtnCancel.Click -= CancelHandler;

            YesHandler = (s, e) => { HideConfirm(); yesAction?.Invoke(); };
            NoHandler = (s, e) => { HideConfirm(); noAction?.Invoke(); };
            CancelHandler = (s, e) => { HideConfirm(); cancelAction?.Invoke(); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);
        }

        public Task<bool?> ShowConfirmAsync(string message)
        {
            var tcs = new TaskCompletionSource<bool?>();

            ConfirmText.Text = message;
            ConfirmOverlay.Visibility = Visibility.Visible;
            this.Visibility = Visibility.Visible;
            Panel.SetZIndex(this, 9999);
            SetWebView2HiddenState(ref _confirmHidesWebView2, true);

            ConfirmPopup.RenderTransform = new ScaleTransform(0.8, 0.8);

            if (YesHandler != null) BtnYes.Click -= YesHandler;
            if (NoHandler != null) BtnNo.Click -= NoHandler;
            if (CancelHandler != null) BtnCancel.Click -= CancelHandler;

            YesHandler = (s, e) => { HideConfirm(); tcs.TrySetResult(true); };
            NoHandler = (s, e) => { HideConfirm(); tcs.TrySetResult(false); };
            CancelHandler = (s, e) => { HideConfirm(); tcs.TrySetResult(null); };

            BtnYes.Click += YesHandler;
            BtnNo.Click += NoHandler;
            BtnCancel.Click += CancelHandler;

            if (this.Resources["ShowConfirm"] is Storyboard sb)
                sb.Begin(this);

            return tcs.Task;
        }

        private void HideConfirm()
        {
            if (this.Resources["HideConfirm"] is Storyboard sb)
            {
                EventHandler onComplete = null;
                onComplete = (s, e) =>
                {
                    ConfirmOverlay.Visibility = Visibility.Collapsed;
                    SetWebView2HiddenState(ref _confirmHidesWebView2, false);
                    if (BusyOverlay.Visibility != Visibility.Visible && Math.Abs(Toast.Opacity - 0) < 0.0001)
                        this.Visibility = Visibility.Collapsed;

                    sb.Completed -= onComplete;
                };

                sb.Completed += onComplete;
                sb.Begin(this);
            }
            else
            {
                // Fallback when storyboard is missing
                ConfirmOverlay.Visibility = Visibility.Collapsed;
                SetWebView2HiddenState(ref _confirmHidesWebView2, false);
                if (BusyOverlay.Visibility != Visibility.Visible && Math.Abs(Toast.Opacity - 0) < 0.0001)
                    this.Visibility = Visibility.Collapsed;
            }
        }

        // Add helper methods for blur
        private void ApplyBlurToSiblings()
        {
            try
            {
                // find parent panel that contains this overlay
                if (this.Parent is Panel parentPanel)
                {
                    _blurredElements.Clear();
                    foreach (UIElement child in parentPanel.Children)
                    {
                        if (child == this) continue;

                        var originalEffect = child.Effect;
                        var originalHit = child.IsHitTestVisible;

                        // store original state
                        _blurredElements.Add((child, originalEffect, originalHit));

                        // apply blur and disable hit testing (overlay + blocker will handle input)
                        child.Effect = new BlurEffect { Radius = 6 };
                        child.IsHitTestVisible = false;
                        SetBlurApplied(child, true);
                    }
                }
                else
                {
                    // fallback: try window content
                    var wnd = Window.GetWindow(this);
                    if (wnd?.Content is Panel wndPanel)
                    {
                        _blurredElements.Clear();
                        foreach (UIElement child in wndPanel.Children)
                        {
                            if (child == this) continue;

                            var originalEffect = child.Effect;
                            var originalHit = child.IsHitTestVisible;
                            _blurredElements.Add((child, originalEffect, originalHit));

                            child.Effect = new BlurEffect { Radius = 6 };
                            child.IsHitTestVisible = false;
                            SetBlurApplied(child, true);
                        }
                    }
                }

                SetWebView2HiddenState(ref _toastHidesWebView2, true);
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"ApplyBlurToSiblings failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Updates whether a given overlay type needs WebView2 hidden, and hides or restores
        /// WebView2 only when the combined active state across all overlay types actually changes.
        /// </summary>
        private void SetWebView2HiddenState(ref bool flag, bool hidden)
        {
            bool wasAnyActive = _toastHidesWebView2 || _busyHidesWebView2 || _confirmHidesWebView2;
            flag = hidden;
            bool isAnyActive = _toastHidesWebView2 || _busyHidesWebView2 || _confirmHidesWebView2;

            if (!wasAnyActive && isAnyActive)
            {
                HideWebView2DescendantsOfSiblings();
            }
            else if (wasAnyActive && !isAnyActive)
            {
                RestoreHiddenWebView2Descendants();
            }
        }

        /// <summary>
        /// Finds this overlay's sibling elements and hides any WebView2 descendants of each.
        /// </summary>
        private void HideWebView2DescendantsOfSiblings()
        {
            try
            {
                _hiddenWebViewElements.Clear();

                if (this.Parent is Panel parentPanel)
                {
                    foreach (UIElement sibling in parentPanel.Children)
                    {
                        if (sibling == this) continue;
                        HideWebView2Descendants(sibling);
                    }
                    return;
                }

                var wnd = Window.GetWindow(this);
                if (wnd?.Content is Panel wndPanel)
                {
                    foreach (UIElement sibling in wndPanel.Children)
                    {
                        if (sibling == this) continue;
                        HideWebView2Descendants(sibling);
                    }
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"HideWebView2DescendantsOfSiblings failed: {ex.Message}");
            }
        }

        /// <summary>
        /// Recursively finds any WebView2 descendants of <paramref name="root"/> and temporarily
        /// hides them (Visibility.Hidden) so this overlay renders on top of them.
        /// </summary>
        private void HideWebView2Descendants(DependencyObject root)
        {
            if (root == null)
            {
                return;
            }

            try
            {
                if (root is Microsoft.Web.WebView2.Wpf.WebView2 webView)
                {
                    if (webView.Visibility == Visibility.Visible)
                    {
                        webView.Visibility = Visibility.Hidden;
                        _hiddenWebViewElements.Add(webView);
                    }
                    return;
                }

                int childCount = System.Windows.Media.VisualTreeHelper.GetChildrenCount(root);
                for (int i = 0; i < childCount; i++)
                {
                    HideWebView2Descendants(System.Windows.Media.VisualTreeHelper.GetChild(root, i));
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"HideWebView2Descendants failed: {ex.Message}");
            }
        }

        private void RestoreHiddenWebView2Descendants()
        {
            if (_hiddenWebViewElements.Count == 0)
            {
                return;
            }

            foreach (var element in _hiddenWebViewElements)
            {
                try
                {
                    element.Visibility = Visibility.Visible;
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Failed to restore WebView2 visibility: {ex.Message}");
                }
            }

            _hiddenWebViewElements.Clear();
        }

        private void RemoveBlurFromSiblings()
        {
            try
            {
                SetWebView2HiddenState(ref _toastHidesWebView2, false);

                if (_blurredElements == null || _blurredElements.Count == 0) return;

                foreach (var entry in _blurredElements)
                {
                    try
                    {
                        if (entry.Element == null) continue;
                        entry.Element.Effect = entry.OriginalEffect;
                        entry.Element.IsHitTestVisible = entry.OriginalHitTest;
                        SetBlurApplied(entry.Element, false);
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogWarn($"Restore blur failed for element: {ex.Message}");
                    }
                }

                _blurredElements.Clear();
                // Fallback: if other AppOverlay instances applied blur but didn't
                // populate our _blurredElements (different instance), clear any
                // residual blurred elements in all windows that were tagged.
                try
                {
                    if (Application.Current != null)
                    {
                        foreach (Window w in Application.Current.Windows)
                        {
                            try
                            {
                                var root = w?.Content as DependencyObject;
                                if (root == null) continue;
                                ClearBlurFromVisualTree(root);
                            }
                            catch (Exception ex)
                            {
                                LogUtility.LogWarn($"Fallback blur clear failed for window: {ex.Message}");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"Fallback blur clear failed: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogWarn($"RemoveBlurFromSiblings failed: {ex.Message}");
            }
        }

        private void ClearBlurFromVisualTree(DependencyObject node)
        {
            if (node == null) return;

            int children = System.Windows.Media.VisualTreeHelper.GetChildrenCount(node);
            for (int i = 0; i < children; i++)
            {
                try
                {
                    var child = System.Windows.Media.VisualTreeHelper.GetChild(node, i);
                    if (child is UIElement ui)
                    {
                        try
                        {
                            if (GetBlurApplied(ui) || (ui.Effect is BlurEffect))
                            {
                                ui.Effect = null;
                                ui.IsHitTestVisible = true;
                                SetBlurApplied(ui, false);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogWarn($"ClearBlurFromVisualTree - UIElement: {ex.Message}");
                        }
                    }

                    // Recurse
                    ClearBlurFromVisualTree(child);
                }
                catch (Exception ex)
                {
                    LogUtility.LogWarn($"ClearBlurFromVisualTree - child iteration: {ex.Message}");
                }
            }
        }
    }
}