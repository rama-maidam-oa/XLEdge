using MahApps.Metro.IconPacks;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using XLEdge.Helpers;
using XLEdge.Utilities;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeWaitWindow.xaml
    /// </summary>
    public partial class XLEdgeWaitWindow : DpiAwareWindow, IDisposable
    {
        private readonly CancellationHelper _helper;
        private readonly Stopwatch _stopwatch;
        private readonly DispatcherTimer _timer;

        private volatile bool _allowClose = false;
        private volatile bool _isClosing = false;
        private readonly object _closeLock = new object();

        private bool _disposed = false;

        public CancellationToken Token => _helper?.GetToken() ?? CancellationToken.None;

        public bool IsCancellationRequested => _helper?.IsCancellationRequested ?? false;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();

                _helper?.Dispose();

                BtnCancel.Click -= BtnCancel_Click;
                this.Closing -= OnClosingGate;
                this.Closed -= OnClosedCleanup;
            }

            _disposed = true;
        }

        private void DisposeManagedResources()
        {
            CleanupWindow();
            StopTimerAndStopwatch();
        }

        private void CleanupWindow()
        {
            try
            {
                if (IsLoaded) RequestClose();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        private void StopTimerAndStopwatch()
        {
            try
            {
                _timer?.Stop();
                _stopwatch?.Reset();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public XLEdgeWaitWindow(CancellationHelper helper = null)
        {
            InitializeComponent();

            EnableEscapeToClose = false;

            _helper = helper ?? new CancellationHelper();

            EnhancedDragDropHelper.EnableWindowDrag(this);

            BtnCancel.Click += BtnCancel_Click;

            _stopwatch = new Stopwatch();
            _timer = new DispatcherTimer(DispatcherPriority.Normal)
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _timer.Tick += (_, __) =>
            {
                ElapsedTimeLabel.Text = $"Time Elapsed: {_stopwatch.Elapsed:hh\\:mm\\:ss}";
            };

            this.Closing += OnClosingGate;
            this.Closed += OnClosedCleanup;

            BtnCancel.IsEnabled = true;
        }

        private void OnClosingGate(object sender, CancelEventArgs e)
        {
            if (!_allowClose)
            {
                e.Cancel = true;
                this.Activate();
            }
        }

        private void OnClosedCleanup(object sender, EventArgs e)
        {
            lock (_closeLock)
            {
                if (_isClosing) return;
                _isClosing = true;
            }

            try
            {
                _timer?.Stop();
                _stopwatch?.Reset();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
            }
            finally
            {
                try { _helper?.Dispose(); }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex);
                }

                ExcelApplicationHelper.TryGetActiveExcelApplication(out Microsoft.Office.Interop.Excel.Application excelApp);
                ExcelWindowHelper.ActivateExcelMainWindow(excelApp);
            }
        }

        public void StartMonitoring()
        {
            VerifyAccess();

            _stopwatch.Start();
            _timer.Start();

            ProgressBarControl.IsIndeterminate = true;
            ProgressBarControl.Visibility = Visibility.Visible;
            BtnCancel.IsEnabled = true;
        }

        public void SetProcessTitle(string title, PackIconFontAwesomeKind icon)
        {
            VerifyAccess();

            titleIcon.Kind = icon;

            if (!string.IsNullOrWhiteSpace(title))
                titleText.Text = title;
        }

        public void SetProcessMessage(string message)
        {
            VerifyAccess();
            processLabelName.Text = message;
        }

        public void RequestClose()
        {
            lock (_closeLock)
            {
                if (_isClosing) return;
                _allowClose = true;
            }

            if (Dispatcher.CheckAccess())
            {
                SafeClose();
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(SafeClose), DispatcherPriority.Background);
            }
        }

        private void SafeClose()
        {
            lock (_closeLock)
            {
                if (_isClosing) return;
                _isClosing = true;
            }

            try
            {
                Close();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error closing wait window: {ex.Message}");
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            BtnCancel.IsEnabled = false;

            try
            {
                _helper.Cancel();
                processLabelName.Text = "Cancelling...";
                RequestClose();
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex);
                BtnCancel.IsEnabled = true;
            }
        }

        public async Task<bool?> ShowConfirmToastAsync(string message)
        {
            return await Dispatcher
                .InvokeAsync(() => AppOverlayControl.ShowConfirmAsync(message), DispatcherPriority.Normal)
                .Task.Unwrap()
                .ConfigureAwait(false);
        }
    }
}
