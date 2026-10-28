using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using XLEdge.Contracts;

namespace XLEdge.Views
{
    /// <summary>
    /// The Reload window: picks where the next core version comes from. Only Offline works for now (a folder holding a
    /// manifest.json + zip pair, checked by <see cref="OfflineReleaseScanner"/>); Online is shown but disabled until the
    /// server endpoints exist. On Reload the dialog returns true with <see cref="SelectedSource"/> and <see cref="ScanResult"/>
    /// set; the shell stages the release into AddinCore\Manifest\ after the old core is unloaded, and the bootstrapper picks it up.
    /// </summary>
    internal partial class XLEdgeReloadSourcePicker : Window
    {
        private const string LogPrefix = "XLEdgeReloadSourcePicker: ";

        // Step-log glyphs: bullet, check mark, warning sign, ballot X.
        private static readonly string StepGlyph = ((char)0x2022).ToString();
        private static readonly string SuccessGlyph = ((char)0x2713).ToString();
        private static readonly string WarningGlyph = ((char)0x26A0).ToString();
        private static readonly string FailureGlyph = ((char)0x2717).ToString();

        private static readonly Guid FolderIdDownloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");

        private readonly IXLEdgeContext _context;
        private OfflineScanResult _candidate;
        private bool _busy;
        private bool _isClosed;

        public XLEdgeReloadSourcePicker(IXLEdgeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
        }

        /// <summary>Whether the Online source can be chosen (not yet: it arrives with the server endpoints).</summary>
        public bool IsOnlineEnabled => RbOnline.IsEnabled;

        public bool IsOfflineSelected => RbOffline.IsChecked == true;

        /// <summary>"Offline" once a verified release was chosen and the dialog accepted; null until then.</summary>
        public string SelectedSource { get; private set; }

        /// <summary>The verified release the user chose (not staged yet: the shell stages it once the reload goes ahead).</summary>
        public OfflineScanResult ScanResult { get; private set; }

        protected override void OnClosed(EventArgs e)
        {
            _isClosed = true;
            base.OnClosed(e);
        }

        // ------------------------------------------------------------------
        // Offline flow
        // ------------------------------------------------------------------

        private async void BtnBrowse_Click(object sender, RoutedEventArgs e)
        {
            // async void: nothing may escape into WPF / Add-in Express.
            try
            {
                string folder;
                using (var dialog = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dialog.SelectedPath = Directory.Exists(TxtFolder.Text) ? TxtFolder.Text : GetDownloadsFolder();
                    if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;
                    folder = dialog.SelectedPath;
                }

                TxtFolder.Text = folder;
                await ScanFolderAsync(folder);
            }
            catch (Exception ex)
            {
                LogFailure($"Folder scan failed: {ex.Message}", ex);
                SetBusy(false);
            }
        }

        private async Task ScanFolderAsync(string folder)
        {
            ResetValidation();
            SetBusy(true);
            try
            {
                LogStep($"Scanning folder: {folder}");
                if (!Directory.Exists(folder))
                {
                    LogWarning($"Folder not found: {folder}");
                    return;
                }

                LogStep("Checking the manifest.json + zip pair (parse, checksum, contract, version)...");
                string loadedVersion = _context.Version;
                string loadedReleaseDate = _context.ReleaseDate;

                // The checksum reads the whole zip: keep it off the UI thread.
                OfflineScanResult result = await Task.Run(() => OfflineReleaseScanner.Scan(folder, loadedVersion, loadedReleaseDate));
                ApplyScanResult(result);
            }
            catch (Exception ex)
            {
                LogFailure($"Folder scan failed: {ex.Message}", ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        /// <summary>Shows a finished scan. May run on a pool thread (no synchronization context in the add-in), so it marshals.</summary>
        private void ApplyScanResult(OfflineScanResult result)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => ApplyScanResult(result)));
                return;
            }
            if (_isClosed || result == null) return;

            if (result.Ok)
            {
                _candidate = result;
                LogSuccess(result.Message);
            }
            else if (result.NotNewer)
            {
                LogStep("No newer version available.");
                MessageBox.Show(this, result.Message, "No Updates Available", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                LogFailure(result.Message);
            }
        }

        // ------------------------------------------------------------------
        // Reload / Cancel
        // ------------------------------------------------------------------

        private void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            if (_busy || _candidate == null) return;
            try
            {
                ScanResult = _candidate;
                SelectedSource = "Offline";
                SafeLog(l => l.LogDebug(LogPrefix + $"version {_candidate.Release?.Version} chosen - proceeding with reload."));
                CloseWith(true);
            }
            catch (Exception ex)
            {
                ScanResult = null;
                SelectedSource = null;
                LogFailure($"Failed to start the reload: {ex.Message}", ex);
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            try { CloseWith(false); }
            catch (Exception ex) { SafeLog(l => l.LogError(LogPrefix + "closing the window failed.", ex.ToString())); }
        }

        /// <summary>Sets DialogResult when shown with ShowDialog (it throws otherwise), then closes.</summary>
        private void CloseWith(bool result)
        {
            try
            {
                DialogResult = result;
            }
            catch (InvalidOperationException)
            {
                // Not shown modally.
            }
            if (!_isClosed) Close();
        }

        // ------------------------------------------------------------------
        // Busy state + step log (callable from any thread)
        // ------------------------------------------------------------------

        private void ResetValidation()
        {
            _candidate = null;
            BtnReload.IsEnabled = false;
            TxtStatus.Text = string.Empty;
        }

        private void SetBusy(bool busy)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => SetBusy(busy)));
                return;
            }

            _busy = busy;
            BusyProgress.Visibility = busy ? Visibility.Visible : Visibility.Collapsed;
            BtnBrowse.IsEnabled = !busy;
            BtnReload.IsEnabled = !busy && _candidate != null;
            BtnCancel.IsEnabled = !busy;
            RbOffline.IsEnabled = !busy;
        }

        private void AppendLine(string glyph, string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => AppendLine(glyph, message)));
                return;
            }

            if (!string.IsNullOrEmpty(TxtStatus.Text)) TxtStatus.AppendText(Environment.NewLine);
            TxtStatus.AppendText($"{glyph} {message}");
            TxtStatus.ScrollToEnd();
        }

        /// <summary>Routine narration: the window plus a debug log line.</summary>
        private void LogStep(string message)
        {
            AppendLine(StepGlyph, message);
            SafeLog(l => l.LogDebug(LogPrefix + message));
        }

        private void LogSuccess(string message)
        {
            AppendLine(SuccessGlyph, message);
            SafeLog(l => l.LogDebug(LogPrefix + message));
        }

        /// <summary>An expected blocking condition; always written to the log.</summary>
        private void LogWarning(string message)
        {
            AppendLine(WarningGlyph, message);
            SafeLog(l => l.LogWarn(LogPrefix + message));
        }

        /// <summary>A real failure; always written to the log (with the exception when there is one).</summary>
        private void LogFailure(string message, Exception ex = null)
        {
            try { AppendLine(FailureGlyph, message); }
            catch { /* the window may already be gone */ }
            SafeLog(l => l.LogError(LogPrefix + message, ex?.ToString()));
        }

        private void SafeLog(Action<ILogger> write)
        {
            try
            {
                ILogger logger = _context.Logger;
                if (logger != null) write(logger);
            }
            catch
            {
                // Logging must never break the window.
            }
        }

        // ------------------------------------------------------------------
        // Downloads folder
        // ------------------------------------------------------------------

        [DllImport("shell32.dll")]
        private static extern int SHGetKnownFolderPath(ref Guid rfid, uint dwFlags, IntPtr hToken, out IntPtr pszPath);

        /// <summary>The user's Downloads folder (it may be redirected), or %USERPROFILE%\Downloads.</summary>
        private static string GetDownloadsFolder()
        {
            try
            {
                Guid id = FolderIdDownloads;
                if (SHGetKnownFolderPath(ref id, 0, IntPtr.Zero, out IntPtr pathPtr) == 0)
                {
                    try { return Marshal.PtrToStringUni(pathPtr); }
                    finally { Marshal.FreeCoTaskMem(pathPtr); }
                }
            }
            catch
            {
                // Fall back to the profile folder below.
            }
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
        }
    }
}
