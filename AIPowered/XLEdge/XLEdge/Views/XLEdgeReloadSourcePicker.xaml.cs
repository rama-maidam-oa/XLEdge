using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Online;
using XLEdge.Shared;

namespace XLEdge.Views
{
    /// <summary>
    /// The Reload window: picks where the next core version comes from.
    /// Offline: a folder holding a manifest.json + zip pair, checked by <see cref="OfflineReleaseScanner"/>; on Reload the dialog
    /// returns true with <see cref="ScanResult"/> and the shell stages it after the old core is unloaded.
    /// Online (Stage 6, logged in only): the ORBIT server's XLEdge releases (<see cref="IOnlineReleaseSource"/>), classified against
    /// this machine's release history; on Reload every ticked release is downloaded and catalogued (UpdateBootstrapper
    /// .ExtractAndCatalog: checksum, contract, folder) before anything is unloaded, and the dialog returns true with
    /// <see cref="OnlineRelease"/> = the newest one that succeeded (GLSense's multi-release design). A failed one does not undo the
    /// others; if none succeeded the window stays open with the errors.
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

        internal const string LoggedOutToolTip = "Log in to XLEdge to check for online updates.";
        internal const string OnlineToolTipText = "Download XLEdge versions from the XLEdge server you are logged in to.";

        private readonly IXLEdgeContext _context;
        private readonly LoginInfo _login;
        private readonly Func<LoginInfo, IOnlineReleaseSource> _sourceFactory;
        private OfflineScanResult _candidate;
        private bool _busy;
        private bool _isClosed;

        /// <param name="login">The core's login (GetLoginInfo); Online is enabled only when it is logged in.</param>
        /// <param name="sourceFactory">Where online releases come from; the ORBIT server by default (tests pass a stand-in).</param>
        public XLEdgeReloadSourcePicker(IXLEdgeContext context, LoginInfo login = null, Func<LoginInfo, IOnlineReleaseSource> sourceFactory = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _login = login;
            _sourceFactory = sourceFactory ?? (l => new OnlineReleaseClient(l, context.Logger));
            InitializeComponent();
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            bool online = login != null && login.IsLoggedIn && !string.IsNullOrWhiteSpace(login.LoginUrl);
            RbOnline.IsEnabled = online;
            RbOnline.ToolTip = online ? OnlineToolTipText : LoggedOutToolTip;
            LstOnline.ItemsSource = OnlineRows;
        }

        /// <summary>Whether the Online source can be chosen (logged in to XLEdge).</summary>
        public bool IsOnlineEnabled => RbOnline.IsEnabled;

        /// <summary>What the Online choice says (why it is disabled, when it is).</summary>
        public string OnlineToolTip => RbOnline.ToolTip as string;

        public bool IsReloadEnabled => BtnReload.IsEnabled;

        /// <summary>The step log shown in the window.</summary>
        public string StatusText => TxtStatus.Text;

        /// <summary>Not closed yet.</summary>
        public bool IsVisibleOrOpen => !_isClosed;

        /// <summary>The server's releases from the last Check (Online).</summary>
        public ObservableCollection<OnlineReleaseRow> OnlineRows { get; } = new ObservableCollection<OnlineReleaseRow>();

        /// <summary>Online: the newest ticked release that was downloaded and catalogued - the one to reload into; null otherwise.</summary>
        public ResolvedRelease OnlineRelease { get; private set; }

        private bool IsOnlineSelected => RbOnline.IsChecked == true;

        /// <summary>Chooses Online (as a click on it does).</summary>
        public void SelectOnline() => RbOnline.IsChecked = true;

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
        // Online flow
        // ------------------------------------------------------------------

        private void Mode_Checked(object sender, RoutedEventArgs e)
        {
            if (OnlinePanel == null || OfflinePanel == null) return; // during InitializeComponent
            OnlinePanel.Visibility = IsOnlineSelected ? Visibility.Visible : Visibility.Collapsed;
            OfflinePanel.Visibility = IsOnlineSelected ? Visibility.Collapsed : Visibility.Visible;
            UpdateReloadEnabled();
        }

        private async void BtnCheck_Click(object sender, RoutedEventArgs e)
        {
            try { await CheckForUpdatesAsync(); }
            catch (Exception ex) { LogFailure($"Checking for updates failed: {ex.Message}", ex); SetBusy(false); }
        }

        /// <summary>Asks the server for its releases and lists them, classified against this machine's release history.</summary>
        internal async Task CheckForUpdatesAsync()
        {
            OnUi(() =>
            {
                foreach (OnlineReleaseRow row in OnlineRows) row.PropertyChanged -= Row_PropertyChanged;
                OnlineRows.Clear();
                TxtStatus.Text = string.Empty;
            });
            SetBusy(true);
            try
            {
                LogStep("Asking the XLEdge server for its versions...");
                IReadOnlyList<VersionInfo> releases;
                using (IOnlineReleaseSource source = _sourceFactory(_login))
                    releases = await source.GetReleasesAsync(CancellationToken.None);
                ReleaseHistoryStore.TryReadAll(_context.Paths.ReleaseHistoryFile, out List<ReleaseEntry> local);
                List<OnlineReleaseRow> rows = OnlineReleaseClassifier.Classify(releases, local, _context.Version, _context.ReleaseDate);
                // After an await this may run on a pool thread (no synchronization context in the add-in): the list is bound.
                OnUi(() =>
                {
                    foreach (OnlineReleaseRow row in rows)
                    {
                        row.PropertyChanged += Row_PropertyChanged;
                        OnlineRows.Add(row);
                    }
                });
                int fresh = rows.Count(r => r.IsSelectable);
                if (rows.Count == 0) LogWarning("The server has no XLEdge versions.");
                else if (fresh == 0) LogStep($"The server has {rows.Count} version(s); you already have all of them.");
                else LogSuccess($"The server has {rows.Count} version(s), {fresh} new. Tick the ones to download; the newest one ticked is loaded.");
            }
            catch (OnlineReleaseException ex)
            {
                LogFailure(ex.Message);
            }
            catch (Exception ex)
            {
                LogFailure($"Checking for updates failed: {ex.Message}", ex);
            }
            finally
            {
                SetBusy(false);
            }
        }

        private void Row_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e) => UpdateReloadEnabled();

        /// <summary>
        /// Downloads and catalogues every ticked release, oldest first; then closes with <see cref="OnlineRelease"/> = the newest
        /// one that succeeded. Nothing is unloaded here: the shell reloads after the dialog closes.
        /// </summary>
        internal async Task ReloadOnlineAsync()
        {
            List<OnlineReleaseRow> chosen = null;
            OnUi(() => chosen = OnlineRows.Where(r => r.IsChecked).OrderBy(r => OnlineReleaseClassifier.ReleaseDateOrMin(r.ReleaseDate)).ToList());
            if (chosen.Count == 0) return;
            var done = new List<(OnlineReleaseRow Row, ResolvedRelease Release)>();
            var failed = new List<string>();
            SetBusy(true);
            try
            {
                var bootstrapper = new UpdateBootstrapper();
                using (IOnlineReleaseSource source = _sourceFactory(_login))
                {
                    foreach (OnlineReleaseRow row in chosen)
                    {
                        LogStep($"Downloading {row.Version} (released {row.ReleaseDate})...");
                        try
                        {
                            (string manifest, byte[] zip) = await source.DownloadAsync(row.FolderName, CancellationToken.None);
                            ResolvedRelease release = await Task.Run(() => bootstrapper.ExtractAndCatalog(_context, manifest, zip, "Online"));
                            if (release == null)
                            {
                                failed.Add(row.Version);
                                LogFailure($"{row.Version} was refused (checksum, contract or folder - see the XLEdge log).");
                            }
                            else
                            {
                                done.Add((row, release));
                                LogSuccess($"{row.Version} downloaded and checked.");
                            }
                        }
                        catch (OnlineReleaseException ex)
                        {
                            failed.Add(row.Version);
                            LogFailure($"{row.Version}: {ex.Message}");
                        }
                        catch (Exception ex)
                        {
                            failed.Add(row.Version);
                            LogFailure($"{row.Version} could not be downloaded: {ex.Message}", ex);
                        }
                    }
                }
            }
            finally
            {
                SetBusy(false);
            }

            if (done.Count == 0)
            {
                LogWarning("No version was downloaded - nothing to reload.");
                return;
            }
            if (failed.Count > 0) LogWarning($"Not downloaded: {string.Join(", ", failed)}. The others are kept.");
            ResolvedRelease newest = done.OrderByDescending(d => OnlineReleaseClassifier.ReleaseDateOrMin(d.Row.ReleaseDate)).First().Release;
            SafeLog(l => l.LogDebug(LogPrefix + $"online version {newest.Version} ({newest.FolderName}) chosen - proceeding with reload."));
            OnUi(() =>
            {
                OnlineRelease = newest;
                SelectedSource = "Online";
                CloseWith(true);
            });
        }

        /// <summary>Runs <paramref name="action"/> on the window's thread (code after an await may be on a pool thread).</summary>
        private void OnUi(Action action)
        {
            if (Dispatcher.CheckAccess()) action();
            else Dispatcher.Invoke(action);
        }

        private void UpdateReloadEnabled()
        {
            if (BtnReload == null) return;
            if (!Dispatcher.CheckAccess()) { Dispatcher.Invoke(new Action(UpdateReloadEnabled)); return; }
            BtnReload.IsEnabled = !_busy && (IsOnlineSelected ? OnlineRows.Any(r => r.IsChecked) : _candidate != null);
        }

        // ------------------------------------------------------------------
        // Reload / Cancel
        // ------------------------------------------------------------------

        private async void BtnReload_Click(object sender, RoutedEventArgs e)
        {
            if (_busy) return;
            if (IsOnlineSelected)
            {
                try { await ReloadOnlineAsync(); }
                catch (Exception ex) { LogFailure($"Failed to start the reload: {ex.Message}", ex); SetBusy(false); }
                return;
            }
            if (_candidate == null) return;
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
            UpdateReloadEnabled();
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
            BtnCheck.IsEnabled = !busy;
            LstOnline.IsEnabled = !busy;
            UpdateReloadEnabled();
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
