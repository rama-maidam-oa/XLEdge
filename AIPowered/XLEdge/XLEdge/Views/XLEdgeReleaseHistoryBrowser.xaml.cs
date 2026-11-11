using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Threading.Tasks;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Views
{
    /// <summary>
    /// The Version History window: every core release adopted on this machine (ReleaseHistory.json, reconciled against
    /// AddinCore\Versions\ first so a deleted folder is never offered). Load returns the chosen release in <see cref="Chosen"/>
    /// (older versions may be loaded on purpose); Remove deletes a release that is not running.
    /// </summary>
    internal partial class XLEdgeReleaseHistoryBrowser : Window
    {
        private const string LogPrefix = "XLEdgeReleaseHistoryBrowser: ";

        private readonly IXLEdgeContext _context;
        private List<ReleaseRow> _rows = new List<ReleaseRow>();

        public XLEdgeReleaseHistoryBrowser(IXLEdgeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            Loaded += XLEdgeReleaseHistoryBrowser_Loaded;
        }

        /// <summary>The rows shown, newest first.</summary>
        public IReadOnlyList<ReleaseRow> Rows => _rows;

        /// <summary>The release to load once Load was clicked; null until then.</summary>
        public ResolvedRelease Chosen { get; private set; }

        private async void XLEdgeReleaseHistoryBrowser_Loaded(object sender, RoutedEventArgs e)
        {
            Loaded -= XLEdgeReleaseHistoryBrowser_Loaded;
            GridReleases.IsEnabled = false;
            TxtStatus.Text = "Loading version history...";
            TxtLoadedStatus.Text = LoadedStatusText(null);

            try
            {
                List<ReleaseEntry> entries = await Task.Run(() =>
                    ReleaseHistoryStore.Reconcile(_context.Paths.ReleaseHistoryFile, _context.Paths.VersionsPath));

                if (!IsLoaded)
                {
                    return;
                }

                PopulateEntries(entries);
            }
            catch (Exception ex)
            {
                if (!IsLoaded)
                {
                    return;
                }

                ShowLoadFailure(ex);
            }
            finally
            {
                if (IsLoaded)
                {
                    GridReleases.IsEnabled = true;
                }
            }
        }

        private void LoadEntries()
        {
            try
            {
                PopulateEntries(ReleaseHistoryStore.Reconcile(_context.Paths.ReleaseHistoryFile, _context.Paths.VersionsPath));
            }
            catch (Exception ex)
            {
                ShowLoadFailure(ex);
            }
        }

        private void PopulateEntries(List<ReleaseEntry> entries)
        {
            _rows = ReleaseHistoryRows.Build(entries, _context);
            GridReleases.ItemsSource = _rows;

            if (_rows.Count == 0)
            {
                TxtStatus.Text = "No version recorded yet.";
            }
            else
            {
                TxtStatus.Inlines.Clear();
                TxtStatus.Inlines.Add(new Run("Select a version, then click "));
                TxtStatus.Inlines.Add(new Run("Load or Remove") { FontWeight = FontWeights.SemiBold });
                TxtStatus.Inlines.Add(new Run("."));
            }

            ReleaseRow loaded = _rows.FirstOrDefault(r => r.IsCurrentlyLoaded);
            TxtLoadedStatus.Text = LoadedStatusText(loaded);
            if (loaded != null)
            {
                GridReleases.SelectedItem = loaded;
                GridReleases.ScrollIntoView(loaded);
            }
        }

        private void ShowLoadFailure(Exception ex)
        {
            SafeLog(l => l.LogError(LogPrefix + "reading the version history failed.", ex.ToString()));
            _rows = new List<ReleaseRow>();
            GridReleases.ItemsSource = _rows;
            TxtStatus.Text = $"The version history could not be read: {ex.Message}";
            TxtLoadedStatus.Text = LoadedStatusText(null);
        }

        private string LoadedStatusText(ReleaseRow loaded)
        {
            if (loaded != null) return $"Currently loaded: version {VersionText.Display(loaded.Version)}, version date {loaded.ReleaseDateDisplay}.";
            string version = _context.Version;
            // Not necessarily an error: the running release's folder may have been reconciled away.
            return string.IsNullOrWhiteSpace(version)
                ? "No version is currently loaded."
                : $"Currently loaded: version {VersionText.Display(version)} (not found in the version history).";
        }

        private void GridReleases_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var row = GridReleases.SelectedItem as ReleaseRow;
            bool selected = row != null;
            BtnLoad.IsEnabled = selected;
            BtnRemove.IsEnabled = selected;
        }

        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(GridReleases.SelectedItem is ReleaseRow row)) return;

                if (row.IsCurrentlyLoaded)
                {
                    MessageBox.Show(this, $"Version {VersionText.Display(row.Version)} is already running in this Excel session.", "Version Already Running",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // No version gate: loading an older release on purpose is what this window is for.
                Chosen = new ResolvedRelease
                {
                    Version = row.Version,
                    ReleaseDate = row.ReleaseDate,
                    FolderName = row.FolderName,
                    ContractVersion = row.ContractVersion,
                };
                CloseWith(true);
            }
            catch (Exception ex)
            {
                SafeLog(l => l.LogError(LogPrefix + "Load failed.", ex.ToString()));
            }
        }

        private void BtnRemove_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                RemoveSelected();
            }
            catch (Exception ex)
            {
                SafeLog(l => l.LogError(LogPrefix + "Remove failed.", ex.ToString()));
            }
        }

        private void RemoveSelected()
        {
            if (!(GridReleases.SelectedItem is ReleaseRow row)) return;

            if (row.IsCurrentlyLoaded)
            {
                MessageBox.Show(this,
                    "The currently loaded version is active and cannot be removed. Load another version first, then remove this one.",
                    "Cannot Remove Active Version", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MessageBoxResult confirmation = MessageBox.Show(this,
                $"Remove version {VersionText.Display(row.Version)} dated {row.ReleaseDateDisplay} from this computer?\n\n" +
                "Its files and its entry in this list are deleted from this computer only. " +
                "Nothing is removed from the server: if the server still offers this version, " +
                "Reload Add-in > Online > Check for Updates lists it again and it can be downloaded again.",
                "Remove Version", MessageBoxButton.YesNo, MessageBoxImage.Warning, MessageBoxResult.No);
            if (confirmation != MessageBoxResult.Yes) return;

            ReleaseDeleteResult result;
            Exception error;
            try
            {
                result = ReleaseHistoryStore.Delete(_context.Paths.ReleaseHistoryFile, _context.Paths.VersionsPath, row.FolderName,
                    _context.ActiveFolderName, out error);
            }
            catch (Exception ex)
            {
                SafeLog(l => l.LogError(LogPrefix + $"failed to remove version '{VersionText.Display(row.Version)}'.", ex.ToString()));
                MessageBox.Show(this, $"The version could not be removed.{Environment.NewLine}{Environment.NewLine}{ex.Message}",
                    "Remove Version Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            switch (result)
            {
                case ReleaseDeleteResult.Deleted:
                    SafeLog(l => l.LogDebug(LogPrefix + $"removed version '{VersionText.Display(row.Version)}' folder '{row.FolderName}'."));
                    LoadEntries();
                    break;

                case ReleaseDeleteResult.ActiveRelease:
                    MessageBox.Show(this, "The selected version is currently active and cannot be removed. Load another version first.",
                        "Cannot Remove Active Version", MessageBoxButton.OK, MessageBoxImage.Warning);
                    break;

                case ReleaseDeleteResult.FolderDeleteFailed:
                    SafeLog(l => l.LogError(LogPrefix + $"could not remove folder '{row.FolderName}'.", error?.ToString()));
                    MessageBox.Show(this,
                        $"The version could not be removed because its files could not be deleted.{Environment.NewLine}{Environment.NewLine}{error?.Message}",
                        "Remove Version Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;

                case ReleaseDeleteResult.CatalogWriteFailed:
                    SafeLog(l => l.LogError(LogPrefix + $"folder '{row.FolderName}' was deleted but the version history could not be updated.", error?.ToString()));
                    MessageBox.Show(this,
                        $"The version files were deleted, but the version history could not be updated.{Environment.NewLine}{Environment.NewLine}{error?.Message}",
                        "History Update Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    LoadEntries();
                    break;

                case ReleaseDeleteResult.InvalidFolderName:
                    SafeLog(l => l.LogWarn(LogPrefix + $"refused unsafe version folder '{row.FolderName}'."));
                    MessageBox.Show(this, "The selected version has an invalid local folder name and was not removed.",
                        "Remove Version Failed", MessageBoxButton.OK, MessageBoxImage.Error);
                    break;

                case ReleaseDeleteResult.NotFound:
                    MessageBox.Show(this, "That version is no longer present in the local history. The list will be refreshed.",
                        "Version Not Found", MessageBoxButton.OK, MessageBoxImage.Information);
                    LoadEntries();
                    break;
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
            if (IsLoaded) Close();
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
    }
}
