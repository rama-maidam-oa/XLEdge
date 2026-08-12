using XLEdge.Helpers;
using XLEdge.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using System.Xml.Linq;

namespace XLEdge.Views
{
    /// <summary>
    /// Interaction logic for XLEdgeServerConfiguration.xaml
    /// </summary>
    public partial class XLEdgeServerConfiguration : DpiAwareWindow
    {
        private enum StatusMessageType
        {
            Success,
            Error,
            Warning,
            Info
        }

        private sealed class UrlInstanceSnapshot
        {
            public string Name { get; set; }
            public string Address { get; set; }
            public bool IsDefault { get; set; }
        }

        private static readonly object CachedConfigurationLock = new object();
        private static List<UrlInstanceSnapshot> CachedConfiguration;

        private readonly string xmlFilePath = XLEdgeAppPaths.TempUrlsPath;
        private readonly ObservableCollection<UrlInstance> urlInstances;
        private bool isInternalUpdate = false;
        private string persistedDefaultName;

        // Auto-hides the status message/border 10 seconds after it's shown - previously it stayed
        // on screen indefinitely until the next button click, which is confusing since it doesn't
        // reflect anything changing after that.
        private readonly DispatcherTimer statusAutoHideTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(10) };

        public XLEdgeServerConfiguration()
        {
            InitializeComponent();

            urlInstances = new ObservableCollection<UrlInstance>();
            dgInstances.ItemsSource = urlInstances;

            statusAutoHideTimer.Tick += StatusAutoHideTimer_Tick;
            this.Closed += (s, e) => statusAutoHideTimer.Stop();

            LoadConfiguration();
        }

        private void StatusAutoHideTimer_Tick(object sender, EventArgs e)
        {
            HideStatus();
        }

        private void HideStatus()
        {
            statusAutoHideTimer.Stop();
            StatusBorder.Visibility = Visibility.Collapsed;
            txtStatus.Text = string.Empty;
            AdjustWindowWidthToContent();
        }

        // Replaces EnhancedDragDropHelper.EnableWindowDrag(this) now that the window has a real
        // title bar (WindowStyle="SingleBorderWindow" + ExtendsContentIntoTitleBar="True").
        private void TitleBar_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                if (e.LeftButton == System.Windows.Input.MouseButtonState.Pressed)
                {
                    this.DragMove();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "XLEdgeServerConfiguration.TitleBar_MouseLeftButtonDown");
            }
        }

        private void XLEdgeServerConfiguration_Loaded(object sender, RoutedEventArgs e)
        {
            // Force refresh of all bindings
            dgInstances.Items.Refresh();
        }

        private void DgInstances_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            // Check if we're editing the Default column
            if (e.Column.Header.ToString() == "Default")
            {
                var instance = e.Row.Item as UrlInstance;
                e.Cancel = true;

                if (instance != null && instance.IsDefault)
                {
                    UpdateStatus($"'{instance.Name}' is already the default server. Use 'Set as Default' button to change.", StatusMessageType.Warning);
                }
                else
                {
                    UpdateStatus("Please use the 'Set as Default' button to change the default server.", StatusMessageType.Warning);
                }
            }
        }

        private void LoadConfiguration()
        {
            try
            {
                EnsureConfigFilePath();
                if (!File.Exists(xmlFilePath))
                {
                    CreateDefaultConfig();
                }

                var snapshots = GetConfigurationSnapshots();
                urlInstances.Clear();

                foreach (var snapshot in snapshots)
                {
                    var instance = new UrlInstance
                    {
                        Name = snapshot.Name ?? "",
                        Address = snapshot.Address?.Trim() ?? "",
                        IsDefault = snapshot.IsDefault
                    };

                    instance.PropertyChanged += UrlInstance_PropertyChanged;
                    urlInstances.Add(instance);
                }

                isInternalUpdate = true;
                try
                {
                    EnsureSingleDefault();
                    EnsureDefaultExists();
                    ReorderInstancesWithDefaultFirst();
                }
                finally
                {
                    isInternalUpdate = false;
                }

                UpdateCachedConfiguration();
                persistedDefaultName = urlInstances.FirstOrDefault(u => u.IsDefault)?.Name;

                // Ensure the UI selection (checkbox) reflects the persisted default on load
                foreach (var instance in urlInstances)
                {
                    instance.IsSelected = instance.IsDefault;
                }

                UpdateStatus($"Configuration loaded successfully. {urlInstances.Count} instances found.", StatusMessageType.Success);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(LoadConfiguration));
                UpdateStatus($"Error loading configuration: {ex.Message}", StatusMessageType.Error);
            }
        }

        private void UrlInstance_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(UrlInstance.IsDefault) && !isInternalUpdate)
            {
                var changedInstance = sender as UrlInstance;
                if (changedInstance != null && changedInstance.IsDefault)
                {
                    isInternalUpdate = true;
                    try
                    {
                        SetDefaultInstance(changedInstance);
                    }
                    finally
                    {
                        isInternalUpdate = false;
                    }

                    UpdateStatus($"'{changedInstance.Name}' has been set as the default server. Click Save to persist the change.", StatusMessageType.Info);
                }
            }
        }

        private void EnsureSingleDefault()
        {
            var defaultInstances = urlInstances.Where(u => u.IsDefault).ToList();
            if (defaultInstances.Count > 1)
            {
                for (int i = 1; i < defaultInstances.Count; i++)
                {
                    defaultInstances[i].IsDefault = false;
                }
            }
        }

        private void EnsureDefaultExists()
        {
            if (urlInstances.Any() && !urlInstances.Any(u => u.IsDefault))
            {
                urlInstances.First().IsDefault = true;
            }
        }

        private void CreateDefaultConfig()
        {
            try
            {
                EnsureConfigFilePath();
                var emptyConfig = new XDocument(new XElement("ORBIT"));
                emptyConfig.Save(xmlFilePath);
                UpdateStatus("New configuration file created.", StatusMessageType.Success);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(CreateDefaultConfig));
                UpdateStatus($"Error creating configuration file: {ex.Message}", StatusMessageType.Error);
            }
        }

        /// <summary>
        /// Returns true only when the configuration was actually written to disk, so callers other
        /// than the Save button itself (Set as Default, Delete) can tell a real persisted save apart
        /// from a validation failure - and avoid showing a falsely reassuring "saved" message when
        /// SaveConfiguration already surfaced the real error via UpdateStatus.
        /// </summary>
        private bool SaveConfiguration()
        {
            try
            {
                EnsureConfigFilePath();

                // Validate duplicate names
                var duplicateNames = urlInstances
                    .Where(u => !string.IsNullOrWhiteSpace(u.Name))
                    .GroupBy(u => u.Name.ToLower())
                    .Where(g => g.Count() > 1)
                    .Select(g => g.Key)
                    .ToList();

                if (duplicateNames.Any())
                {
                    UpdateStatus($"Duplicate Instance Names found: {string.Join(", ", duplicateNames)}", StatusMessageType.Error);
                    return false;
                }

                // Validate URLs
                var invalidUrls = urlInstances
                    .Where(u => !string.IsNullOrWhiteSpace(u.Address) &&
                               !u.Address.StartsWith("http://") &&
                               !u.Address.StartsWith("https://"))
                    .ToList();

                if (invalidUrls.Any())
                {
                    UpdateStatus($"Invalid URL format found. URLs must start with http:// or https://", StatusMessageType.Error);
                    return false;
                }

                // Validate mandatory fields
                var missingMandatory = urlInstances
                    .Where(u => string.IsNullOrWhiteSpace(u.Name) || string.IsNullOrWhiteSpace(u.Address))
                    .ToList();

                if (missingMandatory.Any())
                {
                    UpdateStatus($"Instance Name and URL Address are mandatory for all entries.", StatusMessageType.Error);
                    return false;
                }

                // Ensure only one default exists before saving
                isInternalUpdate = true;
                try
                {
                    EnsureSingleDefault();
                    EnsureDefaultExists();
                    ReorderInstancesWithDefaultFirst();
                }
                finally
                {
                    isInternalUpdate = false;
                }

                var validInstances = urlInstances.Where(u => !string.IsNullOrWhiteSpace(u.Name)).ToList();

                var doc = new XDocument(
                    new XElement("ORBIT",
                        validInstances.Select(instance =>
                            new XElement("URL",
                                new XElement("Name", instance.Name?.Trim() ?? ""),
                                new XElement("Address", instance.Address?.Trim().TrimEnd('/') ?? ""),
                                new XElement("DefaultURL", instance.IsDefault.ToString())
                            )
                        )
                    )
                );

                Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath));
                doc.Save(xmlFilePath);
                UpdateCachedConfiguration();
                persistedDefaultName = urlInstances.FirstOrDefault(u => u.IsDefault)?.Name;

                UpdateStatus("Configuration saved successfully.", StatusMessageType.Success);
                return true;
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(SaveConfiguration));
                UpdateStatus($"Error saving configuration: {ex.Message}", StatusMessageType.Error);
                return false;
            }
        }

        private void AutoSaveConfiguration()
        {
            UpdateCachedConfiguration();
        }

        private void DgInstances_InitializingNewItem(object sender, InitializingNewItemEventArgs e)
        {
            if (e.NewItem is UrlInstance instance)
            {
                instance.PropertyChanged += UrlInstance_PropertyChanged;

                // Ensure proper initialization with empty strings
                instance.Name = string.Empty;
                instance.Address = string.Empty;

                // Force the binding to update
                instance.OnPropertyChanged(nameof(UrlInstance.HasAnyData));
            }
        }

        private void DgInstances_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                // Ignore placeholder/empty rows
                if (!selectedInstance.HasAnyData)
                {
                    return;
                }

                // Clicking a real row means the user has moved on from whatever the last status
                // message was about - dismiss it immediately rather than leaving it up for the
                // rest of its 10-second auto-hide window.
                HideStatus();

                // Mark the selected row for UI only
                foreach (var instance in urlInstances)
                {
                    instance.IsSelected = instance == selectedInstance;
                }
            }
        }

        private void BtnGo_Click(object sender, RoutedEventArgs e)
        {
            UrlInstance selectedInstance = dgInstances.SelectedItem as UrlInstance ?? urlInstances.FirstOrDefault(u => u.IsDefault);

            if (selectedInstance == null)
            {
                UpdateStatus("Please select an instance or set a default server before continuing.", StatusMessageType.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(selectedInstance.Address))
            {
                UpdateStatus("The selected instance has no URL address configured.", StatusMessageType.Warning);
                return;
            }

            if (AddinModule.CurrentInstance != null && AddinModule.CurrentInstance.NavigateReportsToAddress(selectedInstance.Name, selectedInstance.Address))
            {
                UpdateStatus($"Navigating to {selectedInstance.Name}...", StatusMessageType.Success);
                Close();
                return;
            }

            UpdateStatus("Unable to open the selected address.", StatusMessageType.Error);
        }

        private void BtnSetDefault_Click(object sender, RoutedEventArgs e)
        {
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                var currentPersistedDefault = urlInstances.FirstOrDefault(u => u.IsDefault);

                if (currentPersistedDefault == selectedInstance &&
                    string.Equals(persistedDefaultName, selectedInstance.Name, StringComparison.OrdinalIgnoreCase))
                {
                    UpdateStatus($"'{selectedInstance.Name}' is already the default server. No changes made.", StatusMessageType.Warning);
                    return;
                }

                SetDefaultInstance(selectedInstance);

                // Set as Default now persists immediately instead of requiring a separate Save click -
                // if SaveConfiguration fails validation (duplicate name, bad URL, missing field on some
                // other row), it already shows the real error via UpdateStatus, so don't overwrite that
                // with a falsely reassuring "saved" message.
                if (SaveConfiguration())
                {
                    UpdateStatus($"Default set to '{selectedInstance.Name}' and saved.", StatusMessageType.Success);
                }
            }
            else
            {
                UpdateStatus("Please select an instance to set as default.", StatusMessageType.Error);
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!HasUnsavedChanges())
            {
                UpdateStatus("Nothing to save.", StatusMessageType.Info);
                return;
            }

            SaveConfiguration();
        }

        /// <summary>
        /// Compares the grid's current, in-memory rows against the last-persisted snapshot
        /// (CachedConfiguration, kept in sync with disk by UpdateCachedConfiguration after every
        /// successful save/load) to tell whether Save would actually write anything different.
        /// Order-independent (grid rows can reorder via ReorderInstancesWithDefaultFirst without
        /// that alone counting as a real change) and ignores the empty add-row placeholder, the
        /// same way SaveConfiguration's own validInstances filter does.
        /// </summary>
        private bool HasUnsavedChanges()
        {
            List<UrlInstanceSnapshot> baseline;
            lock (CachedConfigurationLock)
            {
                baseline = CachedConfiguration ?? new List<UrlInstanceSnapshot>();
            }

            var current = urlInstances
                .Where(u => !string.IsNullOrWhiteSpace(u.Name))
                .Select(u => $"{u.Name?.Trim().ToLowerInvariant()}|{u.Address?.Trim().TrimEnd('/')}|{u.IsDefault}")
                .OrderBy(s => s, StringComparer.Ordinal)
                .ToList();

            var currentBaseline = baseline
                .Where(s => !string.IsNullOrWhiteSpace(s.Name))
                .Select(s => $"{s.Name?.Trim().ToLowerInvariant()}|{s.Address?.Trim().TrimEnd('/')}|{s.IsDefault}")
                .OrderBy(s => s, StringComparer.Ordinal)
                .ToList();

            return !current.SequenceEqual(currentBaseline);
        }

        private void BtnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (dgInstances.SelectedItem is UrlInstance selectedInstance)
            {
                if (string.IsNullOrWhiteSpace(selectedInstance.Name) && string.IsNullOrWhiteSpace(selectedInstance.Address))
                {
                    UpdateStatus("Selected row is empty. Enter an Instance Name and URL Address, or select a different row to delete.", StatusMessageType.Warning);
                    return;
                }

                string instanceName = string.IsNullOrWhiteSpace(selectedInstance.Name)
                                        ? "this instance"
                                        : $"'{selectedInstance.Name}'";

                bool wasDefault = selectedInstance.IsDefault;

                AppOverlayControl.ShowConfirm(
                    $"Are you sure you want to delete instance {instanceName}?",
                    yesAction: () =>
                    {
                        urlInstances.Remove(selectedInstance);

                        if (wasDefault && urlInstances.Any())
                        {
                            SetDefaultInstance(urlInstances.First());
                        }

                        // Delete now persists immediately instead of requiring a separate Save click -
                        // if SaveConfiguration fails validation (e.g. another row still has bad/missing
                        // data), it already shows the real error via UpdateStatus, so don't overwrite
                        // that with a falsely reassuring message, and don't reload either (the removal
                        // above still stands in-memory even on a failed save; reloading here would
                        // instead re-fetch the stale pre-delete snapshot, since UpdateCachedConfiguration
                        // is only reached on a successful save - see GetConfigurationSnapshots).
                        if (SaveConfiguration())
                        {
                            // Bug fix: deleting a NON-default row left the "Default" checkbox
                            // (bound to IsSelected, which DgInstances_SelectionChanged had just
                            // overwritten via WPF's automatic reselection of another row after the
                            // deleted one) out of sync with the real default - it only got
                            // re-synced above when the deleted row itself was the default. Reload
                            // from the just-persisted configuration instead of trusting this
                            // session's in-memory selection state - the same full refresh already
                            // done when the window is first opened - so the checkbox and row order
                            // (default first, else alphabetical - see ReorderInstancesWithDefaultFirst)
                            // always reflect what's actually in the XML, whichever row was deleted.
                            LoadConfiguration();
                            UpdateStatus($"Instance {instanceName} deleted and saved.", StatusMessageType.Info);
                        }
                        ResetRibbonIfLoggedOut();
                    },
                    noAction: () =>
                    {
                        UpdateStatus("Delete cancelled.", StatusMessageType.Info);
                        ResetRibbonIfLoggedOut();
                    }
                );
            }
            else
            {
                UpdateStatus("Please select an instance to delete.", StatusMessageType.Error);
            }
        }

        private void ResetRibbonIfLoggedOut()
        {
            try
            {
                if (XLEdge.AddinModule.CurrentInstance?.loginButtonVisibility() == true)
                {
                    XLEdge.AddinModule.CurrentInstance.RibbonInitialize();
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(ResetRibbonIfLoggedOut));
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void DgInstances_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                var column = e.Column.Header.ToString();

                if (column == "Instance Name" && e.EditingElement is TextBox textBox)
                {
                    var instance = e.Row.Item as UrlInstance;
                    if (instance != null)
                    {
                        string newValue = textBox.Text?.Trim();

                        if (string.IsNullOrWhiteSpace(newValue))
                        {
                            UpdateStatus("Instance Name is a mandatory field.", StatusMessageType.Error);
                            e.Cancel = true;
                            return;
                        }

                        var duplicate = urlInstances.Any(u => u != instance &&
                                                              u.Name != null &&
                                                              newValue != null &&
                                                              u.Name.ToLower() == newValue.ToLower());
                        if (duplicate)
                        {
                            UpdateStatus($"Instance Name '{newValue}' already exists. Please use a unique name.", StatusMessageType.Error);
                            e.Cancel = true;
                            return;
                        }

                        instance.Name = newValue;
                        AutoSaveConfiguration();
                        UpdateStatus("Changes made. Click 'Save' to persist.", StatusMessageType.Info);
                    }
                }
                else if (column == "URL Address" && e.EditingElement is TextBox urlTextBox)
                {
                    var instance = e.Row.Item as UrlInstance;
                    if (instance != null)
                    {
                        string url = urlTextBox.Text?.Trim();

                        if (string.IsNullOrWhiteSpace(url))
                        {
                            UpdateStatus("URL Address is a mandatory field.", StatusMessageType.Error);
                            e.Cancel = true;
                            return;
                        }

                        if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                        {
                            UpdateStatus("URL Address must start with http:// or https://", StatusMessageType.Error);
                            e.Cancel = true;
                            return;
                        }

                        instance.Address = url;
                        AutoSaveConfiguration();
                        UpdateStatus("Changes made. Click 'Save' to persist.", StatusMessageType.Info);
                    }
                }
            }
        }

        private void UpdateStatus(string message, bool isSuccess)
        {
            UpdateStatus(message, isSuccess ? StatusMessageType.Success : StatusMessageType.Error);
        }

        private void UpdateStatus(string message, StatusMessageType messageType)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                StatusBorder.Visibility = Visibility.Visible;
                txtStatus.Text = message;
                txtStatus.Style = (Style)FindResource(GetStatusStyleKey(messageType));
                StatusBorder.Style = (Style)FindResource(GetStatusBorderStyleKey(messageType));
                AdjustWindowWidthToContent();

                // Restart rather than just start: a second status shown before the first one's 10
                // seconds are up (e.g. two quick actions in a row) should get its own full 10
                // seconds, not inherit whatever's left on the previous message's countdown.
                statusAutoHideTimer.Stop();
                statusAutoHideTimer.Start();
            }));
        }

        private string GetStatusStyleKey(StatusMessageType messageType)
        {
            switch (messageType)
            {
                case StatusMessageType.Warning:
                    return "WarningMessage";
                case StatusMessageType.Info:
                    return "InfoMessage";
                case StatusMessageType.Success:
                    return "SuccessMessage";
                case StatusMessageType.Error:
                default:
                    return "ErrorMessage";
            }
        }

        private string GetStatusBorderStyleKey(StatusMessageType messageType)
        {
            switch (messageType)
            {
                case StatusMessageType.Warning:
                    return "WarningMessageBorder";
                case StatusMessageType.Info:
                    return "InfoMessageBorder";
                case StatusMessageType.Success:
                    return "SuccessMessageBorder";
                case StatusMessageType.Error:
                default:
                    return "ErrorMessageBorder";
            }
        }

        private void AdjustWindowWidthToContent()
        {
            var content = Content as FrameworkElement;
            if (content == null)
            {
                return;
            }

            UpdateLayout();
            content.Measure(new System.Windows.Size(double.PositiveInfinity, double.PositiveInfinity));

            double desiredWidth = content.DesiredSize.Width;
            double clampedWidth = Math.Max(MinWidth, Math.Min(MaxWidth, desiredWidth));

            if (!double.IsNaN(clampedWidth) && clampedWidth > 0)
            {
                Width = clampedWidth;
            }
        }

        private List<UrlInstanceSnapshot> GetConfigurationSnapshots()
        {
            lock (CachedConfigurationLock)
            {
                if (CachedConfiguration != null)
                {
                    return CachedConfiguration
                        .Select(instance => new UrlInstanceSnapshot
                        {
                            Name = instance.Name,
                            Address = instance.Address,
                            IsDefault = instance.IsDefault
                        })
                        .ToList();
                }
            }

            if (!File.Exists(xmlFilePath))
            {
                return new List<UrlInstanceSnapshot>();
            }

            var doc = XDocument.Load(xmlFilePath);
            return doc.Descendants("URL")
                .Select(urlElement => new UrlInstanceSnapshot
                {
                    Name = urlElement.Element("Name")?.Value ?? string.Empty,
                    Address = urlElement.Element("Address")?.Value?.Trim() ?? string.Empty,
                    IsDefault = bool.TryParse(urlElement.Element("DefaultURL")?.Value, out bool isDefault) && isDefault
                })
                .ToList();
        }

        private void UpdateCachedConfiguration()
        {
            lock (CachedConfigurationLock)
            {
                CachedConfiguration = urlInstances
                    .Select(instance => new UrlInstanceSnapshot
                    {
                        Name = instance.Name,
                        Address = instance.Address,
                        IsDefault = instance.IsDefault
                    })
                    .ToList();
            }
        }

        private void SetDefaultInstance(UrlInstance selectedInstance)
        {
            if (selectedInstance == null)
            {
                return;
            }

            isInternalUpdate = true;
            try
            {
                foreach (var instance in urlInstances)
                {
                    instance.IsDefault = instance == selectedInstance;
                }

                ReorderInstancesWithDefaultFirst();
                foreach (var instance in urlInstances)
                {
                    instance.IsSelected = instance.IsDefault;
                }
            }
            finally
            {
                isInternalUpdate = false;
            }
            UpdateCachedConfiguration();
        }

        private void ReorderInstancesWithDefaultFirst()
        {
            if (urlInstances.Count <= 1)
            {
                return;
            }

            var defaultInstance = urlInstances.FirstOrDefault(u => u.IsDefault);
            var ordered = urlInstances
                .Where(u => u != defaultInstance)
                .OrderBy(u => u.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (defaultInstance != null)
            {
                ordered.Insert(0, defaultInstance);
            }

            for (int targetIndex = 0; targetIndex < ordered.Count; targetIndex++)
            {
                int currentIndex = urlInstances.IndexOf(ordered[targetIndex]);
                if (currentIndex != targetIndex)
                {
                    urlInstances.Move(currentIndex, targetIndex);
                }
            }
        }

        private void EnsureConfigFilePath()
        {
            try
            {
                var directory = Path.GetDirectoryName(xmlFilePath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(EnsureConfigFilePath));
                UpdateStatus($"Error preparing configuration path: {ex.Message}", StatusMessageType.Error);
            }
        }
    }

    public class UrlInstance : INotifyPropertyChanged
    {
        private string _name = string.Empty;
        private string _address = string.Empty;
        private bool _isDefault;
        private bool _isSelected;

        public string Name
        {
            get => _name;
            set
            {
                if (_name != value)
                {
                    _name = value ?? string.Empty;
                    OnPropertyChanged(nameof(Name));
                    OnPropertyChanged(nameof(HasAnyData));
                }
            }
        }

        public string Address
        {
            get => _address;
            set
            {
                if (_address != value)
                {
                    _address = value ?? string.Empty;
                    OnPropertyChanged(nameof(Address));
                    OnPropertyChanged(nameof(HasAnyData));
                }
            }
        }

        public bool HasAnyData => !string.IsNullOrWhiteSpace(_name) || !string.IsNullOrWhiteSpace(_address);

        public bool IsDefault
        {
            get => _isDefault;
            set
            {
                if (_isDefault != value)
                {
                    _isDefault = value;
                    OnPropertyChanged(nameof(IsDefault));
                }
            }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}