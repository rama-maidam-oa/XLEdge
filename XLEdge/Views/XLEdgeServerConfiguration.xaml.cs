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

        public XLEdgeServerConfiguration()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            urlInstances = new ObservableCollection<UrlInstance>();
            dgInstances.ItemsSource = urlInstances;

            LoadConfiguration();
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
            SaveConfiguration();
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
                        // that with a falsely reassuring message. The instance is still removed from
                        // this session's grid either way; only the on-disk persistence is affected.
                        if (SaveConfiguration())
                        {
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
                txtStatus.Text = message;
                txtStatus.Style = (Style)FindResource(GetStatusStyleKey(messageType));
                StatusBorder.Style = (Style)FindResource(GetStatusBorderStyleKey(messageType));
                AdjustWindowWidthToContent();
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