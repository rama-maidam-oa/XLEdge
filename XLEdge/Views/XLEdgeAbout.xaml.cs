using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Xml.Linq;
using XLEdge.Helpers;
using XLEdge.Utilities;

namespace XLEdge.Views
{
#nullable enable
    /// <summary>
    /// Interaction logic for XLEdgeAbout.xaml
    /// </summary>
    public partial class XLEdgeAbout : DpiAwareWindow
    {
        private readonly ObservableCollection<InstanceCompatibility> instances;
        private readonly string xmlFilePath = XLEdgeAppPaths.TempUrlsPath;
        public XLEdgeAbout()
        {
            InitializeComponent();
            EnhancedDragDropHelper.EnableWindowDrag(this);

            DataContext = this;
            // Initialize the collection correctly
            instances = new ObservableCollection<InstanceCompatibility>();
            dgInstances.ItemsSource = instances;

            Loaded += AboutWindow_Loaded;
        }
        private async void AboutWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Start compatibility checking
            await CheckInstanceCompatibility();
        }
        private async Task CheckInstanceCompatibility()
        {
            try
            {
                // Ensure UI shows the indeterminate marquee before doing network work
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressPanel.Visibility = Visibility.Visible;
                    progressBar.IsIndeterminate = true;
                    txtProgress.Text = "Starting compatibility checks...";
                }, System.Windows.Threading.DispatcherPriority.Background);

                // Give the UI a tiny moment to render the marquee
                await Task.Yield();

                var urlInstances = LoadInstancesFromXml();
                int totalInstances = urlInstances.Count;

                // Clear existing data on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    instances.Clear();
                }, System.Windows.Threading.DispatcherPriority.Background);

                if (totalInstances == 0)
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        instances.Add(new InstanceCompatibility { Instance = "No instances configured", IsCompatible = false });
                        progressPanel.Visibility = Visibility.Collapsed;
                        progressBar.IsIndeterminate = false;
                        txtProgress.Text = "No instances configured";
                    }, System.Windows.Threading.DispatcherPriority.Background);
                    return;
                }

                int processed = 0;

                // Use indeterminate marquee while doing the checks (network I/O off the UI thread)
                foreach (var instance in urlInstances)
                {
                    // Update status text on UI thread before the network call so user sees it immediately
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        txtProgress.Text = $"Checking {instance.Name}... ({processed + 1}/{totalInstances})";
                    }, System.Windows.Threading.DispatcherPriority.Background);

                    // Run the potentially slow operation off the UI thread
                    bool isCompatible = false;
                    try
                    {
                        isCompatible = await Task.Run(async () => await CheckUrlCompatibility(instance.Address).ConfigureAwait(false)).ConfigureAwait(false);
                    }
                    catch (Exception ex)
                    {
                        // Unexpected per-instance error: log and continue (do not add a global "Error" row)
                        LogUtility.LogError($"Per-instance compatibility check failed for '{instance.Address}': {ex.Message}");
                        isCompatible = false;
                    }

                    processed++;

                    // Add result to collection on UI thread
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        instances.Add(new InstanceCompatibility { Instance = instance.Address, IsCompatible = isCompatible });
                    }, System.Windows.Threading.DispatcherPriority.Background);

                    // Small pause to allow UI to breathe and show updates
                    await Task.Delay(150).ConfigureAwait(false);
                }

                // Finish: stop marquee and hide panel
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressBar.IsIndeterminate = false;
                    progressPanel.Visibility = Visibility.Collapsed;
                    txtProgress.Text = "Compatibility check completed";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
            catch (Exception ex)
            {
                // Log the failure and update the UI, without adding an "Error" row to the grid.
                LogUtility.LogException(ex);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    progressPanel.Visibility = Visibility.Collapsed;
                    progressBar.IsIndeterminate = false;
                    txtProgress.Text = "Error checking instances (see logs)";
                }, System.Windows.Threading.DispatcherPriority.Background);
            }
        }

        private System.Collections.Generic.List<UrlInstance> LoadInstancesFromXml()
        {
            var instances1 = new System.Collections.Generic.List<UrlInstance>();

            try
            {
                if (!File.Exists(xmlFilePath))
                    return instances1;

                XDocument doc = XDocument.Load(xmlFilePath);

                foreach (var urlElement in doc.Descendants("URL"))
                {
                    instances1.Add(new UrlInstance
                    {
                        Name = urlElement.Element("Name")?.Value ?? "",
                        Address = urlElement.Element("Address")?.Value ?? "",
                        IsDefault = bool.Parse(urlElement.Element("DefaultURL")?.Value ?? "false")
                    });
                }
            }
            catch (Exception ex)
            {
                // Log error but don't throw - we'll show empty grid
                LogUtility.LogError($"Error loading instances: {ex.Message}");
            }

            return instances1;
        }
        private async Task<bool> CheckUrlCompatibility(string url)
        {
            try
            {
                var handler = new HttpClientHandler()
                {
                    SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
                    ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate
                };

                using var httpClient = new HttpClient(handler);
                httpClient.Timeout = Timeout.InfiniteTimeSpan;
                httpClient.DefaultRequestHeaders.ExpectContinue = false;

                var ReqURL = url.Trim();

                var patterns = new string[] { "/bypass-saml-login-flow", "/bypass-sso-login-flow" };

                foreach (var pattern in patterns)
                {
                    ReqURL = Regex.Replace(ReqURL, Regex.Escape(pattern), "", RegexOptions.IgnoreCase);
                }

                try
                {
                    // Log request
                    LogUtility.LogDebug($"Sending request: {ReqURL}");

                    // Create request object
                    var request = new HttpRequestMessage(HttpMethod.Get, $"{ReqURL}/rest/public/orbit-version");

                    // Send request and get response
                    var responseMessage = await httpClient.SendAsync(request).ConfigureAwait(false);

                    // Capture status code
                    var statusCode = (int)responseMessage.StatusCode;

                    // Read full response body
                    var responseBody = await responseMessage.Content.ReadAsStringAsync().ConfigureAwait(false);

                    // Log detailed response
                    LogUtility.LogDebug($"Status Code: {statusCode}");

                    // Parse JSON (with additional logging for unexpected structures)
                    try
                    {
                        if (string.IsNullOrWhiteSpace(responseBody))
                        {
                            LogUtility.LogError($"Empty response body from {ReqURL}");
                            return false;
                        }

                        using var doc = JsonDocument.Parse(responseBody);
                        var root = doc.RootElement;

                        var xlEdgeVersion = root
                            .EnumerateObject()
                            .FirstOrDefault(p =>
                                string.Equals(p.Name, "verionInfo",
                                              StringComparison.OrdinalIgnoreCase))
                            .Value
                            .EnumerateObject()
                            .FirstOrDefault(p =>
                                string.Equals(p.Name, "xlEdgeVersion",
                                              StringComparison.OrdinalIgnoreCase))
                            .Value
                            .GetString();

                        if (string.Equals(xlEdgeVersion,
                                          XLEdgeAppConstants.DefaultVersion,
                                          StringComparison.Ordinal))
                        {
                            return true;
                        }

                        LogUtility.LogDebug(
                            $"Version mismatch or missing: Expected='{XLEdgeAppConstants.DefaultVersion}', " +
                            $"Received='{xlEdgeVersion ?? "(null)"}'");

                        return false;
                    }
                    catch (JsonException jsonEx)
                    {
                        LogUtility.LogError(
                            $"JSON Parsing Error at {ReqURL}: {jsonEx.Message} | Raw Response: {responseBody}");
                        return false;
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogError(
                            $"Unexpected error parsing response from {ReqURL}: {ex.Message} | Raw Response: {responseBody}");
                        return false;
                    }


                }
                catch (HttpRequestException ex)
                {
                    LogUtility.LogError($"Network error for {ReqURL}: {ex.Message} | InnerException: {ex.InnerException?.Message} | StackTrace: {ex.StackTrace}");
                }
                catch (TaskCanceledException ex)
                {
                    LogUtility.LogError($"Request Timeout for {ReqURL}: {ex.Message} | StackTrace: {ex.StackTrace}");
                }
                catch (Exception ex)
                {
                    LogUtility.LogError($"General Exception for {ReqURL}: {ex.Message} | StackTrace: {ex.StackTrace}");
                }

                return false;
            }
            catch (Exception ex)
            {
                // Outer safety net around the whole compatibility check - a real failure here means
                // the URL-compatibility check couldn't run at all, worth investigating if seen
                // often, but not fatal (caller treats false as "not confirmed compatible").
                LogUtility.LogException(ex, nameof(CheckUrlCompatibility));
                return false;
            }
        }
        private void SupportLink_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                // Open support URL in default browser
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "https://www.orbitanalytics.com",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Cannot open browser: {ex.Message}");
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
    public class InstanceCompatibility : INotifyPropertyChanged
    {
        private string? _instance;
        private bool _isCompatible;

        public string? Instance
        {
            get => _instance;
            set
            {
                _instance = value;
                OnPropertyChanged(nameof(Instance));
            }
        }
        public bool IsCompatible
        {
            get => _isCompatible;
            set
            {
                _isCompatible = value;
                OnPropertyChanged(nameof(IsCompatible));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
#nullable restore
}
