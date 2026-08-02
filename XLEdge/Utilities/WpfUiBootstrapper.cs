using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using Wpf.Ui.Appearance;

namespace XLEdge.Utilities
{
    /// <summary>
    /// Initializes WPF-UI theming and resource dictionaries for the application. Loads the real
    /// Wpf.Ui theme/control resource dictionaries via pack URIs, and falls back to hand-rolled
    /// resources for any key that fails to load, so the UI never ends up with an undefined
    /// DynamicResource lookup.
    /// </summary>
    public static class WpfUiBootstrapper
    {
        private static bool _initialized;
        private static bool _warmedUp;
        private static readonly object _lock = new object();
        private static string _currentBaseTheme;

        public static bool IsInitialized => _initialized;

        /// <summary>
        /// Initializes WPF-UI once during ribbon load. accentHex is accepted for call-site
        /// compatibility, but WPF-UI's own theming is Light/Dark-only, so only baseTheme actually
        /// drives behavior here.
        /// </summary>
        public static void Init(string accentHex, string baseTheme)
        {
            if (_initialized && string.Equals(_currentBaseTheme, baseTheme, StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            lock (_lock)
            {
                if (_initialized)
                {
                    ApplyTheme(baseTheme);
                    return;
                }

                try
                {
                    WpfAppManager.EnsureApplication();

                    var app = Application.Current;
                    if (app == null)
                    {
                        return;
                    }

                    void DoInit()
                    {
                        using (DpiAwarenessHelper.SetPerMonitorAware())
                        {
                            LoadAllResourcesManually(app);
                            ApplyTheme(baseTheme);
                        }

                        _currentBaseTheme = baseTheme;
                        _initialized = true;
                    }

                    if (app.Dispatcher.CheckAccess())
                    {
                        DoInit();
                    }
                    else
                    {
                        app.Dispatcher.Invoke(DoInit);
                    }
                }
                catch (Exception ex)
                {
                    LogUtility.LogError($"WpfUiBootstrapper.Init: initialization failed - {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Forces every merged resource dictionary's Keys to be enumerated once, which triggers any
        /// lazily-loaded pack:// ResourceDictionary's Source to resolve immediately.
        /// </summary>
        public static void PreloadResources()
        {
            if (!_initialized)
            {
                LogUtility.LogWarn("WpfUiBootstrapper.PreloadResources: called before Init - ignoring.");
                return;
            }

            try
            {
                var app = Application.Current;
                if (app == null)
                {
                    return;
                }

                WpfAppManager.InvokeOnWpfThread(() =>
                {
                    try
                    {
                        foreach (var dict in app.Resources.MergedDictionaries)
                        {
                            var keys = dict.Keys; // force loading
                        }
                    }
                    catch (Exception ex)
                    {
                        LogUtility.LogError($"WpfUiBootstrapper.PreloadResources: error preloading resources - {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"WpfUiBootstrapper.PreloadResources: failed - {ex.Message}");
            }
        }

        /// <summary>
        /// Creates, shows, and immediately closes one throwaway, fully invisible window so the WPF
        /// window-construction/first-render code path - HwndSource creation, DpiAwareWindow's DPI/
        /// layout plumbing (including its GetDpiForWindow/SetWindowPos P/Invoke stubs), and the first
        /// JIT pass over real ControlTemplates like ModernToggleSwitch's track+thumb - gets paid for
        /// once during ribbon load, instead of on whichever real window the user happens to open
        /// first. Deliberately synchronous, same as Init/PreloadResources above: the whole point is to
        /// move this cost into ribbon load's existing startup latency rather than the user's first
        /// click, so it's fine if this takes a moment.
        /// </summary>
        public static void WarmUpFirstWindow()
        {
            if (_warmedUp || !_initialized)
            {
                return;
            }

            try
            {
                var app = Application.Current;
                if (app == null)
                {
                    return;
                }

                if (app.Dispatcher.CheckAccess())
                {
                    RunWarmUpWindow();
                }
                else
                {
                    app.Dispatcher.Invoke(RunWarmUpWindow);
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"WpfUiBootstrapper.WarmUpFirstWindow: failed - {ex.Message}");
            }
            finally
            {
                // Only ever attempted once - if it failed, ribbon load already paid whatever cost it
                // was going to pay, and the user's first real window falls back to warming up on its
                // own the way it always has, so there's nothing to gain by retrying.
                _warmedUp = true;
            }
        }

        private static void RunWarmUpWindow()
        {
            DpiAwareWindow warmupWindow = null;

            try
            {
                warmupWindow = new DpiAwareWindow
                {
                    Title = "warmup",
                    WindowStyle = WindowStyle.None,
                    AllowsTransparency = true,
                    Background = Brushes.Transparent,
                    ShowInTaskbar = false,
                    ShowActivated = false,
                    ResizeMode = ResizeMode.NoResize,
                    SizeToContent = SizeToContent.Manual,
                    Width = 50,
                    Height = 50,
                    MinWidth = 10,
                    MinHeight = 10,
                    Left = -32000,
                    Top = -32000,
                    Opacity = 0
                };

                // A representative slice of the controls/styles real windows use, so their layout/
                // render and ControlTemplate code paths get JIT'd here rather than on the first real
                // window the user opens.
                var stack = new StackPanel();
                stack.Children.Add(new TextBlock { Text = "warmup" });

                if (TryGetStyle("ModernToggleSwitch", out Style toggleStyle))
                {
                    stack.Children.Add(new CheckBox { Style = toggleStyle, Content = "warmup", IsChecked = true });
                }

                if (TryGetStyle("DynamicContentButton", out Style buttonStyle))
                {
                    stack.Children.Add(new Button { Style = buttonStyle, Content = "warmup" });
                }

                warmupWindow.Content = new Border { Padding = new Thickness(4), Child = stack };

                warmupWindow.Loaded += (s, e) =>
                {
                    warmupWindow.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
                    {
                        try
                        {
                            warmupWindow.Close();
                        }
                        catch (Exception ex)
                        {
                            LogUtility.LogDebug($"WpfUiBootstrapper.RunWarmUpWindow: close failed - {ex.Message}");
                        }
                    }));
                };

                warmupWindow.Show();
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"WpfUiBootstrapper.RunWarmUpWindow: failed - {ex.Message}");
                try
                {
                    warmupWindow?.Close();
                }
                catch (Exception closeEx)
                {
                    LogUtility.LogDebug($"WpfUiBootstrapper.RunWarmUpWindow: cleanup close failed - {closeEx.Message}");
                }
            }
        }

        private static bool TryGetStyle(string key, out Style style)
        {
            style = Application.Current?.TryFindResource(key) as Style;
            return style != null;
        }

        private static void ApplyTheme(string baseTheme)
        {
            try
            {
                var theme = string.Equals(baseTheme, "Dark", StringComparison.OrdinalIgnoreCase)
                    ? ApplicationTheme.Dark
                    : ApplicationTheme.Light;

                ApplicationThemeManager.Apply(theme);
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"WpfUiBootstrapper.ApplyTheme: failed to apply '{baseTheme}' theme - {ex.Message}");
            }
        }

        private static void LoadAllResourcesManually(Application app)
        {
            // Load the real Wpf.Ui theme/control dictionaries first, then add hand-rolled fallbacks
            // for anything still missing so a resource-loading failure never crashes the UI.
            LoadWpfUiFromPackUris(app);

            // XLEdge's own global styles (SuggestAppendComboBox and other custom controls).
            LoadResourceIfMissing(app, "pack://application:,,,/XLEdge;component/Themes/GlobalStyles.xaml");

            AddFallbackResources(app);
        }

        private static void LoadWpfUiFromPackUris(Application app)
        {
            var mergedDictionaries = app.Resources.MergedDictionaries;

            try
            {
                var themeUri = new Uri("pack://application:,,,/Wpf.Ui;component/Resources/Theme/Light.xaml", UriKind.Absolute);
                if (!mergedDictionaries.Any(d => d.Source == themeUri))
                {
                    mergedDictionaries.Add(new ResourceDictionary { Source = themeUri });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfUiBootstrapper.LoadWpfUiFromPackUris: failed to load Wpf.Ui theme");
            }

            try
            {
                var controlsUri = new Uri("pack://application:,,,/Wpf.Ui;component/Resources/Wpf.Ui.xaml", UriKind.Absolute);
                if (!mergedDictionaries.Any(d => d.Source == controlsUri))
                {
                    mergedDictionaries.Add(new ResourceDictionary { Source = controlsUri });
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, "WpfUiBootstrapper.LoadWpfUiFromPackUris: failed to load Wpf.Ui controls");
            }
        }

        private static void LoadResourceIfMissing(Application app, string uriString)
        {
            try
            {
                var uri = new Uri(uriString, UriKind.Absolute);

                if (app.Resources.MergedDictionaries.Any(d => d.Source == uri))
                {
                    return;
                }

                var dict = new ResourceDictionary { Source = uri };
                var keys = dict.Keys; // force loading

                app.Resources.MergedDictionaries.Add(dict);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, $"WpfUiBootstrapper.LoadResourceIfMissing: resource not found at {uriString}");
            }
        }

        /// <summary>
        /// Defines fallback values for the Wpf.Ui/Fluent design-token resource keys used by a
        /// FluentWindow's default chrome (background/border/text brushes), applied only for keys the
        /// real Wpf.Ui pack URI resources above failed to load.
        /// </summary>
        private static void AddFallbackResources(Application app)
        {
            var resources = app.Resources;

            var accentColor = (Color)ColorConverter.ConvertFromString(XLEdgeAppConstants.GLAccentHex);
            var accentBrush = CreateFrozenBrush(accentColor);

            AddResourceIfMissing(resources, "SystemAccentColor", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorPrimaryBrush", accentBrush);
            AddResourceIfMissing(resources, "SystemAccentColorSecondary", accentColor);
            AddResourceIfMissing(resources, "SystemAccentColorTertiary", accentColor);

            AddResourceIfMissing(resources, "ControlBackgroundBrush", CreateFrozenBrush(Colors.White));
            AddResourceIfMissing(resources, "ControlSubtleBackgroundBrush", CreateFrozenBrush(Color.FromArgb(255, 240, 240, 240)));
            AddResourceIfMissing(resources, "ControlSubtleSecondaryBrush", CreateFrozenBrush(Color.FromArgb(255, 225, 225, 225)));
            AddResourceIfMissing(resources, "ControlSolidAccentBrush", accentBrush);
            AddResourceIfMissing(resources, "ControlTextBrush", CreateFrozenBrush(Colors.Black));
            AddResourceIfMissing(resources, "ControlBorderBrush", CreateFrozenBrush(Color.FromArgb(255, 200, 200, 200)));

            AddResourceIfMissing(resources, "CardBackgroundFillColorDefaultBrush", CreateFrozenBrush(Colors.White));
            AddResourceIfMissing(resources, "CardStrokeColorDefaultBrush", CreateFrozenBrush(Color.FromArgb(255, 220, 220, 220)));

            AddResourceIfMissing(resources, "TextFillColorPrimaryBrush", CreateFrozenBrush(Colors.Black));
            AddResourceIfMissing(resources, "TextFillColorSecondaryBrush", CreateFrozenBrush(Color.FromArgb(255, 100, 100, 100)));

            AddResourceIfMissing(resources, "ApplicationBackgroundBrush", CreateFrozenBrush(Color.FromArgb(255, 240, 240, 240)));
        }

        private static void AddResourceIfMissing(ResourceDictionary resources, string key, object value)
        {
            if (!resources.Contains(key))
            {
                resources[key] = value;
            }
        }

        private static SolidColorBrush CreateFrozenBrush(Color color)
        {
            var brush = new SolidColorBrush(color);
            if (brush.CanFreeze)
            {
                brush.Freeze();
            }

            return brush;
        }
    }
}
