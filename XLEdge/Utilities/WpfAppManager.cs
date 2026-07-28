using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace XLEdge.Utilities
{
    public static class WpfAppManager
    {
        private static readonly object _lock = new object();
        private static bool _dispatcherInitialized = false;

        public static void EnsureApplication()
        {
            if (Application.Current != null)
                return;

            lock (_lock)
            {
                if (Application.Current == null)
                {
                    Dispatcher _wpfDispatcher;

                    // Set DPI awareness before creating Application
                    using (DpiAwarenessHelper.SetPerMonitorAware())
                    {
                        var app = new Application
                        {
                            ShutdownMode = ShutdownMode.OnExplicitShutdown
                        };

                        // Store the dispatcher for later use
                        _wpfDispatcher = app.Dispatcher;

                        // Initialize dispatcher properly
                        if (!_dispatcherInitialized)
                        {
                            // Create a dummy control on the UI thread to initialize dispatcher
                            _wpfDispatcher.Invoke(() =>
                            {
                                var dummy = new System.Windows.Controls.Control();
                            });
                            _dispatcherInitialized = true;
                        }

                        app.DispatcherUnhandledException += OnDispatcherUnhandledException;
                    }
                }
            }
        }

        private static void OnDispatcherUnhandledException(object sender,
            DispatcherUnhandledExceptionEventArgs e)
        {
            LogUtility.LogError($"WPF Dispatcher Error: {e.Exception.Message}");
            if (e.Exception is System.IO.IOException)
            {
                LogUtility.LogError($"IO Exception in WPF: {e.Exception.StackTrace}");
                // Don't mark as handled for IO exceptions so we can see them
            }
            else
            {
                e.Handled = true; // Prevent application crash for non-IO exceptions
            }
        }

        public static void InvokeOnWpfThread(Action action)
        {
            if (Application.Current != null)
            {
                var dispatcher = Application.Current.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    SafeExecute(action);
                }
                else
                {
                    dispatcher.Invoke(() => SafeExecute(action));
                }
            }
        }
        public static T InvokeOnWpfThread<T>(Func<T> func)
        {
            if (Application.Current != null)
            {
                var dispatcher = Application.Current.Dispatcher;
                if (dispatcher.CheckAccess())
                {
                    return SafeExecute(func);
                }
                else
                {
                    return dispatcher.Invoke(() => SafeExecute(func));
                }
            }

            // Application.Current is null, so the action cannot run; log a warning and return the default value.
            LogUtility.LogWarn($"{nameof(InvokeOnWpfThread)}: Application.Current is null - action was not run, returning default({typeof(T).Name}).");
            return default(T);
        }

        private static T SafeExecute<T>(Func<T> func)
        {
            try
            {
                return func();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception in WPF operation: {ex.Message}");
                // Retry once after a small delay
                System.Threading.Thread.Sleep(100);
                try
                {
                    return func();
                }
                catch (Exception retryEx)
                {
                    LogUtility.LogError($"Retry failed: {retryEx.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in WPF operation: {ex.Message}");
                throw;
            }
        }
        private static void SafeExecute(Action action)
        {
            try
            {
                action();
            }
            catch (System.IO.IOException ex)
            {
                LogUtility.LogError($"IO Exception in WPF operation: {ex.Message}");
                // Retry once after a small delay
                System.Threading.Thread.Sleep(100);
                try
                {
                    action();
                }
                catch (Exception retryEx)
                {
                    LogUtility.LogError($"Retry failed: {retryEx.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                LogUtility.LogError($"Error in WPF operation: {ex.Message}");
                throw;
            }
        }
    }
}
