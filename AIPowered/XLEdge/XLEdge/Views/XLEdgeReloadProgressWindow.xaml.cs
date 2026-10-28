using System;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace XLEdge.Views
{
    /// <summary>
    /// Busy window shown while the shell unloads the core, finds the version to load and loads it. That work blocks the UI
    /// thread, so every update forces a render pass itself. The user cannot close it (X, Alt+F4, Escape all raise Closing):
    /// only <see cref="AllowCloseAndClose"/> closes it, once the reload has finished either way.
    /// </summary>
    internal partial class XLEdgeReloadProgressWindow : Window
    {
        /// <summary>How long each stage stays on screen at least, so fast stages remain readable.</summary>
        private const int MinimumStageMilliseconds = 160;

        private const string DateFormat = "yyyy-MM-ddTHH:mm:ss";
        private const string DisplayDateFormat = "dd MMM yyyy, hh:mm tt";

        private static readonly Brush DoneStepBrush = Frozen(Color.FromRgb(0x2E, 0x86, 0xAB));
        private static readonly Brush PendingStepBrush = Frozen(Color.FromRgb(0x6B, 0x78, 0x83));

        private bool _allowClose;

        public XLEdgeReloadProgressWindow(string currentVersion, string currentReleaseDate)
        {
            InitializeComponent();
            SetReleaseDetails(TxtCurrentVersion, TxtCurrentReleaseDate, currentVersion, currentReleaseDate);
            SourceInitialized += (s, e) => WindowChromeHelper.RemoveMinimizeMaximizeButtons(this);
            Closing += OnClosing;
        }

        /// <summary>The message currently shown (for tests).</summary>
        public string CurrentMessage { get; private set; }

        /// <summary>The stage currently shown, 0 (none yet) to 4 (for tests).</summary>
        public int CurrentStep { get; private set; }

        /// <summary>
        /// Shows the window and pumps a Render and a Background pass, so its first frame is on screen before the caller starts
        /// blocking work on this thread. Set the owner (WindowInteropHelper) before calling this.
        /// </summary>
        public void ShowAndRender()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(ShowAndRender));
                return;
            }

            Show();
            Pump();
        }

        /// <summary>Shows <paramref name="message"/> as stage <paramref name="step"/> (1 closing, 2 finding, 3 loading, 4 done).</summary>
        public void ShowStage(string message, int step)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => ShowStage(message, step)));
                return;
            }

            int clamped = Math.Max(0, Math.Min(4, step));
            CurrentMessage = message;
            CurrentStep = clamped;

            TxtMessage.Text = string.IsNullOrWhiteSpace(message) ? "Reloading XLEdge Add-in..." : message;
            ReloadProgress.Value = clamped;
            SetStepState(TxtStep1, clamped >= 1);
            SetStepState(TxtStep2, clamped >= 2);
            SetStepState(TxtStep3, clamped >= 3);
            SetStepState(TxtStep4, clamped >= 4);

            if (!IsVisible) return;
            Pump();
            HoldFor(MinimumStageMilliseconds);
        }

        /// <summary>Fills "RELOADING TO" once the version to load is known.</summary>
        public void SetTarget(string version, string releaseDate)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(() => SetTarget(version, releaseDate)));
                return;
            }

            SetReleaseDetails(TxtTargetVersion, TxtTargetReleaseDate, version, releaseDate);
            if (IsVisible) Pump();
        }

        /// <summary>Lets the window close and closes it (call once the reload has finished, success or failure).</summary>
        public void AllowCloseAndClose()
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(new Action(AllowCloseAndClose));
                return;
            }

            _allowClose = true;
            Close();
        }

        private void OnClosing(object sender, CancelEventArgs e)
        {
            if (!_allowClose) e.Cancel = true;
        }

        /// <summary>Runs the pending Render and Background work now (works whether or not the dispatcher loop is running).</summary>
        private void Pump()
        {
            try
            {
                Dispatcher.Invoke(DispatcherPriority.Render, new Action(() => { }));
                Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { }));
            }
            catch (InvalidOperationException)
            {
                // Dispatcher processing is suspended (e.g. called during layout): skip the repaint, it is cosmetic.
            }
        }

        /// <summary>
        /// Keeps the current stage on screen for <paramref name="milliseconds"/> while still painting: a nested frame that a
        /// DispatcherTimer ends. PushFrame runs its own message loop, so this also works when the thread's dispatcher loop is not
        /// running. When a nested frame is not allowed, it falls back to a plain sleep.
        /// </summary>
        private void HoldFor(int milliseconds)
        {
            if (milliseconds <= 0) return;
            DispatcherTimer timer = null;
            try
            {
                var frame = new DispatcherFrame();
                timer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher) { Interval = TimeSpan.FromMilliseconds(milliseconds) };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    frame.Continue = false;
                };
                timer.Start();
                Dispatcher.PushFrame(frame);
            }
            catch (InvalidOperationException)
            {
                timer?.Stop();
                Thread.Sleep(milliseconds);
            }
        }

        private static void SetReleaseDetails(TextBlock versionText, TextBlock releaseDateText, string version, string releaseDate)
        {
            versionText.Text = string.IsNullOrWhiteSpace(version) ? "Not available" : version.Trim();
            releaseDateText.Text = FormatReleaseDate(releaseDate);
        }

        private static string FormatReleaseDate(string releaseDate)
        {
            if (string.IsNullOrWhiteSpace(releaseDate)) return "Not available";
            return DateTime.TryParseExact(releaseDate.Trim(), DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed)
                ? parsed.ToString(DisplayDateFormat, CultureInfo.InvariantCulture)
                : releaseDate;
        }

        private static void SetStepState(TextBlock step, bool reached)
        {
            step.Foreground = reached ? DoneStepBrush : PendingStepBrush;
            step.FontWeight = reached ? FontWeights.SemiBold : FontWeights.Normal;
        }

        private static Brush Frozen(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
