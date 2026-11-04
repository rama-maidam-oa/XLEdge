using System;
using System.Threading;
using System.Windows.Interop;
using System.Windows.Threading;
using XLEdge.Shared;
using XLEdge.Views;
using WinForms = System.Windows.Forms;

namespace XLEdge.Hosting
{
    /// <summary>
    /// The add-in's <see cref="IHostUi"/>: message boxes owned by the Excel window and the WPF reload progress window. Nothing
    /// here throws back into <see cref="CoreHost"/> (and so into Add-in Express): a window that cannot be shown is skipped.
    /// </summary>
    internal sealed class WpfHostUi : IHostUi
    {
        private const string ReloadTitle = "Reload XLEdge Add-in";

        private readonly Func<IntPtr> _ownerHwnd;

        /// <param name="ownerHwnd">
        /// Reads the Excel main window each time a window is shown (Excel can recreate it); IntPtr.Zero, or a read that throws,
        /// shows the windows unowned.
        /// </param>
        public WpfHostUi(Func<IntPtr> ownerHwnd)
        {
            _ownerHwnd = ownerHwnd ?? throw new ArgumentNullException(nameof(ownerHwnd));
        }

        private IntPtr OwnerHandle()
        {
            try { return _ownerHwnd(); }
            catch { return IntPtr.Zero; }
        }

        public bool ConfirmCloseOpenWindows(string[] titles)
        {
            try
            {
                string text = "The following XLEdge window(s) are still open:\n\n" +
                              string.Join(Environment.NewLine, titles ?? new string[0]) +
                              "\n\nReloading now will close them and stop anything they are doing. Close them first for a clean reload.\n\nReload anyway?";
                return WinForms.MessageBox.Show(Owner(), text, ReloadTitle, WinForms.MessageBoxButtons.YesNo,
                           WinForms.MessageBoxIcon.Warning, WinForms.MessageBoxDefaultButton.Button2) == WinForms.DialogResult.Yes;
            }
            catch
            {
                // The question could not be asked: do not close the user's windows without an answer.
                return false;
            }
        }

        public IReloadProgress ShowProgress(string currentVersion, string currentReleaseDate)
        {
            XLEdgeReloadProgressWindow window = null;
            try
            {
                window = new XLEdgeReloadProgressWindow(VersionText.Display(currentVersion), currentReleaseDate);
                IntPtr owner = OwnerHandle();
                if (owner != IntPtr.Zero) new WindowInteropHelper(window).Owner = owner;
                window.ShowAndRender();
                return new ProgressAdapter(window);
            }
            catch
            {
                // The reload goes on without a progress window.
                try { window?.AllowCloseAndClose(); }
                catch { /* already gone */ }
                return null;
            }
        }

        public void ShowError(string message, string title)
        {
            try
            {
                WinForms.MessageBox.Show(Owner(), message, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
            }
            catch
            {
                // Nothing more can be done; CoreHost has already logged the failure.
            }
        }

        public bool WaitWhile(Func<bool> condition, TimeSpan timeout)
        {
            DateTime until = DateTime.UtcNow + timeout;
            while (true)
            {
                bool waiting;
                try { waiting = condition(); }
                catch { return true; } // nothing more to wait for
                if (!waiting) return true;
                if (DateTime.UtcNow >= until) return false;
                PumpFor(150);
            }
        }

        /// <summary>
        /// Runs this thread's message loop for <paramref name="milliseconds"/>: a nested dispatcher frame that a timer ends (it
        /// dispatches every Win32 message, so the core's dispatcher and its WebView2 run too). When a nested frame is not
        /// allowed here, WinForms' DoEvents and a short sleep instead.
        /// </summary>
        private static void PumpFor(int milliseconds)
        {
            DispatcherTimer timer = null;
            try
            {
                var frame = new DispatcherFrame();
                timer = new DispatcherTimer(DispatcherPriority.Background, Dispatcher.CurrentDispatcher) { Interval = TimeSpan.FromMilliseconds(milliseconds) };
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
                WinForms.Application.DoEvents();
                Thread.Sleep(milliseconds);
            }
        }

        private WinForms.IWin32Window Owner()
        {
            IntPtr owner = OwnerHandle();
            return owner != IntPtr.Zero ? new HwndOwner(owner) : null;
        }

        /// <summary>Lets a WinForms message box be owned by a native window (the Excel main window).</summary>
        private sealed class HwndOwner : WinForms.IWin32Window
        {
            public HwndOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; }
        }

        /// <summary>Forwards <see cref="IReloadProgress"/> to the window; the progress display is cosmetic, so failures are swallowed.</summary>
        private sealed class ProgressAdapter : IReloadProgress
        {
            private readonly XLEdgeReloadProgressWindow _window;

            public ProgressAdapter(XLEdgeReloadProgressWindow window) { _window = window; }

            public void ShowStage(string message, int step)
            {
                try { _window.ShowStage(message, step); }
                catch { /* cosmetic */ }
            }

            public void SetTarget(string version, string releaseDate)
            {
                try { _window.SetTarget(VersionText.Display(version), releaseDate); }
                catch { /* cosmetic */ }
            }

            public void Close()
            {
                try { _window.AllowCloseAndClose(); }
                catch { /* already gone */ }
            }
        }
    }
}
