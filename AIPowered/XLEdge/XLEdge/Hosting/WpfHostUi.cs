using System;
using System.Windows.Interop;
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
