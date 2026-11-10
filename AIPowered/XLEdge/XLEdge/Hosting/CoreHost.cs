using System;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Hosting
{
    /// <summary>
    /// Starts, reloads and shuts down the XLEdge core for the shell. Nothing here throws back into Add-in Express: failures are
    /// logged and shown with <see cref="IHostUi.ShowError"/> (after the progress window has closed), and the add-in stays usable
    /// for another Reload.
    /// <para>
    /// While a reload runs, <see cref="Addin"/> is null, so ribbon clicks, Excel events and GLSense calls that Excel processes
    /// meanwhile (the progress window pumps messages) never reach the core being replaced. Excel closing in the middle of a
    /// reload (<see cref="Shutdown"/>) stops it from loading anything.
    /// </para>
    /// <para>
    /// A reload first logs XLEdge out - and GLSense with it - when the core is logged in (<c>LogOutBeforeReload</c>), so both
    /// add-ins end logged out and the next core starts logged out.
    /// </para>
    /// </summary>
    internal sealed class CoreHost
    {
        private const string StartTitle = "XLEdge";
        private const string ReloadTitle = "Reload XLEdge Add-in";
        internal const string StuckUnloadMessage = "The previous XLEdge version did not close within 5 seconds. XLEdge keeps working; restart Excel before installing or switching versions again.";

        private readonly IXLEdgeContext _context;
        private readonly ICoreLoader _loader;
        private readonly IReleaseCatalog _catalog;
        private readonly IHostUi _ui;
        private readonly IPaneHost _panes;
        private bool _busy;
        private bool _shutdownRequested;

        public CoreHost(IXLEdgeContext context, ICoreLoader loader, IReleaseCatalog catalog, IHostUi ui, IPaneHost panes)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _loader = loader ?? throw new ArgumentNullException(nameof(loader));
            _catalog = catalog ?? throw new ArgumentNullException(nameof(catalog));
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));
            _panes = panes ?? throw new ArgumentNullException(nameof(panes));
        }

        /// <summary>The loaded core (a remoting proxy), or null when none is loaded or a reload is running.</summary>
        public IXLEdgeAddin Addin { get; private set; }

        /// <summary>Loads the release the bootstrapper chooses. Returns false (after telling the user) when nothing could be loaded.</summary>
        public bool Start()
        {
            if (_busy || _shutdownRequested) return false;
            _busy = true;
            // Debug lines of the start are written as one block after the "Start of XLEdge version" marker.
            using (Scope("XLEdge start"))
            try
            {
                ResolvedRelease release = null;
                try { release = _catalog.ResolveVersionToLoad(_context, "Install"); }
                catch (Exception ex) { Log(l => l.LogError("CoreHost.Start: choosing the core release failed.", ex.ToString())); }

                if (release == null)
                {
                    Log(l => l.LogError("CoreHost.Start: no usable XLEdge core release."));
                    _ui.ShowError("XLEdge could not find a usable version of the add-in. Use Reload Add-in to load one, or reinstall XLEdge. Check the XLEdge log for details.", StartTitle);
                    return false;
                }

                Exception failure = TryLoad(release);
                if (failure != null)
                {
                    _ui.ShowError($"XLEdge could not load version {VersionText.Display(release.Version)}: {failure.Message}", StartTitle);
                    return false;
                }
                return true;
            }
            finally
            {
                _busy = false;
            }
        }

        /// <summary>
        /// Replaces the loaded core with the release <paramref name="resolve"/> returns (it runs after the old core is unloaded,
        /// so it may stage or extract a release). False when refused (busy core, another reload), cancelled or failed.
        /// </summary>
        public bool Reload(Func<ResolvedRelease> resolve)
        {
            if (resolve == null) throw new ArgumentNullException(nameof(resolve));
            if (_busy || _shutdownRequested)
            {
                Log(l => l.LogWarn("CoreHost.Reload: a start, reload or shutdown is already running; ignored."));
                return false;
            }
            _busy = true;
            // Debug lines of the reload are written as one block after the new "Start of XLEdge version" marker.
            using (Scope("XLEdge reload"))
            try
            {
                if (CoreIsBusy())
                {
                    Log(l => l.LogWarn("CoreHost.Reload: refused - the core is busy."));
                    _ui.ShowError("XLEdge is busy (a report or refresh is running, or an XLEdge window is open). Use Reload Add-in again when it has finished.", ReloadTitle);
                    return false;
                }

                string[] titles = OpenWindowTitles();
                if (titles.Length > 0 && !_ui.ConfirmCloseOpenWindows(titles))
                {
                    Log(l => l.LogDebug("CoreHost.Reload: cancelled by the user because core windows are open."));
                    return false;
                }

                // Before the progress window (topmost; it would hide the core's "Logging Off" window).
                LogOutBeforeReload();

                IXLEdgeAddin old = Addin;
                Addin = null;

                IReloadProgress progress = null;
                try { progress = _ui.ShowProgress(_context.Version, _context.ReleaseDate); }
                catch (Exception ex) { Log(l => l.LogWarn("CoreHost.Reload: the progress window could not be shown: " + ex.Message)); }

                string error = null;
                bool reloaded = false, stuckUnload = false;
                try
                {
                    Stage(progress, 1, "Closing the current XLEdge session...");
                    if (old != null) Log(l => l.EndReleaseSession(_context.Version, _context.ReleaseDate));
                    bool unloaded = UnloadCore(old, "CoreHost.Reload");
                    if (!unloaded) stuckUnload = true;

                    Stage(progress, 2, "Finding the selected XLEdge version...");
                    ResolvedRelease release = null;
                    Exception resolveFailure = null;
                    try { release = resolve(); }
                    catch (Exception ex)
                    {
                        resolveFailure = ex;
                        Log(l => l.LogError("CoreHost.Reload: choosing the core release failed.", ex.ToString()));
                    }

                    // Excel closing meanwhile is handled before loading and once more after this block (no message then).
                    if (resolveFailure != null)
                    {
                        error = $"Reload failed: {resolveFailure.Message}. Use Reload Add-in again, or restart Excel.";
                    }
                    else if (release == null)
                    {
                        error = "Reload failed - no usable XLEdge version was found. The previous version has been closed; use Reload Add-in again.";
                    }
                    else
                    {
                        try { progress?.SetTarget(release.Version, release.ReleaseDate); } catch (Exception) { }
                        Stage(progress, 3, "Loading the selected XLEdge version...");
                        Exception failure = _shutdownRequested ? null : TryLoad(release);
                        if (_shutdownRequested)
                        {
                            Log(l => l.LogDebug("CoreHost.Reload: Excel is closing - not loading a new core."));
                        }
                        else if (failure != null)
                        {
                            error = $"Reload failed: {failure.Message}. Use Reload Add-in again, or restart Excel.";
                        }
                        else
                        {
                            Stage(progress, 4, "XLEdge reload complete.");
                            reloaded = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log(l => l.LogError("CoreHost.Reload: unexpected failure.", ex.ToString()));
                    error = $"Reload failed: {ex.Message}. Use Reload Add-in again, or restart Excel.";
                }
                finally
                {
                    // Closed before any message: the progress window is topmost and would hide it.
                    try { progress?.Close(); } catch (Exception) { }
                }

                if (_shutdownRequested)
                {
                    // Excel closed while the new core was being loaded: take it down again.
                    if (Addin != null) UnloadCore(Addin, "CoreHost.Reload");
                    return false;
                }
                if (error != null)
                {
                    Log(l => l.LogError("CoreHost.Reload: " + error));
                    _ui.ShowError(error, ReloadTitle);
                }
                else if (reloaded && stuckUnload)
                {
                    // The new core runs; the old one may still hold its files (a later update would fail to replace them).
                    Log(l => l.LogWarn("CoreHost.Reload: the previous core did not unload in time; the user is told."));
                    _ui.ShowError(StuckUnloadMessage, ReloadTitle);
                }
                return reloaded;
            }
            finally
            {
                _busy = false;
            }
        }

        /// <summary>
        /// Shuts the core down and unloads it (Excel is closing). When a reload is running, it finishes the unload and loads
        /// nothing. Never throws.
        /// </summary>
        public void Shutdown()
        {
            _shutdownRequested = true;
            if (_busy) return;
            UnloadCore(Addin, "CoreHost.Shutdown");
        }

        // ---- helpers -------------------------------------------------------------------

        /// <summary>The first contract whose cores read and keep session values (IXLEdgeContext.GetSessionValue).</summary>
        private const int FirstContractWithSessionValues = 3;

        /// <summary>The first contract whose cores keep the login across a reload (session value "XLEdge.Login", Stage 2e).</summary>
        private const int FirstContractWithSessionLogin = 4;

        private Exception TryLoad(ResolvedRelease release)
        {
            // An older core (Release History) neither reads nor keeps session values: by the time a newer core loads again
            // they would be stale (e.g. session-only options coming back over options saved in the meantime).
            if (release.ContractVersion < FirstContractWithSessionValues && _context is ISessionBag bag)
            {
                try
                {
                    bag.ClearSession();
                    Log(l => l.LogDebug($"CoreHost: core contract {release.ContractVersion} keeps no session values - cleared them."));
                }
                catch (Exception ex) { Log(l => l.LogWarn("CoreHost: could not clear the session values: " + ex.Message)); }
            }
            // A core that predates the session login shows a logged-out ribbon and drops the login at its own Shutdown: the one it
            // leaves here would log the user in again silently when a newer core loads.
            else if (release.ContractVersion < FirstContractWithSessionLogin)
            {
                try
                {
                    _context?.SetSessionValue("XLEdge.Login", null);
                    Log(l => l.LogDebug($"CoreHost: core contract {release.ContractVersion} keeps no login - dropped it."));
                }
                catch (Exception ex) { Log(l => l.LogWarn("CoreHost: could not drop the session login: " + ex.Message)); }
            }
            _context.Version = release.Version;
            _context.ReleaseDate = release.ReleaseDate;
            _context.ActiveFolderName = release.FolderName;
            try
            {
                Log(l => l.BeginReleaseSession(release.Version, release.ReleaseDate));
                Addin = _loader.Load(_context, release);
            }
            catch (Exception ex)
            {
                Addin = null;
                ClearRelease();
                Log(l => l.LogError($"CoreHost: loading core release '{release.FolderName}' failed.", ex.ToString()));
                return ex;
            }
            try { _catalog.RememberLoaded(_context, release.FolderName); }
            catch (Exception ex) { Log(l => l.LogWarn("CoreHost: could not remember the loaded release: " + ex.Message)); }
            try { _panes.AttachAll(); }
            catch (Exception ex) { Log(l => l.LogWarn("CoreHost: could not attach the task panes: " + ex.Message)); }
            return null;
        }

        /// <summary>Shuts the core down and unloads it; false when the unload did not finish (or failed).</summary>
        private bool UnloadCore(IXLEdgeAddin old, string caller)
        {
            Addin = null;
            if (old != null)
            {
                try { _panes.DetachAll(old); }
                catch (Exception ex) { Log(l => l.LogError(caller + ": detaching the task panes failed.", ex.ToString())); }
                try { old.Shutdown(); }
                catch (Exception ex) { Log(l => l.LogError(caller + ": the core's Shutdown failed.", ex.ToString())); }
            }
            bool unloaded;
            try { unloaded = _loader.Unload(_context); }
            catch (Exception ex)
            {
                Log(l => l.LogError(caller + ": unloading the core failed.", ex.ToString()));
                unloaded = false;
            }
            ClearRelease();
            return unloaded;
        }

        private void ClearRelease()
        {
            _context.Version = null;
            _context.ReleaseDate = null;
            _context.ActiveFolderName = null;
        }

        /// <summary>How long a reload waits for the core's logout (its panes' applogout navigations have their own limits).</summary>
        internal static TimeSpan LogoutTimeout = TimeSpan.FromSeconds(60);

        /// <summary>
        /// A hot reload logs XLEdge out, and GLSense with it: the outgoing core runs its Logout button's sequence (each pane's
        /// applogout, GLSense's LogoutSession, the login cleared, the logged-out ribbon), so it leaves no login for the next
        /// core and both add-ins end logged out. The shell keeps Excel's messages flowing until the core is logged out and no
        /// longer busy; a logout that does not end in time is logged and the reload goes on. Nothing happens when not logged in.
        /// </summary>
        private void LogOutBeforeReload()
        {
            if (!CoreIsLoggedIn()) return;
            try
            {
                Log(l => l.LogInfo("CoreHost.Reload: logging out of XLEdge (and GLSense) before the reload."));
                Addin.OnRibbonAction(RibbonControlIds.RibEdgeLogout, null);
                bool loggedOut = _ui.WaitWhile(() => !_shutdownRequested && CoreIsLoggedIn(), LogoutTimeout)
                                 && _ui.WaitWhile(() => !_shutdownRequested && CoreIsBusy(), LogoutTimeout);
                if (loggedOut) Log(l => l.LogInfo("CoreHost.Reload: logged out; reloading."));
                else Log(l => l.LogWarn($"CoreHost.Reload: the logout did not end within {LogoutTimeout.TotalSeconds:0} s; reloading anyway."));
            }
            catch (Exception ex)
            {
                Log(l => l.LogWarn("CoreHost.Reload: the logout before the reload failed; reloading anyway: " + ex.Message));
            }
        }

        private bool CoreIsLoggedIn()
        {
            try { return Addin?.GetLoginInfo()?.IsLoggedIn ?? false; }
            catch (Exception ex)
            {
                Log(l => l.LogDebug("CoreHost.Reload: could not ask the core whether it is logged in: " + ex.Message));
                return false;
            }
        }

        private bool CoreIsBusy()
        {
            try { return Addin?.IsBusy() ?? false; }
            catch (Exception ex)
            {
                Log(l => l.LogDebug("CoreHost.Reload: could not ask the core whether it is busy: " + ex.Message));
                return false;
            }
        }

        private string[] OpenWindowTitles()
        {
            try { return Addin?.GetOpenWindowTitles() ?? new string[0]; }
            catch (Exception ex)
            {
                // An older core (Release History) may not have this member: treat it as "no windows open".
                Log(l => l.LogDebug("CoreHost.Reload: could not list the core's open windows: " + ex.Message));
                return new string[0];
            }
        }

        private void Stage(IReloadProgress progress, int step, string message)
        {
            Log(l => l.LogDebug($"CoreHost.Reload: {message}"));
            try { progress?.ShowStage(message, step); } catch (Exception) { }
        }

        /// <summary>A log scope (debug lines buffered until it ends), or null when there is no logger.</summary>
        private IDisposable Scope(string name)
        {
            try { return _context.Logger?.BeginLogScope(name); }
            catch (Exception) { return null; }
        }

        private void Log(Action<ILogger> write)
        {
            try
            {
                ILogger logger = _context.Logger;
                if (logger != null) write(logger);
            }
            catch (Exception)
            {
                // Logging must never break starting or reloading the core.
            }
        }
    }
}
