using System;
using XLEdge.Contracts;
using XLEdge.Loader.Core;

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
    /// </summary>
    internal sealed class CoreHost
    {
        private const string StartTitle = "XLEdge";
        private const string ReloadTitle = "Reload XLEdge Add-in";

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
            try
            {
                ResolvedRelease release = null;
                try { release = _catalog.ResolveVersionToLoad(_context, "Install"); }
                catch (Exception ex) { Log(l => l.LogError("CoreHost.Start: choosing the core release failed.", ex.ToString())); }

                if (release == null)
                {
                    Log(l => l.LogError("CoreHost.Start: no usable XLEdge core release."));
                    _ui.ShowError("XLEdge could not find a usable core version. Use Reload Add-in to load one, or reinstall XLEdge. Check the XLEdge log for details.", StartTitle);
                    return false;
                }

                Exception failure = TryLoad(release);
                if (failure != null)
                {
                    _ui.ShowError($"XLEdge could not load core version {release.Version}: {failure.Message}", StartTitle);
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

                IXLEdgeAddin old = Addin;
                Addin = null;

                IReloadProgress progress = null;
                try { progress = _ui.ShowProgress(_context.Version, _context.ReleaseDate); }
                catch (Exception ex) { Log(l => l.LogWarn("CoreHost.Reload: the progress window could not be shown: " + ex.Message)); }

                string error = null;
                bool reloaded = false;
                try
                {
                    Stage(progress, 1, "Closing the current XLEdge session...");
                    if (old != null) Log(l => l.EndReleaseSession(_context.Version, _context.ReleaseDate));
                    UnloadCore(old, "CoreHost.Reload");

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
                        error = "Reload failed - no usable XLEdge core version was found. The previous version has been closed; use Reload Add-in again.";
                    }
                    else
                    {
                        try { progress?.SetTarget(release.Version, release.ReleaseDate); } catch (Exception) { }
                        Stage(progress, 3, "Loading the selected XLEdge version...");
                        Exception failure = _shutdownRequested ? null : TryLoad(release);
                        if (_shutdownRequested)
                        {
                            Log(l => l.LogInfo("CoreHost.Reload: Excel is closing - not loading a new core."));
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

        private void UnloadCore(IXLEdgeAddin old, string caller)
        {
            Addin = null;
            if (old != null)
            {
                try { _panes.DetachAll(old); }
                catch (Exception ex) { Log(l => l.LogError(caller + ": detaching the task panes failed.", ex.ToString())); }
                try { old.Shutdown(); }
                catch (Exception ex) { Log(l => l.LogError(caller + ": the core's Shutdown failed.", ex.ToString())); }
            }
            try { _loader.Unload(_context); }
            catch (Exception ex) { Log(l => l.LogError(caller + ": unloading the core failed.", ex.ToString())); }
            ClearRelease();
        }

        private void ClearRelease()
        {
            _context.Version = null;
            _context.ReleaseDate = null;
            _context.ActiveFolderName = null;
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
