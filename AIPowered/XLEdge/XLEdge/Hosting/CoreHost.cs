using System;
using XLEdge.Contracts;
using XLEdge.Loader.Core;

namespace XLEdge.Hosting
{
    /// <summary>
    /// Starts, reloads and shuts down the XLEdge core for the shell. Nothing here throws back into Add-in Express: failures are
    /// logged and shown with <see cref="IHostUi.ShowError"/> (after the progress window has closed), and the add-in stays usable
    /// for another Reload.
    /// </summary>
    internal sealed class CoreHost
    {
        private const string StartTitle = "XLEdge";
        private const string ReloadTitle = "Reload XLEdge Add-in";

        private readonly IXLEdgeContext _context;
        private readonly ICoreLoader _loader;
        private readonly IReleaseCatalog _catalog;
        private readonly IHostUi _ui;
        private bool _busy;

        public CoreHost(IXLEdgeContext context, ICoreLoader loader, IReleaseCatalog catalog, IHostUi ui)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _loader = loader ?? throw new ArgumentNullException(nameof(loader));
            _catalog = catalog ?? throw new ArgumentNullException(nameof(catalog));
            _ui = ui ?? throw new ArgumentNullException(nameof(ui));
        }

        /// <summary>The loaded core (a remoting proxy), or null when none is loaded.</summary>
        public IXLEdgeAddin Addin { get; private set; }

        /// <summary>Loads the release the bootstrapper chooses. Returns false (after telling the user) when nothing could be loaded.</summary>
        public bool Start()
        {
            if (_busy) return false;
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
        /// so it may extract a staged release). False when cancelled, refused (another reload running) or failed.
        /// </summary>
        public bool Reload(Func<ResolvedRelease> resolve)
        {
            if (resolve == null) throw new ArgumentNullException(nameof(resolve));
            if (_busy)
            {
                Log(l => l.LogWarn("CoreHost.Reload: a start or reload is already running; ignored."));
                return false;
            }
            _busy = true;
            try
            {
                string[] titles = OpenWindowTitles();
                if (titles.Length > 0 && !_ui.ConfirmCloseOpenWindows(titles))
                {
                    Log(l => l.LogDebug("CoreHost.Reload: cancelled by the user because core windows are open."));
                    return false;
                }

                IReloadProgress progress = null;
                try { progress = _ui.ShowProgress(_context.Version, _context.ReleaseDate); }
                catch (Exception ex) { Log(l => l.LogWarn("CoreHost.Reload: the progress window could not be shown: " + ex.Message)); }

                string error = null;
                bool reloaded = false;
                try
                {
                    Stage(progress, 1, "Closing the current XLEdge session...");
                    if (Addin != null) Log(l => l.EndReleaseSession(_context.Version, _context.ReleaseDate));
                    UnloadCore("CoreHost.Reload");

                    Stage(progress, 2, "Finding the selected XLEdge version...");
                    ResolvedRelease release = null;
                    try { release = resolve(); }
                    catch (Exception ex) { Log(l => l.LogError("CoreHost.Reload: choosing the core release failed.", ex.ToString())); }

                    if (release == null)
                    {
                        error = "Reload failed - no usable XLEdge core version was found. The previous version has been closed; use Reload Add-in again.";
                    }
                    else
                    {
                        try { progress?.SetTarget(release.Version, release.ReleaseDate); } catch (Exception) { }
                        Stage(progress, 3, "Loading the selected XLEdge version...");
                        Exception failure = TryLoad(release);
                        if (failure != null)
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

        /// <summary>Shuts the core down and unloads it (Excel is closing). Never throws.</summary>
        public void Shutdown() => UnloadCore("CoreHost.Shutdown");

        // ---- helpers -------------------------------------------------------------------

        private Exception TryLoad(ResolvedRelease release)
        {
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
            return null;
        }

        private void UnloadCore(string caller)
        {
            IXLEdgeAddin old = Addin;
            Addin = null;
            if (old != null)
            {
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
