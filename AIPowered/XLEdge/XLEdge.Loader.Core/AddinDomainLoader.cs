using System;
using System.IO;
using System.Threading.Tasks;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Loads the XLEdge core into its own AppDomain ("XLEdgeCore_{folder}") with the release folder as ApplicationBase, and
    /// unloads it again. One instance lives in the shell for the whole Excel session and is reused for every reload.
    /// <para>
    /// No shadow copying: release folders are never rewritten (a release is extracted once into its own folder and the
    /// active one cannot be deleted), and the core's DLLs are released when its AppDomain is unloaded.
    /// </para>
    /// </summary>
    public class AddinDomainLoader
    {
        private static readonly TimeSpan UnloadTimeout = TimeSpan.FromSeconds(5);
        private AppDomain _domain;
        private string _domainName;

        /// <summary>Friendly name of the loaded core's AppDomain, or null when no core is loaded.</summary>
        public string CurrentDomainName => _domain == null ? null : _domainName;

        /// <summary>
        /// Creates the core's AppDomain, creates the core there and initializes it with <paramref name="context"/>.
        /// Refuses a release whose contract version this shell does not support, before anything is created. On any failure
        /// the half-created AppDomain is unloaded and the exception rethrown.
        /// </summary>
        public IXLEdgeAddin Load(IXLEdgeContext context, ResolvedRelease release)
        {
            if (context == null) throw new ArgumentNullException(nameof(context));
            if (release == null) throw new ArgumentNullException(nameof(release));
            if (_domain != null) throw new InvalidOperationException("An XLEdge core is already loaded; unload it first.");
            ILogger logger = context.Logger;

            string refusal = ContractCompatibility.RefusalReason(release.ContractVersion);
            if (refusal != null)
            {
                string message = $"XLEdge core release '{release.FolderName}' cannot be loaded: {refusal}.";
                logger?.LogError("AddinDomainLoader.Load: " + message);
                throw new InvalidOperationException(message);
            }
            if (!ReleaseHistoryStore.IsSafeFolderName(release.FolderName))
            {
                logger?.LogError($"AddinDomainLoader.Load: invalid release folder name '{release.FolderName}'.");
                throw new InvalidOperationException($"Invalid release folder name '{release.FolderName}'.");
            }

            string folder = Path.Combine(context.Paths.VersionsPath, release.FolderName);
            string coreDll = Path.Combine(folder, RemoteLoader.CoreAssemblyFile);
            if (!File.Exists(coreDll))
            {
                logger?.LogError($"AddinDomainLoader.Load: '{coreDll}' does not exist.");
                throw new FileNotFoundException("XLEdge core not found: " + coreDll, coreDll);
            }

            // SingleDomain: nothing is loaded domain-neutral, whatever the host process uses, so every assembly of the core
            // domain really goes away when it is unloaded.
            var setup = new AppDomainSetup { ApplicationBase = folder, ApplicationName = "XLEdgeCore", LoaderOptimization = LoaderOptimization.SingleDomain };
            string config = coreDll + ".config";
            if (File.Exists(config)) setup.ConfigurationFile = config;

            string name = "XLEdgeCore_" + release.FolderName;
            logger?.LogInfo($"AddinDomainLoader.Load: loading XLEdge core {release.Version} (released {release.ReleaseDate}) from '{folder}'.");
            _domain = AppDomain.CreateDomain(name, null, setup);
            _domainName = name;
            logger?.LogDebug($"AddinDomainLoader.Load: created AppDomain '{name}'.");
            try
            {
                // The shell's own copy of XLEdge.Loader.Core creates the core, so a rolled-back (older) release never runs an
                // older loader against this shell. Its dependencies (XLEdge.Contracts) resolve from the release folder.
                var remoteLoader = (RemoteLoader)_domain.CreateInstanceFromAndUnwrap(typeof(RemoteLoader).Assembly.Location, typeof(RemoteLoader).FullName);
                IXLEdgeAddin addin = remoteLoader.Create(folder, context.Paths.SharedDependenciesPath, logger);
                addin.Initialize(context);
                logger?.LogInfo($"AddinDomainLoader.Load: XLEdge core {release.Version} initialized.");
                return addin;
            }
            catch (Exception ex)
            {
                logger?.LogError($"AddinDomainLoader.Load: failed to load the XLEdge core from '{folder}'.", ex.ToString());
                Unload(context);
                throw;
            }
        }

        /// <summary>
        /// Unloads the core's AppDomain, waiting at most 5 seconds (AppDomain.Unload can block forever on a thread stuck in a
        /// COM call; it runs on a background thread so it can never keep Excel alive). Never throws. True when the domain is gone
        /// (or none was loaded); false when it did not unload in time or the unload failed - its files may stay locked.
        /// Never call it from inside a call that came from the core (a core-to-shell callback): that thread still has core frames
        /// on its stack and the unload would abort it.
        /// </summary>
        public bool Unload(IXLEdgeContext context)
        {
            ILogger logger = SafeLogger(context);
            AppDomain domain = _domain;
            string name = _domainName;
            _domain = null;
            _domainName = null;
            if (domain == null)
            {
                logger?.LogDebug("AddinDomainLoader.Unload: no XLEdge core loaded.");
                return true;
            }

            try
            {
                Task unload = Task.Run(() => AppDomain.Unload(domain));
                // Observe a late failure so it never surfaces as an unobserved task exception.
                unload.ContinueWith(t => { _ = t.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
                if (unload.Wait(UnloadTimeout))
                {
                    logger?.LogInfo($"AddinDomainLoader.Unload: AppDomain unloaded ('{name}').");
                    return true;
                }
                logger?.LogError($"AddinDomainLoader.Unload: AppDomain '{name}' did not unload within {UnloadTimeout.TotalSeconds:0} s - continuing; it may still be torn down in the background.");
                return false;
            }
            catch (Exception ex)
            {
                logger?.LogError($"AddinDomainLoader.Unload: error unloading AppDomain '{name}'.", ex.ToString());
                return false;
            }
        }

        private static ILogger SafeLogger(IXLEdgeContext context)
        {
            try { return context?.Logger; }
            catch (Exception) { return null; }
        }
    }
}
