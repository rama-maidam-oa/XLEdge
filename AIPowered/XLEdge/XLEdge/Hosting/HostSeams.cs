using XLEdge.Contracts;
using XLEdge.Loader.Core;

namespace XLEdge.Hosting
{
    /// <summary>The windows CoreHost needs while it starts or reloads the core (WPF in the add-in, fakes in tests).</summary>
    internal interface IHostUi
    {
        /// <summary>Asks whether to reload although these core windows are open; false cancels the reload.</summary>
        bool ConfirmCloseOpenWindows(string[] titles);

        /// <summary>Shows the reload progress window; may return null when it cannot be shown (the reload goes on).</summary>
        IReloadProgress ShowProgress(string currentVersion, string currentReleaseDate);

        void ShowError(string message, string title);
    }

    internal interface IReloadProgress
    {
        /// <summary>Shows a stage (1 closing, 2 finding, 3 loading, 4 done).</summary>
        void ShowStage(string message, int step);

        void SetTarget(string version, string releaseDate);
        void Close();
    }

    /// <summary>Loads and unloads the core (AppDomain loader in the add-in).</summary>
    internal interface ICoreLoader
    {
        IXLEdgeAddin Load(IXLEdgeContext context, ResolvedRelease release);
        void Unload(IXLEdgeContext context);
    }

    /// <summary>Chooses and remembers the release to load (UpdateBootstrapper in the add-in).</summary>
    internal interface IReleaseCatalog
    {
        ResolvedRelease ResolveVersionToLoad(IXLEdgeContext context, string source);
        void RememberLoaded(IXLEdgeContext context, string folderName);
    }

    internal sealed class DomainCoreLoader : ICoreLoader
    {
        private readonly AddinDomainLoader _loader;
        public DomainCoreLoader(AddinDomainLoader loader) { _loader = loader; }
        public IXLEdgeAddin Load(IXLEdgeContext context, ResolvedRelease release) => _loader.Load(context, release);
        public void Unload(IXLEdgeContext context) => _loader.Unload(context);
    }

    internal sealed class BootstrapperReleaseCatalog : IReleaseCatalog
    {
        private readonly UpdateBootstrapper _bootstrapper;
        public BootstrapperReleaseCatalog(UpdateBootstrapper bootstrapper) { _bootstrapper = bootstrapper; }
        public ResolvedRelease ResolveVersionToLoad(IXLEdgeContext context, string source) => _bootstrapper.ResolveVersionToLoad(context, source);
        public void RememberLoaded(IXLEdgeContext context, string folderName) => _bootstrapper.RememberLoaded(context, folderName);
    }
}
