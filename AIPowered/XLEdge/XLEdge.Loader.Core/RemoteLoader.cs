using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using XLEdge.Contracts;

namespace XLEdge.Loader.Core
{
    /// <summary>
    /// Created inside the core's AppDomain by <see cref="AddinDomainLoader"/>: loads XLEdge.Addin.Core.dll from the release
    /// folder and returns its entry point wrapped in a <see cref="RemoteAddinProxy"/>.
    /// </summary>
    public class RemoteLoader : MarshalByRefObject
    {
        internal const string CoreAssemblyFile = "XLEdge.Addin.Core.dll";

        // Per AppDomain (a fresh domain per load), read by the AssemblyResolve handler.
        private static readonly object ResolverLock = new object();
        private static string _releaseFolder;
        private static string _sharedFolder;
        private static ILogger _logger;
        private static bool _resolverRegistered;

        public IXLEdgeAddin Create(string releaseFolder, string sharedDependenciesPath, ILogger logger)
        {
            lock (ResolverLock)
            {
                _releaseFolder = releaseFolder;
                _sharedFolder = sharedDependenciesPath;
                _logger = logger;
                if (!_resolverRegistered)
                {
                    AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;
                    _resolverRegistered = true;
                }
            }

            string path = Path.Combine(releaseFolder, CoreAssemblyFile);
            try
            {
                if (!File.Exists(path)) throw new FileNotFoundException("XLEdge core not found: " + path, path);

                Assembly core = Assembly.LoadFrom(path);
                Type entryType = FindEntryType(core, logger)
                    ?? throw new InvalidOperationException("No IXLEdgeAddin entry type found in " + CoreAssemblyFile);
                var instance = (IXLEdgeAddin)Activator.CreateInstance(entryType);
                return new RemoteAddinProxy(instance);
            }
            catch (Exception ex)
            {
                logger?.LogError($"RemoteLoader.Create: failed to create the XLEdge core from '{path}'.", ex.ToString());
                if (CrossDomain.NeedsWrapping(ex)) throw CrossDomain.Wrap(ex);
                throw;
            }
        }

        public override object InitializeLifetimeService() => null;

        private static Type FindEntryType(Assembly core, ILogger logger)
        {
            IEnumerable<Type> types;
            try
            {
                types = core.GetTypes();
            }
            catch (ReflectionTypeLoadException ex)
            {
                foreach (var loaderException in ex.LoaderExceptions.Where(e => e != null))
                    logger?.LogError("RemoteLoader.Create: type load failure in the XLEdge core: " + loaderException.Message);
                types = ex.Types.Where(t => t != null);
            }
            return types.FirstOrDefault(t => typeof(IXLEdgeAddin).IsAssignableFrom(t) && t.IsClass && !t.IsAbstract);
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            string release, shared;
            ILogger logger;
            lock (ResolverLock)
            {
                release = _releaseFolder;
                shared = _sharedFolder;
                logger = _logger;
            }
            return ResolveFromFolders(args.Name, release, shared, logger);
        }

        /// <summary>
        /// Resolves an assembly the core's AppDomain could not find by itself: an assembly already loaded (exact identity, then
        /// same simple name), else {name}.dll / {name}.exe from the release folder, then from SharedDependencies. A file that
        /// fails to load is skipped in favour of the next candidate. Satellite resource assemblies are not resolved.
        /// </summary>
        internal static Assembly ResolveFromFolders(string requestedName, string releaseFolder, string sharedFolder, ILogger logger)
        {
            AssemblyName requested;
            try
            {
                requested = new AssemblyName(requestedName);
            }
            catch (Exception ex)
            {
                logger?.LogError($"RemoteLoader.ResolveFromFolders: invalid assembly name '{requestedName}'.", ex.ToString());
                return null;
            }
            string simpleName = requested.Name;
            // Satellite resources and pre-generated XmlSerializers assemblies are probed routinely and are optional.
            if (string.IsNullOrEmpty(simpleName) ||
                simpleName.EndsWith(".resources", StringComparison.OrdinalIgnoreCase) ||
                simpleName.EndsWith(".XmlSerializers", StringComparison.OrdinalIgnoreCase)) return null;

            var loaded = AppDomain.CurrentDomain.GetAssemblies();
            Assembly exact = loaded.FirstOrDefault(a => string.Equals(a.FullName, requestedName, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact;
            Assembly sameName = loaded.FirstOrDefault(a => string.Equals(a.GetName().Name, simpleName, StringComparison.OrdinalIgnoreCase));
            if (sameName != null)
            {
                if (requested.Version != null && sameName.GetName().Version != requested.Version)
                    logger?.LogDebug($"RemoteLoader.ResolveFromFolders: '{requestedName}' answered with the loaded '{sameName.FullName}'.");
                return sameName;
            }

            foreach (string candidate in Candidates(simpleName, releaseFolder, sharedFolder))
            {
                if (!File.Exists(candidate)) continue;
                try
                {
                    Assembly assembly = Assembly.LoadFrom(candidate);
                    return assembly;
                }
                catch (Exception ex)
                {
                    logger?.LogWarn($"RemoteLoader.ResolveFromFolders: could not load '{candidate}': {ex.Message}");
                }
            }

            logger?.LogWarn($"RemoteLoader.ResolveFromFolders: could not resolve '{requestedName}' from release '{releaseFolder}' or shared dependencies '{sharedFolder}'.");
            return null;
        }

        private static IEnumerable<string> Candidates(string simpleName, string releaseFolder, string sharedFolder)
        {
            foreach (string extension in new[] { ".dll", ".exe" })
                foreach (string folder in new[] { releaseFolder, sharedFolder })
                    if (!string.IsNullOrWhiteSpace(folder))
                        yield return Path.Combine(folder, simpleName + extension);
        }

        /// <summary>
        /// What the shell holds. Its type lives in XLEdge.Loader.Core (already loaded in the shell's domain), so the shell's
        /// transparent proxy never needs XLEdge.Addin.Core.dll to be resolvable there.
        /// </summary>
        public sealed class RemoteAddinProxy : MarshalByRefObject, IXLEdgeAddin
        {
            private readonly IXLEdgeAddin _inner;

            public RemoteAddinProxy(IXLEdgeAddin inner)
            {
                _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            }

            public void Initialize(IXLEdgeContext context) => Call(() => _inner.Initialize(context));
            public void OnRibbonAction(string action, object parameter) => Call(() => _inner.OnRibbonAction(action, parameter));
            public bool OnExcelEvent(string eventName, object[] args) => Call(() => _inner.OnExcelEvent(eventName, args));
            public void Shutdown() => Call(() => _inner.Shutdown());
            public IntPtr CreateTaskPaneContent(IntPtr excelWindowHandle) => Call(() => _inner.CreateTaskPaneContent(excelWindowHandle));
            public void CloseTaskPaneContent(IntPtr excelWindowHandle) => Call(() => _inner.CloseTaskPaneContent(excelWindowHandle));
            public void OnTaskPaneShown(IntPtr excelWindowHandle) => Call(() => _inner.OnTaskPaneShown(excelWindowHandle));
            public void OnLoginFromGLSense(string instanceName, string loginUrl, string token, string userName, bool hasXLEdgePermission) =>
                Call(() => _inner.OnLoginFromGLSense(instanceName, loginUrl, token, userName, hasXLEdgePermission));
            public void OnLogoutFromGLSense() => Call(() => _inner.OnLogoutFromGLSense());
            public LoginInfo GetLoginInfo() => Call(() => _inner.GetLoginInfo());
            public string[] GetOpenWindowTitles() => Call(() => _inner.GetOpenWindowTitles());
            public bool IsBusy() => Call(() => _inner.IsBusy());

            // Exceptions of core-only types are replaced by a framework exception carrying their text (see CrossDomain).
            private static void Call(Action call)
            {
                try { call(); }
                catch (Exception ex) when (CrossDomain.NeedsWrapping(ex)) { throw CrossDomain.Wrap(ex); }
            }

            private static T Call<T>(Func<T> call)
            {
                try { return call(); }
                catch (Exception ex) when (CrossDomain.NeedsWrapping(ex)) { throw CrossDomain.Wrap(ex); }
            }

            public override object InitializeLifetimeService() => null;
        }
    }
}
