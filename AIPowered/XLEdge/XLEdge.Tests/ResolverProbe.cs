using System;
using System.Linq;
using XLEdge.Contracts;
using XLEdge.Loader.Core;

namespace XLEdge.Tests
{
    /// <summary>Runs RemoteLoader's resolver inside a throwaway AppDomain so the files it loads never reach the test domain.</summary>
    public sealed class ResolverProbe : MarshalByRefObject
    {
        public bool IsLoaded(string simpleName) => AppDomain.CurrentDomain.GetAssemblies().Any(a => a.GetName().Name == simpleName);

        public string Resolve(string requestedName, string releaseFolder, string sharedFolder, ILogger logger) =>
            RemoteLoader.ResolveFromFolders(requestedName, releaseFolder, sharedFolder, logger)?.Location;

        public override object InitializeLifetimeService() => null;
    }
}
