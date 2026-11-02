using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// What 11.1.2's GetGLSenseAddinObject returned - GLSense's add-in object, with the two methods XLEdge calls on it late-bound
    /// (GetType().InvokeMember). Each call goes to the shell (contract 5), which invokes GLSense's real COM object in its own
    /// AppDomain, on Excel's thread. A failed call throws, as the COM call did in 11.1.2.
    /// </summary>
    internal sealed class GLSenseAddin
    {
        /// <summary>XLEdge logged in by itself: GLSense takes the session (it loads its cubes with it).</summary>
        public void GetGLCubeInformation(string token, string url, string userName) =>
            Call(nameof(GetGLCubeInformation), new object[] { token, url, userName });

        /// <summary>XLEdge logged out: GLSense logs out too.</summary>
        public void LogoutSession() => Call(nameof(LogoutSession), new object[0]);

        private static void Call(string method, object[] args)
        {
            IXLEdgeContext context = CoreServices.Context;
            if (context == null) return;
            ExcelThreadContext.OnExcelThread(() => context.CallGLSense(method, args));
        }
    }
}
