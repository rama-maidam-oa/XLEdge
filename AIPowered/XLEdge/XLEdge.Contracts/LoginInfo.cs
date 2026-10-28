using System;

namespace XLEdge.Contracts
{
    /// <summary>
    /// Current login state of the XLEdge core, returned by IXLEdgeAddin.GetLoginInfo so the
    /// shell (e.g. the online reload window) can call the server without the core's state
    /// object crossing the AppDomain boundary. [Serializable] because it is passed by value.
    /// </summary>
    [Serializable]
    public class LoginInfo
    {
        public string LoginUrl { get; set; }
        public string LoginToken { get; set; }
        public string LoginUserName { get; set; }
        public bool IsLoggedIn { get; set; }
    }
}
