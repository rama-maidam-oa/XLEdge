using System.Text.Json;
using XLEdge.Contracts;

namespace XLEdge.Addin.Core.Infrastructure
{
    /// <summary>
    /// The login outlives a hot reload: at Shutdown the core leaves it in the shell's session values (they live as long as this
    /// Excel), and the next core takes it back at Initialize - so its panes open logged in, as 11.1.2's did all session. A logout
    /// clears XLEdgeAppState's login first (LogoffFromXLEdgeAddin), so it leaves nothing; Excel closing ends the session values.
    /// </summary>
    internal static class SessionLogin
    {
        internal const string Key = "XLEdge.Login";

        private sealed class Login
        {
            public string LoginUrl { get; set; }
            public string LoginUrlName { get; set; }
            public string LoginToken { get; set; }
            public string LoginUserName { get; set; }
            public bool IsLoginCompleted { get; set; }
            public bool LoginFromGLSense { get; set; }
            public bool LoginSentToGLSense { get; set; }
        }

        public static void Keep(IXLEdgeContext context)
        {
            XLEdgeAppState state = XLEdgeAppState.Instance;
            bool loggedIn = state.IsLoginCompleted || !string.IsNullOrEmpty(state.LoginUrl);
            context?.SetSessionValue(Key, !loggedIn ? null : JsonSerializer.Serialize(new Login
            {
                LoginUrl = state.LoginUrl, LoginUrlName = state.LoginUrlName, LoginToken = state.LoginToken, LoginUserName = state.LoginUserName,
                IsLoginCompleted = state.IsLoginCompleted, LoginFromGLSense = state.LoginFromGLSense, LoginSentToGLSense = state.LoginSentToGLSense,
            }));
        }

        /// <summary>The previous core's login, if it left one; true when there was one.</summary>
        public static bool Restore(IXLEdgeContext context)
        {
            string json = context?.GetSessionValue(Key);
            if (string.IsNullOrEmpty(json)) return false;
            Login login = JsonSerializer.Deserialize<Login>(json);
            if (login == null) return false;
            XLEdgeAppState state = XLEdgeAppState.Instance;
            state.LoginUrl = login.LoginUrl ?? string.Empty;
            state.LoginUrlName = login.LoginUrlName ?? string.Empty;
            state.LoginToken = login.LoginToken ?? string.Empty;
            state.LoginUserName = login.LoginUserName ?? string.Empty;
            state.IsLoginCompleted = login.IsLoginCompleted;
            state.LoginFromGLSense = login.LoginFromGLSense;
            state.LoginSentToGLSense = login.LoginSentToGLSense;
            return true;
        }
    }
}
