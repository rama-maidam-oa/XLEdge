using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static XLEdge.Tests.CoreWindowsTests;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 2f: XLEdge 11.1.2's login/logout sync with GLSense, in the core. GLSense as the shell sees it is FakeContext
    /// (contract 5: IsGLSenseConnected, CallGLSense); the calls GLSense makes into XLEdge come through the entry
    /// (OnLoginFromGLSense, OnLogoutFromGLSense), as the shell's InvokedFromGLSense / LogoffFromAddin forward them.
    /// </summary>
    [TestClass]
    public class CoreGLSenseTests
    {
        private const string State = "XLEdge.XLEdgeAppState";
        private const string Module = "XLEdge.AddinModule";
        private static readonly TimeSpan Wait = TimeSpan.FromSeconds(15);

        private static void WithCore(Action<CoreProbe, FakeContext> use, Action<FakeContext> before = null) => Win32Test.OnSta(() =>
        {
            using (var t = new TempFolder())
            {
                FakeContext context = NewContext(t);
                context.GLSenseConnected = true;
                before?.Invoke(context);
                CoreOnThisThread(context, p =>
                {
                    p.AnswerMessages();
                    try { use(p, context); }
                    finally { p.StopAnsweringMessages(); }
                });
            }
        });

        private static object AppState(CoreProbe p, string property) => p.OnSingleton(State, "Instance", property);

        private static void NotifyOwnLogin(CoreProbe p) => p.CallOnSingleton(Module, "CurrentInstance", "NotifyGLSenseOfLogin", "tok", "http://x", "u");

        private static string[] RibbonCalls(FakeContext context) => ((CoreRibbonFacadeTests.FakeRibbon)context.RibbonController).Calls;

        /// <summary>Logout as GLSense asks for it (11.1.2's LogoffFromAddin: XLEdge's own logout sequence), waited for.</summary>
        private static void LogoutFromGLSense(CoreProbe p)
        {
            p.EntryLogoutFromGLSense();
            Assert.IsTrue(Win32Test.PumpUntil(() => !(bool)AppState(p, "IsLoginCompleted") && !p.EntryIsBusy(), Wait), "logged out");
        }

        [TestMethod]
        public void OwnLogin_TellsGLSenseOnce()
        {
            WithCore((p, context) =>
            {
                NotifyOwnLogin(p);
                NotifyOwnLogin(p);
                CollectionAssert.AreEqual(new[] { "GetGLCubeInformation|tok|http://x|u" }, context.GLSenseCalls);
                Assert.AreEqual(true, AppState(p, "LoginSentToGLSense"));
            });
        }

        [TestMethod]
        public void OwnLogin_GLSenseNotThere_TriesAgainNextTime()
        {
            WithCore((p, context) =>
            {
                context.GLSenseConnected = false;
                NotifyOwnLogin(p);
                Assert.AreEqual(false, AppState(p, "LoginSentToGLSense"), "nothing sent yet");
                context.GLSenseConnected = true;
                NotifyOwnLogin(p);
                CollectionAssert.AreEqual(new[] { "GetGLCubeInformation|tok|http://x|u" }, context.GLSenseCalls);
            });
        }

        [TestMethod]
        public void OwnLogin_NotRepeatedBack_ForALoginFromGLSense()
        {
            WithCore((p, context) =>
            {
                p.OnSingleton(State, "Instance", "LoginFromGLSense", true);
                NotifyOwnLogin(p);
                CollectionAssert.AreEqual(new string[0], context.GLSenseCalls);
            });
        }

        [TestMethod]
        public void OwnLogin_WithoutAToken_IsNotSent()
        {
            WithCore((p, context) =>
            {
                p.CallOnSingleton(Module, "CurrentInstance", "NotifyGLSenseOfLogin", "", "http://x", "u");
                CollectionAssert.AreEqual(new string[0], context.GLSenseCalls);
            });
        }

        [TestMethod]
        public void LoginFromGLSense_LogsXLEdgeIn_AsInXLEdge11_1_2()
        {
            WithCore((p, context) =>
            {
                p.EntryLoginFromGLSense("Prod", "http://x", "tok", "u", true);
                Assert.IsTrue(Win32Test.PumpUntil(() => (bool)AppState(p, "IsLoginCompleted"), Wait), "logged in");
                Assert.AreEqual("Prod", AppState(p, "LoginUrlName"));
                Assert.AreEqual("http://x", AppState(p, "LoginUrl"));
                Assert.AreEqual("tok", AppState(p, "LoginToken"));
                Assert.AreEqual("u", AppState(p, "LoginUserName"));
                Assert.AreEqual(true, AppState(p, "LoginFromGLSense"));
                string[] calls = RibbonCalls(context);
                foreach (string expected in new[] { "Visible:RibEdgeLogin=False", "Visible:RibEdgeLogout=True", "Label:RibEdgeLogout=Prod",
                                                    "Enabled:RibEdgeBroadcast=True", "Enabled:RibEdgeDialogBoxLauncher=True",
                                                    "Pressed:RibEdgeDebug=False", "Label:RibLoginURL=http://x" })
                    CollectionAssert.Contains(calls, expected, string.Join(Environment.NewLine, calls));
                CollectionAssert.AreEqual(new string[0], context.GLSenseCalls, "GLSense's own login is not sent back to it");
            });
        }

        [TestMethod]
        public void LoginFromGLSense_WithoutXLEdgePermission_StaysLoggedOut()
        {
            WithCore((p, context) =>
            {
                p.EntryLoginFromGLSense("Prod", "http://x", "tok", "u", false);
                string[] expected = { "Enabled:RibEdgeLogin=False", "Enabled:RibControlSheet=False", "Enabled:RibEdgeDebug=False",
                                      "Enabled:RibEdgeIncludeOutputData=False", "Enabled:RibEdgeBroadcast=False" };
                Assert.IsTrue(Win32Test.PumpUntil(() => expected.All(RibbonCalls(context).Contains), Wait), string.Join(Environment.NewLine, RibbonCalls(context)));
                Assert.AreEqual(false, AppState(p, "IsLoginCompleted"));
                Assert.IsTrue(string.IsNullOrEmpty((string)AppState(p, "LoginToken")), "GLSense's token is not taken");
            });
        }

        [TestMethod]
        public void LogoutFromGLSense_LogsXLEdgeOut_AndTellsGLSense()
        {
            WithCore((p, context) =>
            {
                p.EntryLoginFromGLSense("Prod", "http://x", "tok", "u", true);
                Assert.IsTrue(Win32Test.PumpUntil(() => (bool)AppState(p, "IsLoginCompleted"), Wait), "logged in");
                LogoutFromGLSense(p);
                Assert.AreEqual(false, AppState(p, "LoginFromGLSense"));
                Assert.AreEqual(string.Empty, AppState(p, "LoginToken"));
                Assert.AreEqual("Visible:RibEdgeLogout=False", RibbonCalls(context).Last(c => c.StartsWith("Visible:RibEdgeLogout")), "LoggedOut ribbon");
                CollectionAssert.AreEqual(new[] { "LogoutSession" }, context.GLSenseCalls, "11.1.2 tells GLSense, as for any XLEdge logout");
            });
        }

        [TestMethod]
        public void LogoutFromGLSense_KeepsTheCoreBusyUntilItEnds()
        {
            // A reload waits while the core is busy: during a logoff GLSense started, it must not save the half-done login
            // (IsLoginCompleted is reset only at the end) for the next core to start LoggedIn while GLSense is logged out.
            WithCore((p, context) =>
            {
                p.EntryLoginFromGLSense("Prod", "http://x", "tok", "u", true);
                Assert.IsTrue(Win32Test.PumpUntil(() => (bool)AppState(p, "IsLoginCompleted"), Wait), "logged in");
                p.EntryLogoutFromGLSense();
                Assert.IsTrue(p.EntryIsBusy(), "busy while the logoff runs");
                Assert.IsTrue(Win32Test.PumpUntil(() => !(bool)AppState(p, "IsLoginCompleted") && !p.EntryIsBusy(), Wait), "logged out, then not busy");
            });
        }

        [TestMethod]
        public void Logout_TellsGLSense_OnlyAfterASharedLogin()
        {
            WithCore((p, context) =>
            {
                p.OnSingleton(State, "Instance", "IsLoginCompleted", true);
                LogoutFromGLSense(p);
                CollectionAssert.AreEqual(new string[0], context.GLSenseCalls, "never shared: GLSense is not told");

                p.OnSingleton(State, "Instance", "IsLoginCompleted", true);
                NotifyOwnLogin(p);
                LogoutFromGLSense(p);
                CollectionAssert.AreEqual(new[] { "GetGLCubeInformation|tok|http://x|u", "LogoutSession" }, context.GLSenseCalls);
            });
        }

        [TestMethod]
        public void GLSenseFailing_DoesNotStopXLEdge()
        {
            WithCore((p, context) =>
            {
                context.GLSenseThrows = true;
                p.OnSingleton(State, "Instance", "IsLoginCompleted", true);
                NotifyOwnLogin(p);
                Assert.AreEqual(true, AppState(p, "LoginSentToGLSense"), "as 11.1.2: marked sent before the call");
                LogoutFromGLSense(p);
                var log = (RecordingLogger)context.Logger;
                Assert.IsTrue(log.Contains("Failed to invoke GetGLCubeInformation on GLSense addin."), log.Dump());
                Assert.IsTrue(log.Contains("Failed to invoke LogoutSession on GLSense addin."), log.Dump());
            });
        }

        [TestMethod]
        public void Reload_KeepsTellingGLSense()
        {
            // The previous core sent its login to GLSense; this one must still tell GLSense at logout (Ruling 2).
            WithCore((p, context) =>
            {
                Assert.AreEqual(true, AppState(p, "IsLoginCompleted"), "the previous core's login is back");
                LogoutFromGLSense(p);
                CollectionAssert.AreEqual(new[] { "LogoutSession" }, context.GLSenseCalls);
            },
            before: context => context.SetSessionValue("XLEdge.Login",
                "{\"LoginUrl\":\"http://x\",\"LoginToken\":\"tok\",\"IsLoginCompleted\":true,\"LoginSentToGLSense\":true}"));
        }
    }
}
