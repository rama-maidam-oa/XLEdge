using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class XLEdgeContextTests
    {
        [TestMethod]
        public void Context_ShellVersionLeaseAndDefaults()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                string expected = FileVersionInfo.GetVersionInfo(typeof(XLEdgeContext).Assembly.Location).FileVersion;
                Assert.AreEqual(expected, context.ShellVersion);
                Assert.AreEqual("11.1.3.0", context.ShellVersion);
                Assert.IsNull(context.InitializeLifetimeService());
                Assert.IsFalse(context.DebugMode);
                Assert.IsTrue(Directory.Exists(context.Paths.Logs), "paths ensured");
                Assert.IsNotNull(context.Logger);
                Assert.IsNull(context.RibbonController);
            }
        }

        [TestMethod]
        public void Context_WithoutGLSense_ReturnsNullAndLogs()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                Assert.IsNull(context.GetGLSenseAddinInstance());
            }
        }

        public sealed class FakeComAddin
        {
            public string ProgId { get; set; } = "GLSense.AddinModule";
            public bool Connect { get; set; } = true;
            public object Object { get; set; } = new object();
        }

        public sealed class FakeComAddins
        {
            public FakeComAddin Only { get; } = new FakeComAddin();
            public int Count => 1;
            public object Item(object index) => Only;
        }

        public sealed class FakeExcel
        {
            public FakeComAddins COMAddIns { get; } = new FakeComAddins();
        }

        [TestMethod]
        public void Context_GLSense_IsLookedUpEveryTime()
        {
            // GLSense can be disconnected or reloaded during the session: a cached object would go stale.
            using (var t = new TempFolder())
            {
                var excel = new FakeExcel();
                using (var context = new XLEdgeContext(excel, new PathProvider(t.Combine("inst"), t.Combine("data"))))
                {
                    object first = excel.COMAddIns.Only.Object;
                    Assert.AreSame(first, context.GetGLSenseAddinInstance());
                    excel.COMAddIns.Only.Connect = false;
                    Assert.IsNull(context.GetGLSenseAddinInstance());
                    excel.COMAddIns.Only.Connect = true;
                    excel.COMAddIns.Only.Object = new object();
                    Assert.AreSame(excel.COMAddIns.Only.Object, context.GetGLSenseAddinInstance());
                }
            }
        }

        /// <summary>Stands in for GLSense's AddinModule: the two methods XLEdge 11.1.2 calls, late-bound.</summary>
        public sealed class GLSenseRecorder
        {
            public readonly System.Collections.Generic.List<string> Calls = new System.Collections.Generic.List<string>();
            public Exception Throw { get; set; }
            public void GetGLCubeInformation(string token, string url, string userName)
            {
                if (Throw != null) throw Throw;
                Calls.Add($"GetGLCubeInformation|{token}|{url}|{userName}");
            }
            public void LogoutSession()
            {
                if (Throw != null) throw Throw;
                Calls.Add("LogoutSession");
            }
        }

        [TestMethod]
        public void Context_GLSense_IsConnected_FollowsConnect()
        {
            using (var t = new TempFolder())
            {
                var excel = new FakeExcel();
                using (var context = new XLEdgeContext(excel, new PathProvider(t.Combine("inst"), t.Combine("data"))))
                {
                    Assert.IsTrue(context.IsGLSenseConnected());
                    excel.COMAddIns.Only.Connect = false;
                    Assert.IsFalse(context.IsGLSenseConnected());
                    excel.COMAddIns.Only.Connect = true;
                    excel.COMAddIns.Only.ProgId = "Other.AddinModule";
                    Assert.IsFalse(context.IsGLSenseConnected());
                }
            }
        }

        [TestMethod]
        public void Context_CallGLSense_InvokesTheConnectedAddin()
        {
            using (var t = new TempFolder())
            {
                var excel = new FakeExcel();
                var glSense = new GLSenseRecorder();
                excel.COMAddIns.Only.Object = glSense;
                using (var context = new XLEdgeContext(excel, new PathProvider(t.Combine("inst"), t.Combine("data"))))
                {
                    context.CallGLSense("GetGLCubeInformation", new object[] { "tok", "http://x", "u" });
                    context.CallGLSense("LogoutSession", new object[0]);
                    excel.COMAddIns.Only.Connect = false;
                    context.CallGLSense("LogoutSession", null);
                }
                CollectionAssert.AreEqual(new[] { "GetGLCubeInformation|tok|http://x|u", "LogoutSession" }, glSense.Calls,
                    "both calls in order; nothing while GLSense is disconnected");
            }
        }

        [TestMethod]
        public void Context_CallGLSense_FailureComesBackAsInvalidOperation()
        {
            using (var t = new TempFolder())
            {
                var excel = new FakeExcel();
                excel.COMAddIns.Only.Object = new GLSenseRecorder { Throw = new System.Runtime.InteropServices.COMException("boom") };
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                using (var context = new XLEdgeContext(excel, paths))
                {
                    var ex = Assert.ThrowsExactly<InvalidOperationException>(
                        () => context.CallGLSense("GetGLCubeInformation", new object[] { "tok", "http://x", "u" }));
                    StringAssert.Contains(ex.Message, "GetGLCubeInformation");
                    StringAssert.Contains(ex.Message, "boom");
                    Assert.IsNull(ex.InnerException, "only serializable data crosses back to the core");
                }
                string log = string.Join("\n", Directory.GetFiles(paths.Logs).Select(File.ReadAllText));
                StringAssert.Contains(log, "GetGLCubeInformation");
                Assert.IsFalse(log.Contains("tok"), "the token is never logged");
            }
        }

        /// <summary>Uses a context from another AppDomain, the way a core does (through the remoting proxy).</summary>
        public sealed class RemoteSessionUser : MarshalByRefObject
        {
            public string SetThenGet(XLEdge.Contracts.IXLEdgeContext context, string key, string value)
            {
                context.SetSessionValue(key, value);
                return context.GetSessionValue(key);
            }
        }

        [TestMethod]
        public void SessionBag_SetGetOverwriteRemove_AndNullKeys()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                Assert.IsNull(context.GetSessionValue("a"));
                context.SetSessionValue("a", "1");
                Assert.AreEqual("1", context.GetSessionValue("a"));
                context.SetSessionValue("a", "2");
                Assert.AreEqual("2", context.GetSessionValue("a"));
                Assert.IsNull(context.GetSessionValue("A"), "keys are case-sensitive");
                context.SetSessionValue("a", null);
                Assert.IsNull(context.GetSessionValue("a"));
                context.SetSessionValue(null, "x");
                Assert.IsNull(context.GetSessionValue(null));

                context.SetSessionValue("b", "kept");
                ((XLEdge.Hosting.ISessionBag)context).ClearSession();
                Assert.IsNull(context.GetSessionValue("b"));
            }
        }

        [TestMethod]
        public void SessionBag_WorksFromTheCoresAppDomain()
        {
            using (var t = new TempFolder())
            using (var context = new XLEdgeContext(new object(), new PathProvider(t.Combine("inst"), t.Combine("data"))))
            {
                var domain = AppDomain.CreateDomain("SessionUser", null, new AppDomainSetup { ApplicationBase = System.IO.Path.GetDirectoryName(typeof(RemoteSessionUser).Assembly.Location) });
                try
                {
                    var user = (RemoteSessionUser)domain.CreateInstanceAndUnwrap(typeof(RemoteSessionUser).Assembly.FullName, typeof(RemoteSessionUser).FullName);
                    Assert.AreEqual("v", user.SetThenGet(context, "k", "v"));
                    Assert.AreEqual("v", context.GetSessionValue("k"), "kept on the shell's side");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
                Assert.AreEqual("v", context.GetSessionValue("k"), "and survives the other domain's unload");
            }
        }

        [TestMethod]
        public void Context_Dispose_ReleasesTheLogFile()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                var context = new XLEdgeContext(new object(), paths);
                context.Logger.LogInfo("hello");
                context.Dispose();
                foreach (var file in Directory.GetFiles(paths.Logs)) File.Delete(file);
            }
        }

        [TestMethod]
        public void RibbonControlIds_StateListsMatchTodaysAddin()
        {
            CollectionAssert.AreEquivalent(new[] { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide", "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp", "RibEdgeBroadcast" },
                RibbonControlIds.LoggedOutDisabled.ToArray());
            CollectionAssert.AreEquivalent(new[] { "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout" }, RibbonControlIds.LoggedOutEnabled.ToArray());
            CollectionAssert.AreEquivalent(new[] { "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll", "RibEdgeParamRefresh", "RibEdgeShowHide",
                "RibEdgeIncludeOutputData", "RibEdgeOptions", "RibEdgeHelp", "RibEdgeLogin", "RibEdgeDebug", "RibControlSheet", "RibEdgeAbout", "RibEdgeBroadcast" },
                RibbonControlIds.NoPermissionDisabled.ToArray());
        }
    }
}
