using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// The core works on Excel through its own RCW, built from the shell's IDispatch pointer - not through a remoting proxy of
    /// the shell's RCW (thousands of COM calls per report; the proxy path is where GLSense hit CrossAppDomainSink cast
    /// failures after awaits). Scripting.Dictionary, a COM object every Windows has, stands in for Excel.
    /// </summary>
    [TestClass]
    public class CoreExcelTests
    {
        /// <summary>A context whose Excel is a real COM object, handed out the way the shell does (AddRef'ed IDispatch).</summary>
        private sealed class ComContext : MarshalByRefObject, IXLEdgeContext
        {
            private object _com;
            public ComContext(IPathProvider paths, object com) { Paths = paths; _com = com; Logger = new RecordingLogger(); }
            public override object InitializeLifetimeService() => null;
            public object ExcelApp => _com;
            /// <summary>From now on hand out another object (what a reconnected Excel gives after a disconnect).</summary>
            public void SwitchTo(object com) => _com = com;
            public IntPtr GetExcelDispatch() => _com == null ? IntPtr.Zero : Marshal.GetIDispatchForObject(_com);
            public ILogger Logger { get; }
            public IPathProvider Paths { get; }
            public IRibbonController RibbonController { get; private set; }
            public IntPtr ExcelHandle { get; set; }
            public bool DebugMode { get; set; }
            public string Version { get; set; }
            public string ReleaseDate { get; set; }
            public string ActiveFolderName { get; set; }
            public string ShellVersion => "11.1.2.0";
            public void SetRibbonController(IRibbonController controller) => RibbonController = controller;
            public object GetGLSenseAddinInstance() => null;
            public bool IsGLSenseConnected() => false;
            public void CallGLSense(string method, object[] args) { }
            public string GetSessionValue(string key) => null;
            public void SetSessionValue(string key, string value) { }
        }

        private const string CoreServices = "XLEdge.Addin.Core.Infrastructure.CoreServices";

        private static object NewDictionary() => Activator.CreateInstance(Type.GetTypeFromProgID("Scripting.Dictionary", throwOnError: true));

        /// <summary>The COM object's reference count (AddRef + Release returns it without changing it).</summary>
        private static int RefCount(object com)
        {
            IntPtr unknown = Marshal.GetIUnknownForObject(com);
            try
            {
                Marshal.AddRef(unknown);
                return Marshal.Release(unknown) - 1; // minus the GetIUnknownForObject reference released below
            }
            finally
            {
                Marshal.Release(unknown);
            }
        }

        [TestMethod]
        public void CoreExcel_IsANativeComObject_OfTheSameExcel()
        {
            using (var t = new TempFolder())
            {
                object dictionary = NewDictionary();
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), dictionary);
                var (domain, probe) = CoreProbe.Create("CoreExcel");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    Assert.AreEqual("com", probe.DescribeExcel(), "the core must hold its own RCW, not a remoting proxy");
                    probe.InvokeOnExcel("Add", "key", "value");
                    Assert.AreEqual(1, (int)dictionary.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, dictionary, null),
                        "the core's calls reach the shell's Excel");
                    probe.CallStatic(CoreServices, "Shutdown");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void Shutdown_ReleasesTheCoresReference()
        {
            // A reference the core keeps after Shutdown would keep EXCEL.EXE alive after Excel is closed.
            using (var t = new TempFolder())
            {
                object dictionary = NewDictionary();
                int before = RefCount(dictionary);
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), dictionary);
                var (domain, probe) = CoreProbe.Create("CoreExcelRelease");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    Assert.AreEqual("com", probe.DescribeExcel());
                    Assert.IsTrue(RefCount(dictionary) > before, "precondition: the core holds a reference");
                    probe.CallStatic(CoreServices, "Shutdown");
                    Assert.AreEqual(before, RefCount(dictionary), "the core's reference is released at Shutdown");
                    Assert.AreEqual("null", probe.DescribeExcel(), "no Excel after Shutdown");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        /// <summary>The shell's real context (XLEdgeContext), whose GetExcelDispatch the core relies on.</summary>
        private static XLEdgeContext ShellContext(TempFolder t, object excel) => new XLEdgeContext(excel, new PathProvider(t.Combine("inst"), t.Combine("data")));

        [TestMethod]
        public void ShellContext_HandsOutAnAddRefedPointer_ThatTheCoreReleases()
        {
            // A non-AddRef'ed pointer from the shell would make the core over-release Excel's Application (crashes Excel).
            using (var t = new TempFolder())
            using (var context = ShellContext(t, NewDictionary()))
            {
                object dictionary = context.ExcelApp;
                int before = RefCount(dictionary);
                var (domain, probe) = CoreProbe.Create("ShellContextDispatch");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    Assert.AreEqual("com", probe.DescribeExcel());
                    probe.CallStatic(CoreServices, "Shutdown");
                    Assert.AreEqual(before, RefCount(dictionary));
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void ShellContext_WithoutACOMExcel_HandsOutNothing()
        {
            using (var t = new TempFolder())
            using (var context = ShellContext(t, new object()))
            {
                Assert.AreEqual(IntPtr.Zero, context.GetExcelDispatch());
            }
        }

        [TestMethod]
        public void Initialize_TakesExcelAtOnce_OnTheCallingThread()
        {
            // Initialize runs on Excel's main thread; an RCW made later on another thread would be bound to that thread.
            using (var t = new TempFolder())
            {
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), NewDictionary());
                var (domain, probe) = CoreProbe.Create("CoreExcelEager");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    Assert.IsTrue(probe.HasStatic(CoreServices, "ExcelAppIfCreated"), "created by Initialize, not on first use");
                    probe.CallStatic(CoreServices, "Shutdown");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void ResetExcel_KeepsReferencesPortedCodeHolds_AndShutdownStillReleasesEverything()
        {
            // Ported code keeps "var app = XLApp.App" locals; dropping a disconnected RCW must not destroy them.
            using (var t = new TempFolder())
            {
                object dictionary = NewDictionary();
                int before = RefCount(dictionary);
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), dictionary);
                var (domain, probe) = CoreProbe.Create("CoreExcelReset");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    probe.KeepExcel();
                    object reconnected = NewDictionary();
                    int reconnectedBefore = RefCount(reconnected);
                    context.SwitchTo(reconnected);
                    probe.CallStatic(CoreServices, "ResetExcel");
                    Assert.AreEqual(0, (int)probe.InvokeOnKept("Count"), "a reference taken before the reset still works");
                    Assert.AreEqual("com", probe.DescribeExcel(), "the core asks the shell again");
                    probe.InvokeOnExcel("Add", "k", "v");
                    Assert.AreEqual(1, (int)reconnected.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, reconnected, null),
                        "the core now works on the object the shell hands out");
                    probe.CallStatic(CoreServices, "Shutdown");
                    Assert.AreEqual(before, RefCount(dictionary), "the dropped RCW is released at Shutdown too");
                    Assert.AreEqual(reconnectedBefore, RefCount(reconnected), "and the current one");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void TodaysShutdownSequence_DoesNotDoubleRelease()
        {
            using (var t = new TempFolder())
            {
                object dictionary = NewDictionary();
                int before = RefCount(dictionary);
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), dictionary);
                var (domain, probe) = CoreProbe.Create("CoreExcelTodaysShutdown");
                try
                {
                    probe.CallStatic(CoreServices, "Initialize", context);
                    probe.RunTodaysShutdownSequence();
                    Assert.AreEqual(before, RefCount(dictionary));
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void OnlyRealDisconnects_CountAsADeadExcel()
        {
            // Excel rejects calls while a cell is edited or a dialog is open; that is not a dead Excel.
            var (domain, probe) = CoreProbe.Create("CoreExcelDisconnect");
            try
            {
                bool Dead(Exception ex) => (bool)probe.CallStatic(CoreServices, "IsDisconnected", ex);
                Assert.IsTrue(Dead(new COMException("disconnected", unchecked((int)0x80010108))), "RPC_E_DISCONNECTED");
                Assert.IsTrue(Dead(new COMException("server unavailable", unchecked((int)0x800706BA))), "RPC_S_SERVER_UNAVAILABLE");
                Assert.IsTrue(Dead(new COMException("not connected", unchecked((int)0x800401FD))), "CO_E_OBJNOTCONNECTED");
                Assert.IsTrue(Dead(new InvalidComObjectException("separated")));
                Assert.IsFalse(Dead(new COMException("rejected", unchecked((int)0x80010001))), "RPC_E_CALL_REJECTED (busy)");
                Assert.IsFalse(Dead(new COMException("retry later", unchecked((int)0x8001010A))), "RPC_E_SERVERCALL_RETRYLATER");
                Assert.IsFalse(Dead(new InvalidOperationException("other")));
            }
            finally
            {
                AppDomain.Unload(domain);
            }
        }

        [TestMethod]
        public void CoreShutdown_ReleasesExcel()
        {
            // AddinEntry.Shutdown is what the shell calls before every unload and when Excel closes.
            using (var t = new TempFolder())
            {
                object dictionary = NewDictionary();
                int before = RefCount(dictionary);
                var context = new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), dictionary);
                var (domain, probe) = CoreProbe.Create("CoreEntryRelease");
                try
                {
                    probe.InitializeUseAndShutDownTheEntry(context);
                    Assert.AreEqual(before, RefCount(dictionary), "AddinEntry.Shutdown releases the core's Excel");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }

        [TestMethod]
        public void NoExcel_OrNoContext_IsNullNotAnError()
        {
            using (var t = new TempFolder())
            {
                var (domain, probe) = CoreProbe.Create("CoreExcelNone");
                try
                {
                    Assert.AreEqual("null", probe.DescribeExcel(), "before Initialize");
                    probe.CallStatic(CoreServices, "Initialize", new ComContext(new PathProvider(t.Combine("inst"), t.Combine("data")), null));
                    Assert.AreEqual("null", probe.DescribeExcel(), "a context without Excel");
                    probe.CallStatic(CoreServices, "Shutdown");
                    probe.CallStatic(CoreServices, "Shutdown");
                }
                finally
                {
                    AppDomain.Unload(domain);
                }
            }
        }
    }
}
