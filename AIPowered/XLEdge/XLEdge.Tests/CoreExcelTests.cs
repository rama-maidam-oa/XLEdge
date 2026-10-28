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
            private readonly object _com;
            public ComContext(IPathProvider paths, object com) { Paths = paths; _com = com; Logger = new RecordingLogger(); }
            public override object InitializeLifetimeService() => null;
            public object ExcelApp => _com;
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
