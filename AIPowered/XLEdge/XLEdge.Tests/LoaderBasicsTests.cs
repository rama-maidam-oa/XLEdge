using System;
using System.IO;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class LoaderBasicsTests
    {
        [TestMethod]
        public void ContractCompatibility_Rules()
        {
            Assert.IsFalse(ContractCompatibility.IsLoadable(0));
            Assert.AreEqual("built without a contract version", ContractCompatibility.RefusalReason(0));

            Assert.IsTrue(ContractCompatibility.IsLoadable(1));
            Assert.IsNull(ContractCompatibility.RefusalReason(1));
            Assert.IsTrue(ContractCompatibility.IsLoadable(ContractInfo.Version));

            int newer = ContractInfo.Version + 1;
            Assert.IsFalse(ContractCompatibility.IsLoadable(newer));
            Assert.AreEqual($"needs a newer XLEdge installer (contract {newer}, this installer supports up to {ContractInfo.Version})", ContractCompatibility.RefusalReason(newer));
            Assert.IsFalse(ContractCompatibility.IsLoadable(-1));
        }

        [TestMethod]
        public void ResolvedRelease_Fields()
        {
            var r = new ResolvedRelease { Version = "11.1.2.0", ReleaseDate = "d", FolderName = "v1", ContractVersion = 1 };
            Assert.AreEqual("v1", r.FolderName);
            Assert.AreEqual(1, r.ContractVersion);
        }

        [TestMethod]
        public void AddinEntry_IsMarshalByRefWithInfiniteLease()
        {
            var entry = CreateEntry();
            Assert.IsInstanceOfType(entry, typeof(MarshalByRefObject));
            Assert.IsNull(((MarshalByRefObject)entry).InitializeLifetimeService());
        }

        [TestMethod]
        public void AddinEntry_InitializeLogsDomainAndQueriesAreSafe()
        {
            using (var t = new TempFolder())
            {
                var logger = new RecordingLogger();
                var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data"))) { Logger = logger };
                var entry = CreateEntry();

                entry.Initialize(context);
                Assert.IsTrue(logger.Contains($"AddinEntry.Initialize: XLEdge core initialized in AppDomain '{AppDomain.CurrentDomain.FriendlyName}'"), logger.Dump());
                Assert.IsNotNull(entry.GetLoginInfo());
                Assert.IsFalse(entry.GetLoginInfo().IsLoggedIn);
                Assert.AreEqual(0, entry.GetOpenWindowTitles().Length);
                Assert.IsFalse(entry.IsBusy());
                Assert.AreEqual(IntPtr.Zero, entry.CreateTaskPaneContent(new IntPtr(1)));
                entry.OnRibbonAction("RibEdgeRefresh", null);
                Assert.IsTrue(entry.OnExcelEvent("SheetActivate", new object[] { "Sheet1" }));
                entry.Shutdown();
                Assert.IsTrue(logger.Contains("AddinEntry.Shutdown: XLEdge core shut down."), logger.Dump());
            }
        }

        private static IXLEdgeAddin CreateEntry()
        {
            var asm = Assembly.LoadFrom(Path.Combine(ReleaseFolders.CoreBuildFolder, "XLEdge.Addin.Core.dll"));
            return (IXLEdgeAddin)Activator.CreateInstance(asm.GetType("XLEdge.Addin.Core.AddinEntry", throwOnError: true));
        }
    }
}
