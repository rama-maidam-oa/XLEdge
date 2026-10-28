using System;
using System.IO;
using System.Reflection;
using System.Runtime.Remoting;
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

        // The core is created in a throwaway AppDomain so it never gets loaded into the test domain
        // (other tests assert that the shell's domain never loads it).
        private static void WithCoreEntry(Action<IXLEdgeAddin, string> test)
        {
            string coreFolder = ReleaseFolders.CoreBuildFolder;
            var domain = AppDomain.CreateDomain("CoreEntryTest", null, new AppDomainSetup { ApplicationBase = coreFolder });
            try
            {
                var entry = (IXLEdgeAddin)domain.CreateInstanceFromAndUnwrap(Path.Combine(coreFolder, "XLEdge.Addin.Core.dll"), "XLEdge.Addin.Core.AddinEntry");
                test(entry, domain.FriendlyName);
            }
            finally
            {
                AppDomain.Unload(domain);
            }
        }

        [TestMethod]
        public void AddinEntry_IsMarshalByRefWithInfiniteLease()
        {
            WithCoreEntry((entry, _) =>
            {
                Assert.IsTrue(RemotingServices.IsTransparentProxy(entry));
                Assert.IsNull(RemotingServices.GetLifetimeService((MarshalByRefObject)entry));
            });
        }

        [TestMethod]
        public void AddinEntry_InitializeLogsDomainAndQueriesAreSafe()
        {
            using (var t = new TempFolder())
            {
                var logger = new RecordingLogger();
                var context = new FakeContext(new PathProvider(t.Combine("inst"), t.Combine("data"))) { Logger = logger };
                WithCoreEntry((entry, domainName) =>
                {
                    entry.Initialize(context);
                    Assert.IsTrue(logger.Contains($"AddinEntry.Initialize: XLEdge core initialized in AppDomain '{domainName}'"), logger.Dump());
                    Assert.IsFalse(entry.GetLoginInfo().IsLoggedIn);
                    Assert.AreEqual(0, entry.GetOpenWindowTitles().Length);
                    Assert.IsFalse(entry.IsBusy());
                    Assert.AreEqual(IntPtr.Zero, entry.CreateTaskPaneContent(new IntPtr(1)));
                    entry.OnRibbonAction("RibEdgeRefresh", null);
                    Assert.IsTrue(entry.OnExcelEvent("SheetActivate", new object[] { "Sheet1" }));
                    entry.Shutdown();
                    Assert.IsTrue(logger.Contains("AddinEntry.Shutdown: XLEdge core shut down."), logger.Dump());
                });
            }
        }

        // ---- exceptions crossing back to the shell -----------------------------------

        private sealed class CoreOnlyException : Exception
        {
            public CoreOnlyException(string message) : base(message) { }
        }

        [TestMethod]
        public void CrossDomain_FrameworkExceptionsPassThrough_OthersAreWrapped()
        {
            Assert.IsFalse(CrossDomain.NeedsWrapping(new FileNotFoundException("x")));
            Assert.IsFalse(CrossDomain.NeedsWrapping(new TargetInvocationException(new InvalidOperationException("x"))));

            var custom = new CoreOnlyException("boom");
            Assert.IsTrue(CrossDomain.NeedsWrapping(custom));
            Assert.IsTrue(CrossDomain.NeedsWrapping(new TargetInvocationException(custom)));
            Assert.IsTrue(CrossDomain.NeedsWrapping(new AggregateException(new InvalidOperationException("a"), custom)));

            var wrapped = CrossDomain.Wrap(new TargetInvocationException(custom));
            Assert.IsInstanceOfType(wrapped, typeof(InvalidOperationException));
            StringAssert.Contains(wrapped.Message, "CoreOnlyException: boom");
            Assert.IsNull(wrapped.InnerException, "nothing core-specific may be serialized back");
        }
    }
}
