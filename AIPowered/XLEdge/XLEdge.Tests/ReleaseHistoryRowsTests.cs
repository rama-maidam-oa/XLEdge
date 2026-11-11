using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;
using XLEdge.Views;

namespace XLEdge.Tests
{
    [TestClass]
    public class ReleaseHistoryRowsTests
    {
        private static ReleaseEntry E(string folder, string version, string date, string source = "Offline") =>
            new ReleaseEntry { FolderName = folder, Version = version, ReleaseDate = date, Source = source, Notes = "n-" + folder, ContractVersion = 1 };

        [TestMethod]
        public void Build_OrdersNewestFirstAndMarksTheLoadedFolder()
        {
            var context = new FakeContext(new PathProvider(@"C:\i", @"C:\d")) { ActiveFolderName = "vB" };
            var rows = ReleaseHistoryRows.Build(new[]
            {
                E("vA", "11.1.1.0", "2026-10-01T10:00:00"),
                E("vC", "11.1.3.0", "2026-10-20T10:00:00", "Online"),
                E("vB", "11.1.2.0", "2026-10-10T10:00:00"),
                E("vX", "9.9", "garbage"),
            }, context);

            CollectionAssert.AreEqual(new[] { "vC", "vB", "vA", "vX" }, rows.Select(r => r.FolderName).ToArray());
            CollectionAssert.AreEqual(new[] { false, true, false, false }, rows.Select(r => r.IsCurrentlyLoaded).ToArray());
            Assert.AreEqual("Online", rows[0].Source);
        }

        [TestMethod]
        public void Build_FallsBackToVersionAndDateWhenNoFolderIsActive()
        {
            var context = new FakeContext(new PathProvider(@"C:\i", @"C:\d")) { Version = "11.1.2.0", ReleaseDate = " 2026-10-10T10:00:00 " };
            var rows = ReleaseHistoryRows.Build(new[] { E("vA", "11.1.1.0", "2026-10-01T10:00:00"), E("vB", "11.1.2.0", "2026-10-10T10:00:00") }, context);
            Assert.IsTrue(rows.Single(r => r.FolderName == "vB").IsCurrentlyLoaded);
            Assert.IsFalse(rows.Single(r => r.FolderName == "vA").IsCurrentlyLoaded);
        }

        [TestMethod]
        public void Build_NothingLoaded_NoRowRunning()
        {
            var context = new FakeContext(new PathProvider(@"C:\i", @"C:\d"));
            var rows = ReleaseHistoryRows.Build(new[] { E("vA", "11.1.1.0", "2026-10-01T10:00:00") }, context);
            Assert.IsFalse(rows.Single().IsCurrentlyLoaded);
        }
    }
}
