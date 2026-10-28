using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    [TestClass]
    public class VersionParserTests
    {
        private const string Manifest =
            "[{\"version\":\"11.1.2\",\"releaseDate\":\"2026-10-20T09:00:00\",\"folderName\":\"v11.1.2_a\",\"contractVersion\":1}," +
            " {\"version\":\"11.1.10\",\"releaseDate\":\"2026-10-28T14:05:09\",\"folderName\":\"v11.1.10_b\",\"fileName\":\"v11.1.10_b.zip\",\"checksum\":\"sha-b\",\"contractVersion\":2}," +
            " {\"version\":\"11.1.10\",\"releaseDate\":\"2026-10-29T00:00:00\",\"folderName\":\"v11.1.10_tie\",\"contractVersion\":3}]";

        [TestMethod]
        public void Parse_Array_TakesLatest_WithContractVersion()
        {
            var result = new VersionParser().ParseVersionJson(Manifest);
            Assert.IsTrue(result.Success, result.ErrorMessage);
            Assert.AreEqual("11.1.10", result.Version);
            Assert.AreEqual(new DateTime(2026, 10, 28, 14, 5, 9), result.ReleaseDate);
            Assert.AreEqual(2, result.ContractVersion, "ties go to the first entry");
            Assert.AreEqual("v11.1.10_b", result.FolderName);
            Assert.AreEqual("sha-b", result.Checksum);
            Assert.AreEqual("v11.1.10_b.zip", result.FileName);
            Assert.AreEqual(3, result.AllVersions.Count);
            Assert.AreEqual("v11.1.2_a", result.AllVersions[0].FolderName);
        }

        [TestMethod]
        public void Parse_Errors()
        {
            var parser = new VersionParser();
            AssertError(parser.ParseVersionJson(null), "JSON content is empty or null");
            AssertError(parser.ParseVersionJson("  "), "JSON content is empty or null");
            AssertError(parser.ParseVersionJson("[]"), "No version data found in JSON");
            AssertError(parser.ParseVersionJson("null"), "No version data found in JSON");
            AssertError(parser.ParseVersionJson("[{\"version\":\"x\"}]"), "Failed to determine latest version");
            StringAssert.StartsWith(parser.ParseVersionJson("{\"version\":\"11.1.2\"}").ErrorMessage, "Invalid JSON format: ");
            StringAssert.StartsWith(parser.ParseVersionJson("garbage").ErrorMessage, "Invalid JSON format: ");
        }

        [TestMethod]
        public void Parse_ReleaseDate_InvariantUnderDeDE()
        {
            var previous = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("de-DE");
            try
            {
                var result = new VersionParser().ParseVersionJson("[{\"version\":\"11.1.2\",\"releaseDate\":\"2026-10-28T14:05:09\"}]");
                Assert.AreEqual(new DateTime(2026, 10, 28, 14, 5, 9), result.ReleaseDate);
                Assert.AreEqual(new DateTime(2026, 10, 28), new VersionParser().ParseVersionJson("[{\"version\":\"1.0\",\"releaseDate\":\"28-Oct-2026\"}]").ReleaseDate);
                Assert.AreEqual(DateTime.MinValue, new VersionParser().ParseVersionJson("[{\"version\":\"1.0\",\"releaseDate\":\"28.10.2026\"}]").ReleaseDate);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [TestMethod]
        public void ParseFile_MissingAndPresent()
        {
            using (var t = new TempFolder())
            {
                string path = t.Combine("manifest.json");
                AssertError(new VersionParser().ParseVersionFile(path), "Manifest file not found: " + path);
                File.WriteAllText(path, Manifest);
                Assert.AreEqual("11.1.10", new VersionParser().ParseVersionFile(path).Version);
            }
        }

        [TestMethod]
        public void GetLatestVersion_SkipsUnparseable()
        {
            var latest = new VersionParser().GetLatestVersion(new[]
            {
                new VersionInfo { Version = "11.1.9" }, new VersionInfo { Version = null }, new VersionInfo { Version = "v12" },
                new VersionInfo { Version = "11.1.11" },
            });
            Assert.AreEqual("11.1.11", latest.Version);
            Assert.IsNull(new VersionParser().GetLatestVersion(new VersionInfo[0]));
            Assert.IsNull(new VersionParser().GetLatestVersion(null));
        }

        private static void AssertError(VersionParseResult result, string message)
        {
            Assert.IsFalse(result.Success);
            Assert.AreEqual(message, result.ErrorMessage);
        }
    }
}
