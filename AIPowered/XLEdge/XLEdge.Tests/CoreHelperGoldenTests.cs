using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// The ported helpers behave exactly as XLEdge 11.1.2's: the same inputs (GoldenCases) give the outputs 11.1.2's own build
    /// gave (GoldenValues), evaluated inside a loader-like AppDomain over the core.
    /// </summary>
    [TestClass]
    public class CoreHelperGoldenTests
    {
        [TestMethod]
        public void PortedHelpers_GiveXLEdge11_1_2sResults()
        {
            var (domain, probe) = CoreProbe.Create("CoreGolden");
            string[] actual;
            try { actual = probe.RunGoldenCases(); }
            finally { AppDomain.Unload(domain); }

            string[] expected = GoldenValues.Lines;
            Assert.AreEqual(expected.Length, actual.Length, "the case list and the golden table are out of step: regenerate GoldenValues");
            var differences = new List<string>();
            for (int i = 0; i < expected.Length; i++)
                if (expected[i] != actual[i]) differences.Add($"11.1.2: {expected[i]}{Environment.NewLine}   core: {actual[i]}");
            Assert.AreEqual(0, differences.Count, string.Join(Environment.NewLine, differences.Take(15)));
        }

        [TestMethod]
        public void GoldenTable_CoversEveryCultureAndTheBranchesThatMatter()
        {
            string[] lines = GoldenValues.Lines;
            foreach (string culture in new[] { "en-US ", "de-DE ", "th-TH " })
                Assert.IsTrue(lines.Count(l => l.StartsWith(culture, StringComparison.Ordinal)) > 50, culture);
            // Empty attachment anchors (OISR-22767) and HTML error pages are among the cases.
            Assert.IsTrue(lines.Any(l => l.StartsWith("AttachmentLinkHelper.IsEmptyAttachmentAnchor(\"<a href=''></a>\") = b:True", StringComparison.Ordinal)),
                "OISR-22767 empty anchor case");
            Assert.IsTrue(lines.Any(l => l.StartsWith("ApiErrorMessageExtractor.IsHtmlResponse(\"<!DOCTYPE html>", StringComparison.Ordinal) && l.EndsWith("b:True")));
            // The request writer (NumericJsonConverter.WriteDictionary: ORACLE_* always strings, numeric-looking strings).
            Assert.IsTrue(lines.Any(l => l.StartsWith("en-US JSONW.RequestBody() = s:{\"ORACLE_RESP_ID\":\"50559\"", StringComparison.Ordinal)), "request body case");
            // Real EBS/Fusion attachment cells and download URLs.
            Assert.IsTrue(lines.Any(l => l.Contains("linkValue=s:ATTACHMENT|businessObjectId=34|documentId=12|get-fusion-file")), "attachment link parse");
            Assert.IsTrue(lines.Any(l => l.EndsWith("= s:https://host/web/secure/get-fusion-file?businessObjectId=34&documentId=12")), "download URL");
            // The numeric drilldown branches (int, decimal, BigInteger).
            foreach (string kind in new[] { "= i:42", "= m:12.50", "= bi:123456789012345678901234" })
                Assert.IsTrue(lines.Any(l => l.StartsWith("en-US XLEdgeValueFormatter.FormatDrilldownValue(", StringComparison.Ordinal) && l.EndsWith(kind)), kind);
        }
    }
}
