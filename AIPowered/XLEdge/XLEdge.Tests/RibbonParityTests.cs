using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// The new shell's ribbon must be today's XLEdge ribbon (same field names, Ids, captions, images, sizes, tips, visibility),
    /// plus only the two hot-reload buttons. Compares the designer code of both add-ins.
    /// </summary>
    [TestClass]
    public class RibbonParityTests
    {
        private static readonly string[] TodaysControls =
        {
            "TabXLEdge",
            "adxRibbonGroup1", "adxRibbonGroup2", "adxRibbonGroup3", "adxRibbonGroup4", "adxRibbonGroup5", "adxRibbonGroup6",
            "adxRibbonGroup7", "adxRibbonGroup8", "adxRibbonGroup9",
            "RibEdgeLogin", "RibEdgeLogout", "RibLoginURL", "RibEdgeDialogBoxLauncher", "RibEdgeRefresh", "RibEdgeRefreshAll",
            "RibEdgeParamRefresh", "RibEdgeParamRefreshBook", "RibEdgeShowHide", "RibEdgeDebug", "RibEdgeIncludeOutputData",
            "RibEdgeOptions", "RibControlSheet", "RibEdgeAbout", "RibEdgeHelp", "RibSheetLabel",
        };

        private static string TestBin => Path.GetDirectoryName(typeof(RibbonParityTests).Assembly.Location);
        private static string ShellDesigner => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\XLEdge\AddinModule.Designer.cs"));
        private static string TodaysDesigner => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\..\..\XLEdge\AddinModule.Designer.cs"));

        /// <summary>"this.Field.Property = value;" assignments per field (event subscriptions and Controls.Add included).</summary>
        private static Dictionary<string, SortedSet<string>> Assignments(string designerFile)
        {
            var result = new Dictionary<string, SortedSet<string>>(StringComparer.Ordinal);
            var line = new Regex(@"^\s*this\.(?<field>\w+)\.(?<rest>.+?);\s*$");
            var creation = new Regex(@"^\s*this\.(?<field>\w+)\s*=\s*new\s+(?<type>[\w.]+)\(");
            foreach (string text in JoinContinuations(File.ReadAllLines(designerFile)))
            {
                var c = creation.Match(text);
                var m = line.Match(text);
                string field = c.Success ? c.Groups["field"].Value : m.Success ? m.Groups["field"].Value : null;
                if (field == null) continue;
                if (!result.TryGetValue(field, out var set)) result[field] = set = new SortedSet<string>(StringComparer.Ordinal);
                set.Add(c.Success ? "new " + c.Groups["type"].Value : m.Groups["rest"].Value.Trim());
            }
            return result;
        }

        /// <summary>Designer code splits long strings over lines ending in '+'; join them into one statement.</summary>
        private static IEnumerable<string> JoinContinuations(IEnumerable<string> lines)
        {
            string pending = null;
            foreach (string raw in lines)
            {
                string text = pending == null ? raw : pending + " " + raw.Trim();
                if (text.TrimEnd().EndsWith("+")) { pending = text.TrimEnd(); continue; }
                pending = null;
                yield return text;
            }
        }

        [TestMethod]
        public void TodaysRibbon_IsUnchangedInTheNewShell()
        {
            Assert.IsTrue(File.Exists(TodaysDesigner), TodaysDesigner);
            Assert.IsTrue(File.Exists(ShellDesigner), ShellDesigner);
            var today = Assignments(TodaysDesigner);
            var shell = Assignments(ShellDesigner);

            var differences = new List<string>();
            foreach (string field in TodaysControls)
            {
                Assert.IsTrue(today.ContainsKey(field), "today's designer has no " + field);
                if (!shell.TryGetValue(field, out var shellSet)) { differences.Add(field + ": missing"); continue; }
                var todaySet = today[field];
                // The two hot-reload buttons are added to groups 6 and 7.
                var extra = shellSet.Except(todaySet).Where(a => !(field == "adxRibbonGroup6" && a == "Controls.Add(this.RibReload)")
                                                               && !(field == "adxRibbonGroup7" && a == "Controls.Add(this.RibVersionHistory)")).ToList();
                var missing = todaySet.Except(shellSet).ToList();
                if (extra.Count > 0 || missing.Count > 0)
                    differences.Add($"{field}: missing [{string.Join(" | ", missing)}] extra [{string.Join(" | ", extra)}]");
            }
            Assert.AreEqual(0, differences.Count, string.Join(Environment.NewLine, differences));
        }

        [TestMethod]
        public void ShellRibbon_AddsOnlyTheHotReloadButtons()
        {
            var shell = Assignments(ShellDesigner);
            var reload = shell["RibReload"];
            CollectionAssert.IsSubsetOf(new[] { "Caption = \"Reload Add-in\"", "ImageList = this.ImageList_HotReload16", "Image = 40", "ScreenTip = \"Reload Add-in\"" }, reload.ToArray(),
                string.Join(" | ", reload));
            var history = shell["RibVersionHistory"];
            CollectionAssert.IsSubsetOf(new[] { "Caption = \"Version History\"", "ImageList = this.ImageList_HotReload32", "Image = 34",
                "Size = AddinExpress.MSO.ADXRibbonXControlSize.Large", "ScreenTip = \"Version History\"" }, history.ToArray(), string.Join(" | ", history));

            var ribbonFields = shell.Keys.Where(k => k.StartsWith("Rib", StringComparison.Ordinal) || k.StartsWith("adxRibbon", StringComparison.Ordinal) || k == "TabXLEdge");
            CollectionAssert.AreEquivalent(TodaysControls.Concat(new[] { "RibReload", "RibVersionHistory" }).ToArray(), ribbonFields.ToArray());
        }

        private static System.Xml.Linq.XElement ResxValue(string resxFile, string name) =>
            System.Xml.Linq.XDocument.Load(resxFile).Root.Elements("data").Single(d => (string)d.Attribute("name") == name).Element("value");

        [TestMethod]
        public void ShellResx_KeepsTodaysImagesAndTipsByteForByte()
        {
            // Re-encoding an ImageList changes the colours of semi-transparent pixels; today's icons must stay as they are.
            string todayResx = Path.Combine(Path.GetDirectoryName(TodaysDesigner), "AddinModule.resx");
            string shellResx = Path.Combine(Path.GetDirectoryName(ShellDesigner), "AddinModule.resx");
            foreach (string key in new[] { "ImageList_32X32.ImageStream", "RibEdgeRefresh.SuperTip", "RibEdgeRefreshAll.SuperTip", "RibControlSheet.SuperTip" })
                Assert.AreEqual(ResxValue(todayResx, key).Value.Trim(), ResxValue(shellResx, key).Value.Trim(), key);
        }

        [TestMethod]
        public void ShellModule_KeepsTodaysIdentity()
        {
            string module = File.ReadAllText(Path.Combine(Path.GetDirectoryName(ShellDesigner), "AddinModule.cs"));
            StringAssert.Contains(module, "[Guid(\"80B0FB76-A5F4-41E0-B283-C175B7374B60\"), ProgId(\"XLEdge.AddinModule\")]");
            string designer = File.ReadAllText(ShellDesigner);
            StringAssert.Contains(designer, "this.AddinName = \"XLEdge\";");
            StringAssert.Contains(designer, "this.SupportedApps = AddinExpress.MSO.ADXOfficeHostApp.ohaExcel;");
            StringAssert.Contains(module, "public void InvokedFromGLSense(string eeName, string eeUrl, string eeToken, string eeUser, bool hasXLEdgePermission = true)");
            StringAssert.Contains(module, "public void LogoffFromAddin()");
        }
    }
}
