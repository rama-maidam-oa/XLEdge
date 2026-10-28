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
            foreach (string text in File.ReadAllLines(designerFile))
            {
                var m = line.Match(text);
                if (!m.Success) continue;
                string field = m.Groups["field"].Value;
                if (!result.TryGetValue(field, out var set)) result[field] = set = new SortedSet<string>(StringComparer.Ordinal);
                set.Add(m.Groups["rest"].Value.Trim());
            }
            return result;
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
            CollectionAssert.IsSubsetOf(new[] { "Caption = \"Reload Add-in\"", "ImageList = this.ImageList_16X16", "Image = 0", "ScreenTip = \"Reload Add-in\"" }, reload.ToArray(),
                string.Join(" | ", reload));
            var history = shell["RibVersionHistory"];
            CollectionAssert.IsSubsetOf(new[] { "Caption = \"Version History\"", "ImageList = this.ImageList_32X32", "Image = 11",
                "Size = AddinExpress.MSO.ADXRibbonXControlSize.Large", "ScreenTip = \"Version History\"" }, history.ToArray(), string.Join(" | ", history));

            var ribbonFields = shell.Keys.Where(k => k.StartsWith("Rib", StringComparison.Ordinal) || k.StartsWith("adxRibbon", StringComparison.Ordinal) || k == "TabXLEdge");
            CollectionAssert.AreEquivalent(TodaysControls.Concat(new[] { "RibReload", "RibVersionHistory" }).ToArray(), ribbonFields.ToArray());
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
