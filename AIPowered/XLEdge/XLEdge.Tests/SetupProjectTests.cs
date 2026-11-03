using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 5: the twin's setup project (XLEdge\OrbitXLEdgeSetups\OrbitXLEdge.vdproj) - an MSI 11.1.3 that upgrades today's
    /// XLEdge 11.1.2 (same UpgradeCode), installs the shell + signed loaders + the first core, and removes AddinCore\ on a real
    /// uninstall. Source checks; the MSI's real contents are checked with msiexec /a (ledger), the install by the developer.
    /// </summary>
    [TestClass]
    public class SetupProjectTests
    {
        private static string TwinRoot => Path.GetFullPath(Path.Combine(Path.GetDirectoryName(typeof(SetupProjectTests).Assembly.Location), @"..\..\.."));
        private static string Vdproj => Path.Combine(TwinRoot, @"XLEdge\OrbitXLEdgeSetups\OrbitXLEdge.vdproj");
        private static string Today => Path.GetFullPath(Path.Combine(TwinRoot, @"..\..\XLEdge\setupfiles\OrbitXLEdge.vdproj"));

        private static string Value(string text, string name) =>
            Regex.Match(text, "\"" + name + "\" = \"8:([^\"]*)\"").Groups[1].Value;

        /// <summary>The product's code (a GUID - the bootstrapper's prerequisite entries also have a "ProductCode").</summary>
        private static string Code(string text, string name) =>
            Regex.Match(text, "\"" + name + "\" = \"8:(\\{[0-9A-F-]+\\})\"").Groups[1].Value;

        [TestMethod]
        public void SetupProject_UpgradesTodaysXLEdge()
        {
            string twin = File.ReadAllText(Vdproj), today = File.ReadAllText(Today);
            Assert.AreEqual("{D090F1E0-CA8E-4529-B78D-981652FFF708}", Code(twin, "UpgradeCode"), "11.1.2's UpgradeCode: a major upgrade replaces it");
            Assert.AreEqual(Code(today, "UpgradeCode"), Code(twin, "UpgradeCode"));
            Assert.AreEqual("11.1.3", Value(twin, "ProductVersion"));
            foreach (string code in new[] { "ProductCode", "PackageCode" })
            {
                Assert.AreNotEqual("", Code(twin, code), code);
                Assert.AreNotEqual(Code(today, code), Code(twin, code), code + ": a new product");
            }
            foreach (string same in new[] { "ProductName", "Manufacturer", "Title" }) Assert.AreEqual(Value(today, same), Value(twin, same), same);
            StringAssert.Contains(twin, "\"RemovePreviousVersions\" = \"11:TRUE\"");
            StringAssert.Contains(twin, "\"InstallAllUsers\" = \"11:FALSE\"");
            StringAssert.Contains(twin, "\"DefaultLocation\" = \"8:[LocalAppDataFolder][Manufacturer]\\\\[ProductName]\"");
        }

        [TestMethod]
        public void SetupProject_InstallsTheTwinShell_SignedLoaders_AndTheFirstCore()
        {
            string twin = File.ReadAllText(Vdproj);
            StringAssert.Contains(twin, "\"OutputProjectGuid\" = \"8:{98093E3C-A05D-4EE1-BF6D-AC742BE9F782}\"", "the twin shell's Primary Output");
            Assert.IsFalse(twin.Contains("{D090F1E0-CA8E-4529-B78D-981652FFF707}"), "no output of 11.1.2's project");
            Assert.IsFalse(twin.Contains("obj\\\\Debug"), "no stale obj\\Debug path");
            StringAssert.Contains(twin, "\"SourcePath\" = \"8:..\\\\bin\\\\Release\\\\adxloader.dll\"");
            StringAssert.Contains(twin, "\"SourcePath\" = \"8:..\\\\bin\\\\Release\\\\adxloader64.dll\"");
            StringAssert.Contains(twin, "XLEdge.Addin.Core\\\\SetupFiles\\\\Release\\\\InstallerPayload\\\\OrbitXLEdge.zip");
            StringAssert.Contains(twin, "XLEdge.Addin.Core\\\\SetupFiles\\\\Release\\\\Manifest\\\\manifest.json");
            StringAssert.Contains(twin, "\"Name\" = \"8:AddinCore\"");
            StringAssert.Contains(twin, "\"Name\" = \"8:Manifest\"");
            StringAssert.Contains(twin, "\"Arguments\" = \"8:/install=\\\"XLEdge.dll\\\" /privileges=user\"", "registered through adxregistrator, as today");
        }

        /// <summary>The dependency entry (detected by VS) whose source is <paramref name="sourceName"/> and whose display name contains <paramref name="asm"/>.</summary>
        private static string DependencyEntry(string text, string sourceName, string asm)
        {
            foreach (Match m in Regex.Matches(text, "            \"\\{[0-9A-F-]+\\}:_[0-9A-F]{32}\"\n            \\{\n(.*?)\n            \\}\n", RegexOptions.Singleline))
                if (m.Value.Contains("\"SourcePath\" = \"8:" + sourceName + "\"") && m.Value.Contains(asm)) return m.Value;
            return null;
        }

        [TestMethod]
        public void SetupProject_ExcludesWhatMustNotShip_AndNoCoreOnlyDependency()
        {
            string twin = File.ReadAllText(Vdproj).Replace("\r\n", "\n");
            string tlb = DependencyEntry(twin, "AddinExpress.MSO.2005.tlb", "");
            Assert.IsNotNull(tlb, "11.1.2's tlb entry kept");
            StringAssert.Contains(tlb, "\"Exclude\" = \"11:TRUE\"", "the tlb stays excluded, as in 11.1.2");
            string gacCompression = DependencyEntry(twin, "System.IO.Compression.dll", "Version=4.0.0.0");
            Assert.IsNotNull(gacCompression, "the GAC System.IO.Compression (Windows File Protection) has an entry");
            StringAssert.Contains(gacCompression, "\"Exclude\" = \"11:TRUE\"", "excluded, as GLSense 11.1.3");
            foreach (string coreOnly in new[] { "MahApps.Metro", "Microsoft.Web.WebView2", "Wpf.Ui", "Microsoft.Xaml.Behaviors", "XLEdge.Addin.Core" })
                Assert.IsFalse(twin.Contains("\"SourcePath\" = \"8:" + coreOnly), coreOnly + " ships in the core zip, not next to the shell");
        }

        [TestMethod]
        public void SetupProject_RemovesAddinCore_OnARealUninstallOnly()
        {
            string twin = File.ReadAllText(Vdproj);
            StringAssert.Contains(twin, "\"TargetName\" = \"8:XLEdgeUninstallCleanup.exe\"");
            StringAssert.Contains(twin, "\"Condition\" = \"8:REMOVE=\\\"ALL\\\" AND NOT UPGRADINGPRODUCTCODE\"");
            StringAssert.Contains(twin, "rd /s /q \\\"[TARGETDIR]AddinCore\\\"");
        }

        [TestMethod]
        public void SetupProject_SignsTheShellAndLoaders_BeforeItBuilds()
        {
            // The whole line: the value carries escaped quotes (\").
            string pre = File.ReadAllLines(Vdproj).Single(l => l.Contains("\"PreBuildEvent\""));
            foreach (string file in new[] { "XLEdge.dll", "adxloader.dll", "adxloader64.dll" })
                StringAssert.Contains(pre, "XLEdge.Build\\\\sign.ps1\\\" -File \\\"$(ProjectDir)..\\\\bin\\\\Release\\\\" + file, file);
            StringAssert.Contains(pre, "CleanOnlineAppCache");
        }

        [TestMethod]
        public void SetupSolution_BuildsTheSetupProjectInReleaseOnly()
        {
            string sln = File.ReadAllText(Path.Combine(TwinRoot, "OrbitXLEdgeSetup.sln"));
            Match project = Regex.Match(sln, "Project\\(\"\\{54435603-DBB4-11D2-8724-00A0C9A8B90C\\}\"\\) = \"OrbitXLEdge\", \"XLEdge\\\\OrbitXLEdgeSetups\\\\OrbitXLEdge.vdproj\", \"(\\{[0-9A-F-]+\\})\"");
            Assert.IsTrue(project.Success, "the setup project is in OrbitXLEdgeSetup.sln");
            string guid = project.Groups[1].Value;
            StringAssert.Contains(sln, guid + ".Release|Any CPU.Build.0 = Release");
            Assert.IsFalse(sln.Contains(guid + ".Debug|Any CPU.Build.0"), "Debug builds do not make an MSI");
            Assert.IsFalse(File.ReadAllText(Path.Combine(TwinRoot, "XLEdge.sln")).Contains(".vdproj"), "XLEdge.sln stays MSBuild-only");
        }
    }
}
