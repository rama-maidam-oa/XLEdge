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
        public void SetupProject_SignsTheCopyItPackages_AndRefusesUnsignedFiles()
        {
            // Review fix: the shell's Primary Output packages obj\Release\XLEdge.dll (MSBuild's IntermediateAssembly), not the
            // bin copy - verified by making them differ (ledger). Both are signed, then every shipped file of ours is checked.
            string pre = File.ReadAllLines(Vdproj).Single(l => l.Contains("\"PreBuildEvent\""));
            StringAssert.Contains(pre, "-File \\\"$(ProjectDir)..\\\\obj\\\\Release\\\\XLEdge.dll\\\"", "the packaged copy is signed");
            int verify = pre.IndexOf("verify-signed.ps1", StringComparison.Ordinal);
            Assert.IsTrue(verify > pre.LastIndexOf("sign.ps1\\\"", StringComparison.Ordinal), "the signature check runs after the signing");
        }

        [TestMethod]
        public void SetupProject_KeepsTheKeysUpgradesDependOn()
        {
            // A file entry's ComponentId follows its key: a regenerated key makes the next MSI's late removal of this one delete
            // the file the new version just installed (11.1.2's keys are kept for the same reason).
            string twin = File.ReadAllText(Vdproj);
            foreach (var (key, what) in new[]
            {
                ("_7F3C2A91B6D84E0F9A5B1C8D2E6F4A73", "OrbitXLEdge.zip"), ("_2B9E6D14C7A3480FB1D5E9A2C6F8B347", "manifest.json"),
                ("_0AC4BB0B2D0448E3A2CF7CCB287D6D0E", "adxregistrator.exe"), ("_E369ACDA2B434E21BB456611DF57C948", "adxloader.dll"),
                ("_065924003BD749BC864364CB30199C11", "adxloader64.dll"), ("_1223A86871044FCB82F28F63D17AA70D", "XLEdge.dll (Built output)"),
            })
                StringAssert.Contains(twin, ":" + key + "\"", what);
        }

        [TestMethod]
        public void SetupProject_UninstallRegistration_IsSkippedDuringAnUpgrade()
        {
            // Review fix: on a major upgrade the old product is removed after the new one registered XLEdge; its /uninstall must
            // not undo that registration.
            string twin = File.ReadAllText(Vdproj).Replace("\r\n", "\n");
            Match uninstall = Regex.Match(twin, "\"Condition\" = \"8:([^\"]*)\"\n[^\n]*\n[^\n]*\n            \"InstallAction\" = \"3:4\"\n            \"Arguments\" = \"8:/uninstall=\\\\\"XLEdge.dll\\\\\" /privileges=user\"");
            Assert.IsTrue(uninstall.Success, "the uninstall registration action");
            Assert.AreEqual("NOT UPGRADINGPRODUCTCODE", uninstall.Groups[1].Value);
        }

        [TestMethod]
        public void SharedAssemblies_FileVersion_FollowsTheProduct()
        {
            Func<string, string> fileVersion = f => Regex.Match(File.ReadAllText(Path.Combine(TwinRoot, f)), "AssemblyFileVersion\\(\"([^\"]+)\"\\)").Groups[1].Value;
            Assert.AreEqual(fileVersion("XLEdgeSharedVersion.cs"), fileVersion("XLEdgeContractsVersion.cs"), "Contracts, Shared and Loader.Core show the release that built them");
        }

        // ---- the PreBuild signature check (XLEdge.Build\verify-signed.ps1) ----

        private static string Verify => Path.Combine(TwinRoot, @"XLEdge.Build\verify-signed.ps1");

        private static (int Exit, string Output) RunVerify(string shellBin, string shellObj, string payload, string skip = null)
        {
            var info = new System.Diagnostics.ProcessStartInfo("powershell.exe",
                $"-NoProfile -ExecutionPolicy Bypass -File \"{Verify}\" -ShellBin \"{shellBin}\" -ShellObj \"{shellObj}\" -Payload \"{payload}\"")
            { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            info.EnvironmentVariables.Remove("XLEDGE_SKIP_SIGNING");
            if (skip != null) info.EnvironmentVariables["XLEDGE_SKIP_SIGNING"] = skip;
            using (var process = System.Diagnostics.Process.Start(info))
            {
                string output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
                process.WaitForExit();
                return (process.ExitCode, output);
            }
        }

        /// <summary>A shell bin/obj and a payload zip, every file of ours a copy of <paramref name="source"/> (signed or not).</summary>
        private static (string Bin, string Obj, string Payload) Layout(TempFolder t, string source, string unsignedOne = null)
        {
            string bin = t.Combine("bin"), obj = t.Combine("obj"), zipDir = t.Combine("zip");
            foreach (string d in new[] { bin, obj, zipDir }) Directory.CreateDirectory(d);
            foreach (string f in new[] { "XLEdge.dll", "adxloader.dll", "adxloader64.dll", "XLEdge.Contracts.dll", "XLEdge.Shared.dll", "XLEdge.Loader.Core.dll" })
                File.Copy(source, Path.Combine(bin, f));
            File.Copy(source, Path.Combine(obj, "XLEdge.dll"));
            foreach (string f in new[] { "XLEdge.Addin.Core.dll", "XLEdge.Contracts.dll", "XLEdge.Shared.dll" }) File.Copy(source, Path.Combine(zipDir, f));
            if (unsignedOne != null)
                File.Copy(typeof(SetupProjectTests).Assembly.Location, Path.Combine(unsignedOne == "obj" ? obj : bin, unsignedOne == "obj" ? "XLEdge.dll" : unsignedOne), true);
            string payload = t.Combine("OrbitXLEdge.zip");
            System.IO.Compression.ZipFile.CreateFromDirectory(zipDir, payload);
            return (bin, obj, payload);
        }

        private static string SignedMicrosoftDll()
        {
            string dll = typeof(Assert).Assembly.Location;
            var check = new System.Diagnostics.ProcessStartInfo("powershell.exe", $"-NoProfile -Command \"(Get-AuthenticodeSignature '{dll}').Status\"")
            { UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true };
            using (var p = System.Diagnostics.Process.Start(check))
            {
                string status = p.StandardOutput.ReadToEnd().Trim();
                p.WaitForExit();
                if (status != "Valid") Assert.Inconclusive("No Authenticode-signed DLL to test with (" + dll + ": " + status + ").");
            }
            return dll;
        }

        [TestMethod]
        public void VerifySigned_AllSigned_Passes()
        {
            using (var t = new TempFolder())
            {
                var (bin, obj, payload) = Layout(t, SignedMicrosoftDll());
                var (exit, output) = RunVerify(bin, obj, payload);
                Assert.AreEqual(0, exit, output);
            }
        }

        [TestMethod]
        public void VerifySigned_AnUnsignedFile_FailsAndNamesIt()
        {
            foreach (string unsigned in new[] { "obj", "adxloader64.dll" })
                using (var t = new TempFolder())
                {
                    var (bin, obj, payload) = Layout(t, SignedMicrosoftDll(), unsigned);
                    var (exit, output) = RunVerify(bin, obj, payload);
                    Assert.AreEqual(1, exit, output);
                    StringAssert.Contains(output, "error XLEDGE0005");
                    StringAssert.Contains(output, unsigned == "obj" ? @"obj\XLEdge.dll" : "adxloader64.dll");
                }
        }

        [TestMethod]
        public void VerifySigned_UnsignedPayloadDll_Fails()
        {
            using (var t = new TempFolder())
            {
                var (bin, obj, _) = Layout(t, SignedMicrosoftDll());
                string zipDir = t.Combine("zip2");
                Directory.CreateDirectory(zipDir);
                File.Copy(typeof(SetupProjectTests).Assembly.Location, Path.Combine(zipDir, "XLEdge.Addin.Core.dll"));
                string payload = t.Combine("unsigned.zip");
                System.IO.Compression.ZipFile.CreateFromDirectory(zipDir, payload);
                var (exit, output) = RunVerify(bin, obj, payload);
                Assert.AreEqual(1, exit, output);
                StringAssert.Contains(output, "XLEdge.Addin.Core.dll");
            }
        }

        [TestMethod]
        public void VerifySigned_SkipVariable_WarnsThatTheMsiIsUnsigned()
        {
            using (var t = new TempFolder())
            {
                var (bin, obj, payload) = Layout(t, typeof(SetupProjectTests).Assembly.Location);
                var (exit, output) = RunVerify(bin, obj, payload, skip: "1");
                Assert.AreEqual(0, exit, output);
                StringAssert.Contains(output, "warning XLEDGE0004");
            }
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
