using System.IO;
using System.IO.Compression;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Loader.Core;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>The core build publishes a release (zip + manifest.json) that the bootstrapper accepts and the loader runs.</summary>
    [TestClass]
    public class CorePackageTests
    {
        internal static string PackageFolder
        {
            get
            {
                string coreBin = ReleaseFolders.CoreBuildFolder;
                string config = new DirectoryInfo(coreBin).Name;
                return Path.GetFullPath(Path.Combine(coreBin, @"..\..\SetupFiles", config, "Manifest"));
            }
        }

        // ---- Stage 5: the Release signing hook (XLEdge.Build\sign.ps1) -----------------------------------------------------

        private static string TwinRoot => Path.GetFullPath(Path.Combine(Path.GetDirectoryName(typeof(CorePackageTests).Assembly.Location), @"..\..\.."));
        private static string SignScript => Path.Combine(TwinRoot, @"XLEdge.Build\sign.ps1");

        /// <summary>Runs sign.ps1 in a fresh PowerShell (XLEDGE_SKIP_SIGNING as given, not inherited): exit code and output.</summary>
        private static (int Exit, string Output) RunSign(string file, string configuration, string signScript, string skip = null)
        {
            var info = new System.Diagnostics.ProcessStartInfo("powershell.exe",
                $"-NoProfile -ExecutionPolicy Bypass -File \"{SignScript}\" -File \"{file}\" -Configuration {configuration}" +
                (signScript == null ? "" : $" -SignScript \"{signScript}\""))
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

        /// <summary>A stand-in sign_file.cmd: logs "file|configuration" next to itself and exits with <paramref name="exit"/>.</summary>
        private static string FakeSignFile(TempFolder t, int exit = 0)
        {
            string path = t.Combine("sign_file.cmd");
            File.WriteAllText(path, "@echo %~1^|%~2>>\"%~dp0signed.log\"\r\n@exit /b " + exit + "\r\n");
            return path;
        }

        [TestMethod]
        public void Sign_Release_CallsTheSignScriptWithFileAndConfiguration()
        {
            using (var t = new TempFolder())
            {
                var (exit, output) = RunSign(@"C:\x\a.dll", "Release", FakeSignFile(t));
                Assert.AreEqual(0, exit, output);
                CollectionAssert.AreEqual(new[] { @"C:\x\a.dll|Release" }, File.ReadAllLines(t.Combine("signed.log")));
            }
        }

        [TestMethod]
        public void Sign_Debug_DoesNothing()
        {
            using (var t = new TempFolder())
            {
                var (exit, output) = RunSign(@"C:\x\a.dll", "Debug", FakeSignFile(t));
                Assert.AreEqual(0, exit, output);
                Assert.IsFalse(File.Exists(t.Combine("signed.log")), "no signing in Debug");
            }
        }

        [TestMethod]
        public void Sign_SkipVariable_DoesNothingAndSaysSo()
        {
            using (var t = new TempFolder())
            {
                var (exit, output) = RunSign(@"C:\x\a.dll", "Release", FakeSignFile(t), skip: "1");
                Assert.AreEqual(0, exit, output);
                StringAssert.Contains(output, "signing skipped");
                Assert.IsFalse(File.Exists(t.Combine("signed.log")), "XLEDGE_SKIP_SIGNING=1: nothing signed");
            }
        }

        [TestMethod]
        public void Sign_Release_WithoutTheScript_FailsWithTheMessage()
        {
            using (var t = new TempFolder())
            {
                var (exit, output) = RunSign(@"C:\x\a.dll", "Release", t.Combine("sign_file.cmd"));
                Assert.AreEqual(1, exit, output);
                StringAssert.Contains(output, "sign_file.cmd not found");
                StringAssert.Contains(output, "XLEDGE_SKIP_SIGNING=1");
            }
        }

        [TestMethod]
        public void Sign_ScriptFails_TheBuildFails()
        {
            using (var t = new TempFolder())
            {
                var (exit, output) = RunSign(@"C:\x\a.dll", "Release", FakeSignFile(t, exit: 3));
                Assert.AreNotEqual(0, exit, output);
            }
        }

        [TestMethod]
        public void PostBuilds_SignWhatGLSenseSigns()
        {
            foreach (string project in new[] { @"XLEdge.Contracts\XLEdge.Contracts.csproj", @"XLEdge.Shared\XLEdge.Shared.csproj", @"XLEdge.Loader.Core\XLEdge.Loader.Core.csproj" })
            {
                string text = File.ReadAllText(Path.Combine(TwinRoot, project));
                StringAssert.Contains(text, @"XLEdge.Build\sign.ps1", project);
                StringAssert.Contains(text, "$(TargetPath)", project);
                StringAssert.Contains(text, "$(Configuration)", project);
            }
            string core = File.ReadAllText(Path.Combine(TwinRoot, @"XLEdge.Addin.Core\post_build.ps1"));
            int sign = core.IndexOf("sign.ps1", System.StringComparison.Ordinal), zip = core.IndexOf("ZipFile]::Open", System.StringComparison.Ordinal);
            Assert.IsTrue(sign >= 0 && sign < zip, "the core DLL is signed before it is zipped");
            StringAssert.Contains(File.ReadAllText(Path.Combine(TwinRoot, @"..\..\.gitignore")), "/AIPowered/XLEdge/sign_file.cmd");
        }

        // ---- Stage 5: the installer payload and release notes ---------------------------------------------------------------

        /// <summary>
        /// The core's post_build.ps1 run as in a build, in a copy of the project layout (TempFolder: XLEdge.Addin.Core\post_build.ps1
        /// + ReleaseNotes.txt, XLEdge.Build\sign.ps1, bin\ with the built core): its SetupFiles folder.
        /// </summary>
        private static (int Exit, string Output, string SetupFiles) RunCorePostBuild(TempFolder t, string configuration, string notes)
        {
            string project = t.Combine("XLEdge.Addin.Core"), build = t.Combine("XLEdge.Build"), bin = t.Combine("bin");
            Directory.CreateDirectory(project); Directory.CreateDirectory(build); Directory.CreateDirectory(bin);
            File.Copy(Path.Combine(TwinRoot, @"XLEdge.Addin.Core\post_build.ps1"), Path.Combine(project, "post_build.ps1"));
            File.Copy(SignScript, Path.Combine(build, "sign.ps1"));
            if (notes != null) File.WriteAllText(Path.Combine(project, "ReleaseNotes.txt"), notes);
            string coreBin = Path.Combine(TwinRoot, @"XLEdge.Addin.Core\bin\Debug");
            foreach (string dll in new[] { "XLEdge.Addin.Core.dll", "XLEdge.Contracts.dll" }) File.Copy(Path.Combine(coreBin, dll), Path.Combine(bin, dll));
            var info = new System.Diagnostics.ProcessStartInfo("powershell.exe",
                $"-NoProfile -ExecutionPolicy Bypass -File \"{Path.Combine(project, "post_build.ps1")}\" -TargetDir \"{bin} \" -Configuration {configuration}")
            { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
            info.EnvironmentVariables["XLEDGE_SKIP_SIGNING"] = "1";
            using (var process = System.Diagnostics.Process.Start(info))
            {
                string output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
                process.WaitForExit();
                return (process.ExitCode, output, Path.Combine(project, "SetupFiles"));
            }
        }

        private static string Sha256(byte[] bytes)
        {
            using (var sha = System.Security.Cryptography.SHA256.Create()) return System.BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", "");
        }

        [TestMethod]
        public void ReleaseBuild_CopiesThePublishedZipToTheInstallerPayload()
        {
            using (var t = new TempFolder())
            {
                var (exit, output, setup) = RunCorePostBuild(t, "Release", null);
                Assert.AreEqual(0, exit, output);
                string published = Directory.GetFiles(Path.Combine(setup, @"Release\Manifest"), "*.zip").Single();
                byte[] payload = File.ReadAllBytes(Path.Combine(setup, @"Release\InstallerPayload\OrbitXLEdge.zip"));
                CollectionAssert.AreEqual(File.ReadAllBytes(published), payload, "the payload is the published zip, byte for byte");
                StringAssert.Contains(File.ReadAllText(Path.Combine(setup, @"Release\Manifest\manifest.json")), Sha256(payload),
                    "the seeded manifest's checksum matches the payload (else the first core never installs)");
                Assert.AreEqual(1, Directory.GetFiles(Path.Combine(setup, @"Release\InstallerPayload")).Length, "only the payload");
            }
        }

        [TestMethod]
        public void DebugBuild_WritesNoInstallerPayload()
        {
            using (var t = new TempFolder())
            {
                var (exit, output, setup) = RunCorePostBuild(t, "Debug", null);
                Assert.AreEqual(0, exit, output);
                Assert.IsFalse(Directory.Exists(Path.Combine(setup, @"Debug\InstallerPayload")) || Directory.Exists(Path.Combine(setup, @"Release\InstallerPayload")));
            }
        }

        [TestMethod]
        public void ReleaseNotes_FirstLine_BecomesTheNotes_Trimmed()
        {
            foreach (var (notes, expected) in new[] { ("  fixes the totals  \r\n# comment", "fixes the totals"), ("\r\n# comment", "XLEdge core build") })
                using (var t = new TempFolder())
                {
                    var (exit, output, setup) = RunCorePostBuild(t, "Debug", notes);
                    Assert.AreEqual(0, exit, output);
                    StringAssert.Contains(File.ReadAllText(Path.Combine(setup, @"Debug\Manifest\manifest.json")), "\"notes\":  \"" + expected + "\"");
                }
            Assert.IsTrue(File.Exists(Path.Combine(TwinRoot, @"XLEdge.Addin.Core\ReleaseNotes.txt")), "the project carries its ReleaseNotes.txt");
        }

        [TestMethod]
        public void BuiltPackage_IsAValidRelease()
        {
            string folder = PackageFolder;
            Assert.IsTrue(Directory.Exists(folder), "no package folder " + folder);
            var zips = Directory.GetFiles(folder, "*.zip");
            Assert.AreEqual(1, zips.Length, "exactly one zip");
            string manifestPath = Path.Combine(folder, "manifest.json");
            Assert.IsTrue(File.Exists(manifestPath));

            var parsed = new VersionParser().ParseVersionFile(manifestPath);
            Assert.IsTrue(parsed.Success, parsed.ErrorMessage);
            VersionInfo release = parsed.AllVersions.Single();
            byte[] zip = File.ReadAllBytes(zips[0]);
            Assert.AreEqual(ReleaseFolders.Sha256(zip), release.Checksum);
            Assert.AreEqual(ContractInfo.Version, release.ContractVersion);
            Assert.IsTrue(ReleaseHistoryStore.IsSafeFolderName(release.FolderName), release.FolderName);
            Assert.AreEqual(release.FolderName + ".zip", release.FileName);
            Assert.AreEqual(Path.GetFileName(zips[0]), release.FileName);
            Assert.IsFalse(File.ReadAllBytes(manifestPath).Take(3).SequenceEqual(new byte[] { 0xEF, 0xBB, 0xBF }), "manifest.json without BOM");

            using (var archive = new ZipArchive(new MemoryStream(zip)))
            {
                var names = archive.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.IsSubsetOf(new[] { "XLEdge.Addin.Core.dll", "XLEdge.Contracts.dll", "XLEdge.Shared.dll" }, names);
                Assert.IsFalse(names.Any(n => n.EndsWith(".pdb") || n.EndsWith(".xml") || n.EndsWith(".zip")), string.Join(", ", names));
                CollectionAssert.DoesNotContain(names, "XLEdge.Loader.Core.dll", "the shell supplies the loader");
            }
        }

        [TestMethod]
        public void BuiltPackage_CarriesWebView2()
        {
            // The pane content is a WinForms WebView2 in the core; WebView2 finds its native loader under runtimes\ next to Core.dll.
            using (var archive = new ZipArchive(File.OpenRead(Directory.GetFiles(PackageFolder, "*.zip").Single())))
            {
                var names = archive.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.IsSubsetOf(new[]
                {
                    "Microsoft.Web.WebView2.Core.dll", "Microsoft.Web.WebView2.WinForms.dll",
                    "runtimes/win-x86/native/WebView2Loader.dll", "runtimes/win-x64/native/WebView2Loader.dll",
                }, names, string.Join(", ", names));
                Assert.AreEqual(new System.Version(1, 0, 4022, 49), System.Reflection.AssemblyName.GetAssemblyName(
                    Path.Combine(ReleaseFolders.CoreBuildFolder, "Microsoft.Web.WebView2.Core.dll")).Version, "today's XLEdge WebView2 version");
            }
        }

        [TestMethod]
        public void BuiltPackage_CarriesTheJsonDependencies()
        {
            // The ported helpers use System.Text.Json; the tests run from bin\, so only this catches a release zip without them.
            using (var archive = new ZipArchive(File.OpenRead(Directory.GetFiles(PackageFolder, "*.zip").Single())))
            {
                var names = archive.Entries.Select(e => e.FullName).ToList();
                CollectionAssert.IsSubsetOf(new[]
                {
                    "System.Text.Json.dll", "System.Text.Encodings.Web.dll", "System.Memory.dll", "System.Buffers.dll", "System.IO.Pipelines.dll",
                    "System.Numerics.Vectors.dll", "System.Runtime.CompilerServices.Unsafe.dll", "System.Threading.Tasks.Extensions.dll",
                    "Microsoft.Bcl.AsyncInterfaces.dll",
                }, names, string.Join(", ", names));
            }
        }

        [TestMethod]
        public void ShellBuild_StagesExactlyTheCorePackage_AndNoCoreNextToTheShell()
        {
            string coreBin = ReleaseFolders.CoreBuildFolder;
            string config = new DirectoryInfo(coreBin).Name;
            string shellBin = Path.GetFullPath(Path.Combine(coreBin, @"..\..\..\XLEdge\bin", config));
            string staged = Path.Combine(shellBin, "AddinCore", "Manifest");
            var package = Directory.GetFiles(PackageFolder).Select(Path.GetFileName).OrderBy(n => n).ToArray();
            CollectionAssert.AreEqual(package, Directory.GetFiles(staged).Select(Path.GetFileName).OrderBy(n => n).ToArray());
            foreach (string name in package)
                CollectionAssert.AreEqual(File.ReadAllBytes(Path.Combine(PackageFolder, name)), File.ReadAllBytes(Path.Combine(staged, name)), name);
            Assert.IsFalse(File.Exists(Path.Combine(shellBin, "XLEdge.Addin.Core.dll")));
            foreach (string loader in new[] { "adxloader.dll", "adxloader64.dll", "adxloader.dll.manifest" })
                Assert.IsTrue(File.Exists(Path.Combine(shellBin, loader)), loader);
        }

        [TestMethod]
        public void BuiltPackage_RunsWithTheShellsRealLoggerAndPaths()
        {
            // The shell passes XLEdge.Shared's Logger/PathProvider (NLog, System.Text.Json) into the core's AppDomain.
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                var context = new FakeContext(paths) { DebugMode = true };
                using (var logger = new Logger(context))
                {
                    context.Logger = logger;
                    string folder = PackageFolder;
                    var release = new UpdateBootstrapper().ExtractAndCatalog(context,
                        File.ReadAllText(Path.Combine(folder, "manifest.json")), File.ReadAllBytes(Directory.GetFiles(folder, "*.zip").Single()), "Install");
                    Assert.IsNotNull(release);
                    var loader = new AddinDomainLoader();
                    try
                    {
                        var addin = loader.Load(context, release);
                        addin.OnRibbonAction("RibEdgeRefresh", null);
                        addin.OnRibbonAction("RibEdgeLogout_PropertyChanging", "Visible");
                    }
                    finally
                    {
                        loader.Unload(context);
                    }
                }
                string log = File.ReadAllText(Directory.GetFiles(paths.Logs, "XLEdge_Logs_*.log").Single());
                StringAssert.Contains(log, "AddinEntry.Initialize: XLEdge core initialized in AppDomain 'XLEdgeCore_");
                StringAssert.Contains(log, "AddinEntry.OnRibbonAction: 'RibEdgeRefresh'.");
                // Ribbon invalidations are not traced (one per ribbon refresh, not a user action).
                Assert.IsFalse(log.Contains("RibEdgeLogout_PropertyChanging"), log);
            }
        }

        [TestMethod]
        public void BuiltPackage_InstallsAndLoads()
        {
            using (var t = new TempFolder())
            {
                var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
                paths.Ensure();
                var context = new FakeContext(paths) { Logger = new RecordingLogger() };
                string folder = PackageFolder;
                var release = new UpdateBootstrapper().ExtractAndCatalog(context,
                    File.ReadAllText(Path.Combine(folder, "manifest.json")), File.ReadAllBytes(Directory.GetFiles(folder, "*.zip").Single()), "Install");
                Assert.IsNotNull(release, ((RecordingLogger)context.Logger).Dump());

                var loader = new AddinDomainLoader();
                try
                {
                    var addin = loader.Load(context, release);
                    Assert.IsFalse(addin.IsBusy());
                }
                finally
                {
                    loader.Unload(context);
                }
            }
        }
    }
}
