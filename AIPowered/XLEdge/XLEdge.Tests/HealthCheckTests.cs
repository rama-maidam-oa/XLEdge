using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Win32;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 7: tools\XLEdgeHealthCheck.ps1 - what QA runs on a test machine. Each test builds a fake machine: an install root
    /// (the shell's real build output), a data root with a log, and a temporary HKCU key standing in for HKCU\Software (removed
    /// after). The script only reads.
    /// </summary>
    [TestClass]
    public class HealthCheckTests
    {
        private const string Clsid = "{80B0FB76-A5F4-41E0-B283-C175B7374B60}";
        private static string TwinRoot => Path.GetFullPath(Path.Combine(Path.GetDirectoryName(typeof(HealthCheckTests).Assembly.Location), @"..\..\.."));
        private static string Script => Path.Combine(TwinRoot, @"tools\XLEdgeHealthCheck.ps1");
        private static string ShellBuild => Path.Combine(TwinRoot, @"XLEdge\bin\Debug");

        private sealed class Machine : IDisposable
        {
            private readonly TempFolder _temp = new TempFolder();
            public readonly string RegistryKey = @"Software\XLEdgeHealthCheckTests_" + Guid.NewGuid().ToString("N");
            public string InstallRoot => _temp.Combine("install");
            public string DataRoot => _temp.Combine("data");
            public string LogFile => Path.Combine(DataRoot, @"XLEdge_Logs\Logs", "XLEdge_Logs_" + DateTime.Now.ToString("dd-MMM-yyyy", System.Globalization.CultureInfo.InvariantCulture) + ".log");

            /// <param name="addinCore">false: an AddinCore never set up.</param>
            /// <param name="oldSingleDll">11.1.2's install: XLEdge.dll and the loaders only (its MSI ships no Contracts/Shared/Loader.Core), no AddinCore.</param>
            public Machine(bool registered = true, bool addinCore = true, bool oldSingleDll = false)
            {
                Directory.CreateDirectory(InstallRoot);
                if (oldSingleDll) addinCore = false;
                var dlls = oldSingleDll
                    ? new[] { "XLEdge.dll", "adxloader.dll", "adxloader64.dll" }
                    : new[] { "XLEdge.dll", "XLEdge.Contracts.dll", "XLEdge.Shared.dll", "XLEdge.Loader.Core.dll", "adxloader.dll", "adxloader64.dll" };
                foreach (string dll in dlls)
                    File.Copy(Path.Combine(ShellBuild, dll), Path.Combine(InstallRoot, dll));
                if (addinCore)
                {
                    string versions = Path.Combine(InstallRoot, @"AddinCore\Versions");
                    Directory.CreateDirectory(Path.Combine(versions, "vA"));
                    File.WriteAllText(Path.Combine(versions, "vA", "XLEdge.Addin.Core.dll"), "x");
                    ReleaseHistoryStore.Append(Path.Combine(InstallRoot, @"AddinCore\ReleaseHistory.json"),
                        new ReleaseEntry { Version = "11.1.3.0", ReleaseDate = "2026-11-03T10:00:00", FolderName = "vA", ContractVersion = 5 });
                    File.WriteAllText(Path.Combine(InstallRoot, @"AddinCore\SelectedRelease.txt"), "vA");
                }
                Directory.CreateDirectory(Path.GetDirectoryName(LogFile));
                File.WriteAllText(LogFile, LogLine("INFO  ", "10:00:00", "started"));
                if (registered) Register(Path.Combine(InstallRoot, "adxloader64.dll"));
            }

            /// <summary>A line as XLEdge.Shared.Logger writes it: NLog's "${longdate}|" before the level (Logger.cs Layout and Format).</summary>
            public static string LogLine(string level, string time, string message, string indent = "") =>
                DateTime.Now.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture) + " " + time + ".1234|" + indent + level + "| " + time + " | " + message + "\r\n";

            public void Register(string loader, int loadBehavior = 3)
            {
                using (var addin = Registry.CurrentUser.CreateSubKey(RegistryKey + @"\Microsoft\Office\Excel\Addins\XLEdge.AddinModule"))
                    addin.SetValue("LoadBehavior", loadBehavior, RegistryValueKind.DWord);
                using (var server = Registry.CurrentUser.CreateSubKey(RegistryKey + @"\Classes\CLSID\" + Clsid + @"\InprocServer32"))
                    server.SetValue("", loader);
            }

            public (int Exit, Dictionary<string, (string Status, string Detail)> Checks, string Output) Run()
            {
                var info = new System.Diagnostics.ProcessStartInfo("powershell.exe",
                    $"-NoProfile -ExecutionPolicy Bypass -File \"{Script}\" -RegistryRoot \"HKCU:\\{RegistryKey}\" -DataRoot \"{DataRoot}\" -Json")
                { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true, CreateNoWindow = true };
                using (var process = System.Diagnostics.Process.Start(info))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string errors = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    var checks = new Dictionary<string, (string, string)>(StringComparer.Ordinal);
                    try
                    {
                        using (JsonDocument doc = JsonDocument.Parse(output))
                            foreach (JsonElement check in doc.RootElement.EnumerateArray())
                                checks[check.GetProperty("check").GetString()] = (check.GetProperty("status").GetString(), check.GetProperty("detail").GetString());
                    }
                    catch (JsonException) { }
                    return (process.ExitCode, checks, output + errors);
                }
            }

            public void Dispose()
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(RegistryKey, false); } catch (Exception) { }
                _temp.Dispose();
            }
        }

        [TestMethod]
        public void HealthCheck_HealthyInstall_HasNoFail()
        {
            using (var m = new Machine())
            {
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(0, exit, output);
                foreach (string name in new[] { "Registration", "Loader", "Shell", "Signature", "SharedAssemblies", "AddinCore", "SelectedRelease", "StagedRelease", "Log", "GLSense" })
                    Assert.IsTrue(checks.ContainsKey(name), name + " reported: " + output);
                Assert.IsFalse(checks.Values.Any(c => c.Status == "FAIL"), output);
                Assert.AreEqual("PASS", checks["Registration"].Status);
                StringAssert.Contains(checks["Loader"].Detail, "adxloader64.dll");
                StringAssert.Contains(checks["Shell"].Detail, "11.1.3.0");
                Assert.AreEqual("WARN", checks["Signature"].Status, "a development build is not signed");
                StringAssert.Contains(checks["AddinCore"].Detail, "11.1.3.0");
                StringAssert.Contains(checks["SelectedRelease"].Detail, "vA");
                StringAssert.Contains(checks["GLSense"].Detail, "not installed");
            }
        }

        [TestMethod]
        public void HealthCheck_NotRegistered_Fails()
        {
            using (var m = new Machine(registered: false))
            {
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(1, exit, output);
                Assert.AreEqual("FAIL", checks["Registration"].Status, output);
            }
        }

        [TestMethod]
        public void HealthCheck_LoadBehaviorNotThree_Warns()
        {
            using (var m = new Machine(registered: false))
            {
                m.Register(Path.Combine(m.InstallRoot, "adxloader64.dll"), loadBehavior: 2);
                var (_, checks, output) = m.Run();
                Assert.AreEqual("WARN", checks["Registration"].Status, output);
                StringAssert.Contains(checks["Registration"].Detail, "LoadBehavior 2");
            }
        }

        [TestMethod]
        public void HealthCheck_LoaderMissing_FailsWithThePath()
        {
            using (var m = new Machine(registered: false))
            {
                string gone = Path.Combine(m.InstallRoot, @"gone\adxloader64.dll");
                m.Register(gone);
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(1, exit, output);
                Assert.AreEqual("FAIL", checks["Loader"].Status, output);
                StringAssert.Contains(checks["Loader"].Detail, gone);
            }
        }

        [TestMethod]
        public void HealthCheck_SelectedReleaseGone_Fails()
        {
            using (var m = new Machine())
            {
                File.WriteAllText(Path.Combine(m.InstallRoot, @"AddinCore\SelectedRelease.txt"), "vMissing");
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(1, exit, output);
                Assert.AreEqual("FAIL", checks["SelectedRelease"].Status, output);
                StringAssert.Contains(checks["SelectedRelease"].Detail, "vMissing");
            }
        }

        [TestMethod]
        public void HealthCheck_LogErrors_AreCountedAndListed()
        {
            using (var m = new Machine())
            {
                // Review: the real lines start with NLog's longdate; an "ERROR |" inside a message is not a level.
                File.AppendAllText(m.LogFile, Machine.LogLine("ERROR ", "10:01:00", "first failure") + Machine.LogLine("INFO  ", "10:02:00", "fine ERROR | not a level")
                    + Machine.LogLine("ERROR ", "10:03:00", "second failure", indent: "    "));
                var (_, checks, output) = m.Run();
                StringAssert.Contains(checks["Log"].Detail, "2 ERROR", output);
                StringAssert.Contains(checks["Log"].Detail, "first failure");
                StringAssert.Contains(checks["Log"].Detail, "second failure");
            }
        }

        [TestMethod]
        public void HealthCheck_SeveralReleases_AreEachListed()
        {
            // Found on a real machine: Windows PowerShell's ConvertFrom-Json hands a JSON array over as one object.
            using (var m = new Machine())
            {
                Directory.CreateDirectory(Path.Combine(m.InstallRoot, @"AddinCore\Versions\vB"));
                ReleaseHistoryStore.Append(Path.Combine(m.InstallRoot, @"AddinCore\ReleaseHistory.json"),
                    new ReleaseEntry { Version = "11.1.4.0", ReleaseDate = "2026-11-10T10:00:00", FolderName = "vB", ContractVersion = 5, Source = "Online" });
                var (_, checks, output) = m.Run();
                StringAssert.Contains(checks["AddinCore"].Detail, "2 catalogued release(s): 11.1.3.0 (vA, Install); 11.1.4.0 (vB, Online)", output);
            }
        }

        [TestMethod]
        public void HealthCheck_OldSingleDllInstall_WarnsNotSetUp()
        {
            // Review: 11.1.2's MSI ships no Contracts/Shared/Loader.Core - the way back (runbook B2) must not FAIL.
            using (var m = new Machine(oldSingleDll: true))
            {
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(0, exit, output);
                Assert.AreEqual("WARN", checks["AddinCore"].Status, output);
                StringAssert.Contains(checks["AddinCore"].Detail, "not set up");
                Assert.AreEqual("WARN", checks["SharedAssemblies"].Status, output);
                StringAssert.Contains(checks["SharedAssemblies"].Detail, "single-DLL");
            }
        }

        [TestMethod]
        public void HealthCheck_SharedAssemblyMissingNextToTheTwin_Fails()
        {
            using (var m = new Machine())
            {
                File.Delete(Path.Combine(m.InstallRoot, "XLEdge.Shared.dll"));
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(1, exit, output);
                Assert.AreEqual("FAIL", checks["SharedAssemblies"].Status, output);
                StringAssert.Contains(checks["SharedAssemblies"].Detail, "XLEdge.Shared.dll");
            }
        }

        [TestMethod]
        public void HealthCheck_EmptySelectedRelease_IsNoChoice()
        {
            // Review: UpdateBootstrapper.RememberLoaded can write "" - the loader reads that as no choice.
            using (var m = new Machine())
            {
                File.WriteAllText(Path.Combine(m.InstallRoot, @"AddinCore\SelectedRelease.txt"), "");
                var (exit, checks, output) = m.Run();
                Assert.AreEqual(0, exit, output);
                Assert.AreEqual("PASS", checks["SelectedRelease"].Status, output);
                StringAssert.Contains(checks["SelectedRelease"].Detail, "No release chosen yet");
            }
        }
    }
}
