using System.Collections.Generic;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 2f: XLEdge 11.1.2 calls GLSense's GetGLCubeInformation(token, url, userName) and LogoutSession() late-bound; GLSense
    /// AIPowered must expose them on its COM add-in object and route them to its core. GLSense has no test project, so its source
    /// is checked here, on each branch's checkout (Inconclusive where that checkout is not on this machine).
    /// </summary>
    [TestClass]
    public class GLSenseSourceTests
    {
        private static readonly string[] Checkouts =
        {
            @"D:\SQLLite_Test\worktrees\GLSense-11.1.2\AIPowered\GLSense", // 11.1.2
            @"D:\SQLLite_Test\GLSense\AIPowered\GLSense",                 // 11.1.3
        };

        [TestMethod]
        public void GLSenseShell_ExposesTheMethodsXLEdgeCalls()
        {
            var problems = new List<string>();
            int found = 0;
            foreach (string root in Checkouts)
            {
                string shell = Path.Combine(root, @"GLSense\AddinModule.cs"), core = Path.Combine(root, @"GLSense.Addin.Core\AddinEntry.cs");
                if (!File.Exists(shell) || !File.Exists(core)) continue;
                found++;
                string shellText = File.ReadAllText(shell), coreText = File.ReadAllText(core);
                foreach (string expected in new[] { "public void GetGLCubeInformation(string token, string url, string userName)",
                                                    "public void LogoutSession()",
                                                    "OnRibbonAction(\"LoginFromXLEdge\", new object[] { token, url, userName })",
                                                    "OnRibbonAction(\"LogoutFromXLEdge\", null)" })
                    if (!shellText.Contains(expected)) problems.Add($"{shell}: missing {expected}");
                foreach (string expected in new[] { "case \"LoginFromXLEdge\":", "case \"LogoutFromXLEdge\":" })
                    if (!coreText.Contains(expected)) problems.Add($"{core}: missing {expected}");
            }
            if (found == 0) Assert.Inconclusive("No GLSense AIPowered checkout on this machine.");
            Assert.AreEqual(0, problems.Count, string.Join(System.Environment.NewLine, problems));
        }
    }
}
