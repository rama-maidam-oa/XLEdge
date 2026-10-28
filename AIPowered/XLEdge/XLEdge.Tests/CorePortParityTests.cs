using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace XLEdge.Tests
{
    /// <summary>
    /// XLEdge 11.1.2 code moves into the core unchanged unless a ledgered ruling says otherwise: every file ported verbatim must
    /// stay byte-identical to its original (BOM and line endings aside), so a later edit to one copy cannot drift silently.
    /// </summary>
    [TestClass]
    public class CorePortParityTests
    {
        private static string TestBin => Path.GetDirectoryName(typeof(CorePortParityTests).Assembly.Location);
        private static string Today => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\..\..\XLEdge"));
        private static string Core => Path.GetFullPath(Path.Combine(TestBin, @"..\..\..\XLEdge.Addin.Core"));

        /// <summary>Ported unchanged (relative to the project folder in both).</summary>
        internal static readonly string[] Verbatim =
        {
            // Stage 2a
            @"Helpers\ExcelSheetHelper.cs", @"Utilities\DpiAwarenessHelper.cs", @"Utilities\WpfAppManager.cs", @"Utilities\UiDispatcher.cs", "XLApp.cs",
            // Stage 2b
            @"Helpers\ApiErrorMessageExtractor.cs", @"Helpers\XLEdgeValueFormatter.cs", @"Helpers\ApiResponseHelper.cs",
            @"Helpers\ImageDownloadHelper.cs", @"Helpers\XLEdgePreferencesManager.cs", @"Helpers\NumericJsonConverter.cs", @"Helpers\PerformanceHelper.cs",
            @"Helpers\JsonGlobals.cs", @"Helpers\CancellationHelper.cs", @"Helpers\AttachmentLinkHelper.cs",
            @"Helpers\XLEdgeTempFileCleaner.cs", @"Helpers\CancellationTokenHelper.cs", @"Helpers\ExceptionHelper.cs", @"Helpers\ApiExceptions.cs",
            @"Helpers\XLEdgeOperatorMappings.cs", @"Helpers\EdgeRequestParser.cs", @"Helpers\SerializationHelper.cs",
            @"Models\ReportMetaModels.cs", @"Models\AllModels.cs", @"Models\ApiResult.cs", @"Utilities\StrictCertificateValidator.cs", "XLEdgeAppConstants.cs",
        };

        /// <summary>Ported with ledgered changes forced by the AppDomain boundary (Stage 2a rulings): must exist, may differ.</summary>
        internal static readonly string[] Adapted =
        {
            @"Utilities\LogUtility.cs", "XLEdgeAppPaths.cs", "XLEdgeAppState.cs", @"Helpers\XLEdgeRibbonHelper.cs", @"Helpers\ExcelApplicationHelper.cs",
        };

        private static string Normalized(string path) => File.ReadAllText(path).TrimStart('﻿').Replace("\r\n", "\n");

        [TestMethod]
        public void PortedFiles_AreByteIdenticalToXLEdge11_1_2()
        {
            Assert.IsTrue(Directory.Exists(Today), Today);
            var problems = new List<string>();
            foreach (string file in Verbatim)
            {
                string original = Path.Combine(Today, file), copy = Path.Combine(Core, file);
                if (!File.Exists(original)) problems.Add("no original: " + file);
                else if (!File.Exists(copy)) problems.Add("not ported: " + file);
                else if (Normalized(original) != Normalized(copy)) problems.Add("differs from 11.1.2: " + file);
            }
            foreach (string file in Adapted.Where(f => !File.Exists(Path.Combine(Core, f))))
                problems.Add("adapted file missing: " + file);
            Assert.AreEqual(0, problems.Count, string.Join(Environment.NewLine, problems));
        }

        [TestMethod]
        public void EveryPortedFile_IsCompiledIntoTheCore()
        {
            string project = File.ReadAllText(Path.Combine(Core, "XLEdge.Addin.Core.csproj"));
            var missing = Verbatim.Concat(Adapted).Where(f => !project.Contains($"<Compile Include=\"{f}\"")).ToList();
            Assert.AreEqual(0, missing.Count, "not in XLEdge.Addin.Core.csproj: " + string.Join(", ", missing));
        }
    }
}
