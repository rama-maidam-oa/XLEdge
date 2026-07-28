using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLEdge
{
    public static class XLEdgeAppConstants
    {
        // Default version info (can be overridden at runtime if needed)
        public const string DefaultVersion = "11.1.0";
        public const string DefaultCommitDate = "23-Jul-2026";

        // Logging constants (used in LogHelper and AppPaths)
        public const long LogMaxFileSizeBytes = 20 * 1024 * 1024; // 20 MB
        public const int LogMaxArchiveFiles = 30;

        //Theme and accentHex
        public const string GLTheme = "Light";
        public const string GLAccentHex = "#149FDB";
    }
}
