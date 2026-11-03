using System;

namespace XLEdge.Shared
{
    /// <summary>
    /// How a version is shown to people (log lines, windows, messages): 11.1.3, not the assembly/manifest form 11.1.3.0.
    /// Only for display - catalog entries, folder names and comparisons keep the full version.
    /// </summary>
    public static class VersionText
    {
        /// <summary>"11.1.3.0" -> "11.1.3"; a non-zero revision ("11.1.3.5"), a shorter version or text that is not a version is returned trimmed.</summary>
        public static string Display(string version)
        {
            if (string.IsNullOrWhiteSpace(version)) return version;
            string text = version.Trim();
            return text.Split('.').Length == 4 && Version.TryParse(text, out Version parsed) && parsed.Revision == 0
                ? $"{parsed.Major}.{parsed.Minor}.{parsed.Build}"
                : text;
        }
    }
}
