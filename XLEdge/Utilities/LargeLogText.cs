using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace XLEdge.Utilities
{
    /// <summary>
    /// Keeps very large texts (API responses - report CSV data or JSON - and raw JSON) out of the main log file. Ported from
    /// GLSense AIPowered's GLSense.Shared.LargeLogText (OISR-22830). Report data (whatever its size) and any other text over
    /// <see cref="ThresholdChars"/> is written in full to the add-in's Temp folder (XLEdge_Logs\Temp) and the main log gets only its size and that file's path, so the debug
    /// log stays readable. This replaces the "Include Output Data" ribbon option: with Debug on, the full data is always
    /// kept - small texts in the log, large ones in a Temp file. <see cref="CleanUp"/> runs at add-in start: it creates the
    /// folder if missing and removes this helper's own files older than <see cref="KeepDays"/> days (other files in the
    /// folder, e.g. report working CSVs, are left to XLEdgeTempFileCleaner).
    /// </summary>
    public static class LargeLogText
    {
        public const int ThresholdChars = 100 * 1024;
        public const int PreviewChars = 1000;
        public const int KeepDays = 7;

        public const string JsonExtension = ".json";
        public const string CsvExtension = ".csv";
        public const string TextExtension = ".txt";

        // File names written by NewFilePath: yyyyMMdd_HHmmss_fff_<name>.<ext>
        private static readonly Regex OwnFileName = new Regex(@"^\d{8}_\d{6}_\d{3}_", RegexOptions.Compiled);

        /// <summary>True for a file this helper wrote (its name starts with NewFilePath's timestamp).</summary>
        public static bool IsOwnFile(string filePath) =>
            !string.IsNullOrEmpty(filePath) && OwnFileName.IsMatch(Path.GetFileName(filePath));

        private static string Folder(string tempFolder) =>
            string.IsNullOrWhiteSpace(tempFolder) ? Path.GetTempPath() : tempFolder;

        /// <summary>New, unique file path in the Temp folder (folder created if needed).</summary>
        public static string NewFilePath(string tempFolder, string nameHint, string extension = JsonExtension)
        {
            string folder = Folder(tempFolder);
            Directory.CreateDirectory(folder);
            if (string.IsNullOrWhiteSpace(extension)) extension = TextExtension;
            if (!extension.StartsWith(".", StringComparison.Ordinal)) extension = "." + extension;
            return Path.Combine(folder, $"{DateTime.Now:yyyyMMdd_HHmmss_fff}_{SafeName(nameHint)}{extension}");
        }

        /// <summary>
        /// The file extension for <paramref name="text"/>: from the response's media type when it names JSON or CSV,
        /// otherwise from the content - a text starting with '{' or '[' is JSON, an HTML page is .txt, anything else
        /// (XLEdge's report data) is CSV.
        /// </summary>
        public static string ExtensionFor(string text, string mediaType = null)
        {
            if (!string.IsNullOrWhiteSpace(mediaType))
            {
                if (mediaType.IndexOf("json", StringComparison.OrdinalIgnoreCase) >= 0) return JsonExtension;
                if (mediaType.IndexOf("csv", StringComparison.OrdinalIgnoreCase) >= 0) return CsvExtension;
            }

            if (string.IsNullOrEmpty(text)) return TextExtension;

            int i = 0;
            while (i < text.Length && (char.IsWhiteSpace(text[i]) || text[i] == '﻿')) i++;
            if (i >= text.Length) return TextExtension;
            if (text[i] == '{' || text[i] == '[') return JsonExtension;
            if (text[i] == '<') return TextExtension;
            return CsvExtension;
        }

        /// <summary>
        /// Returns <paramref name="text"/> unchanged when it is small, unless <paramref name="alwaysSave"/> is set (report
        /// data: saved whatever its size). Otherwise saves it to the Temp folder (extension from <see cref="ExtensionFor"/>)
        /// and returns only its size and the saved file's path - none of the text. When the file cannot be written, a
        /// preview is returned instead, since the log is then the only record.
        /// </summary>
        public static string ToLogText(string text, string tempFolder, string nameHint, string mediaType = null, bool alwaysSave = false)
        {
            if (string.IsNullOrEmpty(text) || (!alwaysSave && text.Length <= ThresholdChars))
                return text;

            try
            {
                string path = NewFilePath(tempFolder, nameHint, ExtensionFor(text, mediaType));
                File.WriteAllText(path, text);
                return $"[{text.Length:N0} chars saved to {path}]";
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                return $"{text.Substring(0, Math.Min(PreviewChars, text.Length))}... [{text.Length:N0} chars - truncated, could not save the full text: {ex.Message}]";
            }
        }

        /// <summary>
        /// Creates the Temp folder if missing and deletes this helper's own files in it that are older than
        /// <paramref name="keepDays"/> days.
        /// </summary>
        public static void CleanUp(string tempFolder, int keepDays = KeepDays)
        {
            try
            {
                string folder = Folder(tempFolder);
                Directory.CreateDirectory(folder);

                DateTime cutoff = DateTime.Now.AddDays(-keepDays);
                foreach (string file in Directory.EnumerateFiles(folder))
                {
                    try
                    {
                        if (IsOwnFile(file) && File.GetLastWriteTime(file) < cutoff)
                            File.Delete(file);
                    }
                    catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
                    {
                        // In use or locked - leave it for the next start.
                    }
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Clean-up is best effort and must never stop the add-in from starting.
            }
        }

        private static string SafeName(string hint)
        {
            if (string.IsNullOrWhiteSpace(hint))
                return "text";

            string name = new string(hint.Select(c => char.IsLetterOrDigit(c) || c == '-' || c == '_' ? c : '_').ToArray()).Trim('_');
            if (name.Length > 60)
                name = name.Substring(0, 60);
            return name.Length == 0 ? "text" : name;
        }
    }
}
