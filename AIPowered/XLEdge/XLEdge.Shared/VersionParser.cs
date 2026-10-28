using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using XLEdge.Contracts;

namespace XLEdge.Shared
{
    /// <summary>Result of parsing a manifest (a JSON array of release entries); the scalar fields describe the latest entry.</summary>
    public class VersionParseResult
    {
        public bool Success { get; set; }
        public string Version { get; set; }
        public DateTime ReleaseDate { get; set; }
        public int ContractVersion { get; set; }
        public List<VersionInfo> AllVersions { get; set; } = new List<VersionInfo>();
        public string ErrorMessage { get; set; }
    }

    /// <summary>
    /// Reads manifest.json files (staged releases, offline release folders, the server's release index).
    /// Never throws: failures come back in <see cref="VersionParseResult.ErrorMessage"/> and are logged.
    /// </summary>
    public class VersionParser
    {
        // Release dates written by the build are "yyyy-MM-ddTHH:mm:ss" local time; the other two are older forms.
        private static readonly string[] ReleaseDateFormats = { "yyyy-MM-ddTHH:mm:ss", "dd-MMM-yyyy", "yyyy-MM-dd" };

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        };

        private readonly ILogger _logger;

        public VersionParser(ILogger logger = null)
        {
            _logger = logger;
        }

        public VersionParseResult ParseVersionJson(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return Fail("JSON content is empty or null");
            try
            {
                var versions = JsonSerializer.Deserialize<List<VersionInfo>>(json, JsonOptions);
                if (versions == null || versions.Count == 0) return Fail("No version data found in JSON");

                var latest = GetLatestVersion(versions);
                if (latest == null) return Fail("Failed to determine latest version");

                var result = new VersionParseResult
                {
                    Success = true,
                    Version = latest.Version,
                    ReleaseDate = ParseReleaseDate(latest.ReleaseDate),
                    ContractVersion = latest.ContractVersion,
                    AllVersions = versions,
                };
                _logger?.LogDebug($"Parsed version: {result.Version}, Release Date: {result.ReleaseDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture)}");
                return result;
            }
            catch (JsonException ex)
            {
                return Fail("Invalid JSON format: " + ex.Message, ex);
            }
            catch (Exception ex)
            {
                return Fail("Failed to parse version JSON: " + ex.Message, ex);
            }
        }

        public VersionParseResult ParseVersionFile(string path)
        {
            _logger?.LogDebug($"ParseVersionFile: reading manifest file '{path}'.");
            try
            {
                if (!File.Exists(path)) return Fail("Manifest file not found: " + path);
                return ParseVersionJson(File.ReadAllText(path));
            }
            catch (Exception ex)
            {
                return Fail("Failed to read manifest file: " + ex.Message, ex);
            }
        }

        /// <summary>The entry with the highest version (System.Version, 2-4 numeric parts); ties go to the first; null when none parses.</summary>
        public VersionInfo GetLatestVersion(IEnumerable<VersionInfo> versions)
        {
            if (versions == null) return null;
            return versions
                .Where(v => v != null && !string.IsNullOrWhiteSpace(v.Version))
                .Select(v => new { Info = v, Parsed = TryParseVersion(v.Version) })
                .Where(x => x.Parsed != null)
                .OrderByDescending(x => x.Parsed)
                .Select(x => x.Info)
                .FirstOrDefault();
        }

        private Version TryParseVersion(string value)
        {
            if (System.Version.TryParse(value, out var parsed)) return parsed;
            _logger?.LogWarn($"Could not parse version string '{value}' as a System.Version.");
            return null;
        }

        private DateTime ParseReleaseDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return DateTime.MinValue;
            if (DateTime.TryParseExact(value, ReleaseDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsed))
                return parsed;
            _logger?.LogWarn("Could not parse release date: " + value);
            return DateTime.MinValue;
        }

        private VersionParseResult Fail(string message, Exception ex = null)
        {
            _logger?.LogError(message, ex?.ToString());
            return new VersionParseResult { Success = false, ErrorMessage = message };
        }
    }
}
