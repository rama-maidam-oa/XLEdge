using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Parses "isFileAttached" report cells (raw HTML anchor snippets from Oracle EBS/Fusion) into a
    /// compact "ATTACHMENT|key=value|...|downloadType" string used as a hyperlink's ScreenTip, and
    /// rebuilds the actual download URL from that string when the link is clicked.
    /// </summary>
    public static class AttachmentLinkHelper
    {
        private static readonly Regex DownloadTypeRegex = new Regex(@"fileDownLoadByHref\(event,""(?<type>.*?)""", RegexOptions.Compiled);
        private static readonly Regex KeyValueBlockRegex = new Regex(@"\{(?<inner>.*?)\}", RegexOptions.Compiled);
        private static readonly Regex DisplayValueRegex = new Regex(@">(?<display>[^<]+)<", RegexOptions.Compiled);

        public static bool TryParseAttachmentLink(string rawCellText, out string displayValue, out string linkValue)
        {
            displayValue = string.Empty;
            linkValue = string.Empty;

            if (string.IsNullOrWhiteSpace(rawCellText))
            {
                return false;
            }

            try
            {
                string downloadType = string.Empty;
                Match typeMatch = DownloadTypeRegex.Match(rawCellText);
                if (typeMatch.Success)
                {
                    downloadType = typeMatch.Groups["type"].Value;
                }

                var orderedKeyValues = new List<KeyValuePair<string, string>>();
                Match kvMatch = KeyValueBlockRegex.Match(rawCellText);
                if (kvMatch.Success)
                {
                    foreach (string part in kvMatch.Groups["inner"].Value.Split(','))
                    {
                        string[] kv = part.Split(':');
                        if (kv.Length == 2)
                        {
                            string key = kv[0].Trim();
                            string value = kv[1].Trim().Trim('{', '}');
                            orderedKeyValues.Add(new KeyValuePair<string, string>(key, value));
                        }
                    }
                }

                Match dispMatch = DisplayValueRegex.Match(rawCellText);
                if (dispMatch.Success)
                {
                    displayValue = dispMatch.Groups["display"].Value;
                }

                int businessObjIndex = orderedKeyValues.FindIndex(
                    kv => string.Equals(kv.Key, "businessObjectId", StringComparison.OrdinalIgnoreCase));
                List<KeyValuePair<string, string>> others = orderedKeyValues
                    .Where(kv => !string.Equals(kv.Key, "businessObjectId", StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var linkParts = new List<string> { "ATTACHMENT" };
                if (businessObjIndex >= 0)
                {
                    KeyValuePair<string, string> businessObj = orderedKeyValues[businessObjIndex];
                    linkParts.Add($"{businessObj.Key}={businessObj.Value}");
                }

                foreach (KeyValuePair<string, string> kv in others)
                {
                    linkParts.Add($"{kv.Key}={kv.Value}");
                }

                linkParts.Add(downloadType);
                linkValue = string.Join("|", linkParts);

                return !string.IsNullOrWhiteSpace(displayValue) && !string.IsNullOrWhiteSpace(linkValue);
            }
            catch (Exception ex)
            {
                LogUtility.LogException(ex, nameof(TryParseAttachmentLink));
                displayValue = string.Empty;
                linkValue = string.Empty;
                return false;
            }
        }

        /// <summary>
        /// Builds the attachment download URL from a stored "ATTACHMENT|key1=value1|key2=value2|...|downloadType"
        /// ScreenTip. Falls back to the "get-oracle-lob-file" endpoint when no explicit download type is present,
        /// and requires at least the two key=value segments (businessObjectId + one other id).
        /// </summary>
        public static string BuildDownloadUrl(string screenTip, string loginUrl)
        {
            if (string.IsNullOrWhiteSpace(screenTip) || string.IsNullOrWhiteSpace(loginUrl))
            {
                return null;
            }

            string[] parts = screenTip.Split('|');
            string baseUrl = loginUrl.TrimEnd('/');

            if (parts.Length < 3)
            {
                return null;
            }

            if (parts.Length >= 4 && !string.IsNullOrWhiteSpace(parts[3]))
            {
                return $"{baseUrl}/web/secure/{parts[3]}?{parts[1]}&{parts[2]}";
            }

            return $"{baseUrl}/web/secure/get-oracle-lob-file?{parts[1]}&{parts[2]}";
        }
    }
}
