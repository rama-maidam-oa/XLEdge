using System;
using System.Collections.Generic;
using System.Globalization;
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

        /// <summary>
        /// Placeholder a report formula writes for the report's business object id, e.g.
        /// "fileDownLoadByHref(event,"fusion-attachments",{documentId:12,businessObjectId:{businessObjectId}})".
        /// The web report grid replaces it with the report's businessObjectId before the link is used.
        /// </summary>
        public const string BusinessObjectIdPlaceholder = "{businessObjectId}";

        /// <summary>
        /// OISR-22766: parses the cell like <see cref="TryParseAttachmentLink(string, out string, out string)"/> after
        /// replacing <see cref="BusinessObjectIdPlaceholder"/> with the report's business object id, as the web report
        /// does. Without it the link sent "businessObjectId=businessObjectId" and the server's fusion-attachments /
        /// get-oracle-lob-file endpoints (businessObjectId is a number there) answered 400.
        /// </summary>
        public static bool TryParseAttachmentLink(string rawCellText, int businessObjectId, out string displayValue, out string linkValue)
        {
            return TryParseAttachmentLink(ReplaceBusinessObjectIdPlaceholder(rawCellText, businessObjectId), out displayValue, out linkValue);
        }

        /// <summary>
        /// Replaces <see cref="BusinessObjectIdPlaceholder"/> with <paramref name="businessObjectId"/>; the text is
        /// returned unchanged when the id is not known (0 or less).
        /// </summary>
        public static string ReplaceBusinessObjectIdPlaceholder(string rawCellText, int businessObjectId)
        {
            if (businessObjectId <= 0 || string.IsNullOrEmpty(rawCellText))
            {
                return rawCellText;
            }

            return rawCellText.Replace(BusinessObjectIdPlaceholder, businessObjectId.ToString(CultureInfo.InvariantCulture));
        }

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

        private static readonly Regex EmptyAnchorRegex = new Regex(@"^\s*<a\b[^>]*>\s*</a>\s*$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// True when the cell holds an anchor with no display text, i.e. a row without an attachment
        /// (e.g. "...{documentId:,businessObjectId:{businessObjectId}});'></a>"). The web report renders it blank.
        /// </summary>
        public static bool IsEmptyAttachmentAnchor(string rawCellText)
        {
            return !string.IsNullOrWhiteSpace(rawCellText) && EmptyAnchorRegex.IsMatch(rawCellText);
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
