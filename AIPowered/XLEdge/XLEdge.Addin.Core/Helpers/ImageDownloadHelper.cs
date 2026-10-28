using System;
using System.Net;
using System.Net.Http;
using System.Security.Authentication;
using System.Text.RegularExpressions;
using System.Threading;
using XLEdge.Utilities;

namespace XLEdge.Helpers
{
    /// <summary>
    /// Downloads a remote image for embedding into a report cell. Retries with backoff on
    /// throttling/server-busy responses (429/502/503/504, honoring a Retry-After header when present),
    /// enforces a minimum 1-second spacing between downloads, sets an appropriate Referer header, and
    /// falls back to scraping an image URL out of the response HTML when it isn't a direct image.
    /// </summary>
    public static class ImageDownloadHelper
    {
        private const int MaxAttempts = 3;
        private const int MinimumSpacingMs = 1000;
        private static readonly object ThrottleLock = new object();
        private static DateTime _lastDownloadUtc = DateTime.MinValue;

        public static bool TryDownloadImage(string url, string destinationPath)
        {
            if (string.IsNullOrWhiteSpace(url) || string.IsNullOrWhiteSpace(destinationPath))
            {
                return false;
            }

            return DownloadImageInternal(url, destinationPath, allowHtmlFallback: true);
        }

        private static bool DownloadImageInternal(string url, string destinationPath, bool allowHtmlFallback)
        {
            for (int attempt = 1; attempt <= MaxAttempts; attempt++)
            {
                try
                {
                    ThrottleImageDownload();

                    var handler = new HttpClientHandler
                    {
                        SslProtocols = SslProtocols.Tls12 | SslProtocols.Tls13,
                        ServerCertificateCustomValidationCallback = StrictCertificateValidator.Validate,
                        AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate,
                        AllowAutoRedirect = true
                    };

                    using var client = new HttpClient(handler) { Timeout = TimeSpan.FromSeconds(30) };
                    client.DefaultRequestHeaders.UserAgent.ParseAdd(
                        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36");
                    client.DefaultRequestHeaders.Accept.ParseAdd("image/avif,image/webp,image/apng,image/*,*/*;q=0.8");
                    client.DefaultRequestHeaders.AcceptLanguage.ParseAdd("en-US,en;q=0.9");

                    string referer = GetImageReferer(url);
                    if (!string.IsNullOrWhiteSpace(referer))
                    {
                        client.DefaultRequestHeaders.Referrer = new Uri(referer);
                    }

                    using HttpResponseMessage response = client.GetAsync(url).GetAwaiter().GetResult();
                    string contentType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;

                    if (response.IsSuccessStatusCode)
                    {
                        if (contentType.StartsWith("image", StringComparison.OrdinalIgnoreCase))
                        {
                            byte[] bytes = response.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                            System.IO.File.WriteAllBytes(destinationPath, bytes);
                            return true;
                        }

                        if (allowHtmlFallback && contentType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            string html = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                            string resolvedUrl = ExtractImageUrlFromHtml(html, url);
                            if (!string.IsNullOrWhiteSpace(resolvedUrl) && !string.Equals(resolvedUrl, url, StringComparison.OrdinalIgnoreCase))
                            {
                                return DownloadImageInternal(resolvedUrl, destinationPath, allowHtmlFallback: false);
                            }
                        }

                        LogUtility.LogWarn($"TryDownloadImage|Response for {url} was not an image (Content-Type: {contentType}).");
                        return false;
                    }

                    // Non-success response - try the HTML fallback on the error body too.
                    if (allowHtmlFallback && contentType.IndexOf("html", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        string html = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                        string resolvedUrl = ExtractImageUrlFromHtml(html, url);
                        if (!string.IsNullOrWhiteSpace(resolvedUrl) && !string.Equals(resolvedUrl, url, StringComparison.OrdinalIgnoreCase))
                        {
                            return DownloadImageInternal(resolvedUrl, destinationPath, allowHtmlFallback: false);
                        }
                    }

                    if (attempt < MaxAttempts && IsRetryableDownloadError(response.StatusCode))
                    {
                        Thread.Sleep(GetRetryDelayMilliseconds(response, attempt));
                        continue;
                    }

                    LogUtility.LogWarn($"TryDownloadImage|Server returned {(int)response.StatusCode} for {url}");
                    return false;
                }
                catch (Exception ex)
                {
                    LogUtility.LogException(ex, $"TryDownloadImage failed for {url}");
                    return false;
                }
            }

            return false;
        }

        private static void ThrottleImageDownload()
        {
            lock (ThrottleLock)
            {
                int elapsedMs = 0;
                if (_lastDownloadUtc != DateTime.MinValue)
                {
                    elapsedMs = (int)(DateTime.UtcNow - _lastDownloadUtc).TotalMilliseconds;
                }

                if (elapsedMs < MinimumSpacingMs)
                {
                    Thread.Sleep(MinimumSpacingMs - elapsedMs);
                }

                _lastDownloadUtc = DateTime.UtcNow;
            }
        }

        private static bool IsRetryableDownloadError(HttpStatusCode statusCode)
        {
            return (int)statusCode == 429 ||
                   statusCode == HttpStatusCode.ServiceUnavailable ||
                   statusCode == HttpStatusCode.BadGateway ||
                   statusCode == HttpStatusCode.GatewayTimeout;
        }

        private static int GetRetryDelayMilliseconds(HttpResponseMessage response, int attempt)
        {
            if (response.Headers.RetryAfter?.Delta.HasValue == true)
            {
                int seconds = (int)response.Headers.RetryAfter.Delta.Value.TotalSeconds;
                if (seconds > 0)
                {
                    return Math.Min(seconds * 1000, 10000);
                }
            }

            return Math.Min((int)Math.Pow(2, attempt) * 1000, 10000);
        }

        /// <summary>Wikipedia-hosted images need a Wikipedia referer or the CDN rejects the request;
        /// everything else gets its own origin as a reasonable default.</summary>
        private static string GetImageReferer(string imageUrl)
        {
            try
            {
                var uri = new Uri(imageUrl);

                if (string.Equals(uri.Host, "upload.wikimedia.org", StringComparison.OrdinalIgnoreCase))
                {
                    return "https://en.wikipedia.org/";
                }

                return uri.GetLeftPart(UriPartial.Authority) + "/";
            }
            catch (Exception ex)
            {
                // Safe to ignore: best-effort referer derivation from an image URL; falls back to no
                // referer header being sent, which most CDNs tolerate.
                LogUtility.LogDebug($"{nameof(GetImageReferer)}: failed to derive referer from '{imageUrl}' - {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>Scrapes a plausible direct image URL out of an HTML page returned instead of the
        /// expected image (og:image / twitter:image meta tags, Wikipedia's file-page
        /// &lt;img class="...mw-file-element..."&gt;, or a link[rel=image_src]).</summary>
        private static string ExtractImageUrlFromHtml(string html, string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(html))
            {
                return string.Empty;
            }

            const string dq = "\"";
            string[] patterns =
            {
                $"<meta[^>]+property\\s*=\\s*['{dq}]og:image['{dq}][^>]+content\\s*=\\s*['{dq}](?<url>[^'{dq} >]+)",
                $"<meta[^>]+name\\s*=\\s*['{dq}]twitter:image['{dq}][^>]+content\\s*=\\s*['{dq}](?<url>[^'{dq} >]+)",
                $"<img[^>]+class\\s*=\\s*['{dq}][^'{dq}]*mw-file-element[^'{dq}]*['{dq}][^>]+src\\s*=\\s*['{dq}](?<url>[^'{dq} >]+)",
                $"<link[^>]+rel\\s*=\\s*['{dq}]image_src['{dq}][^>]+href\\s*=\\s*['{dq}](?<url>[^'{dq} >]+)"
            };

            foreach (string pattern in patterns)
            {
                Match match = Regex.Match(html, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (match.Success)
                {
                    return NormalizeResolvedImageUrl(match.Groups["url"].Value, baseUrl);
                }
            }

            return string.Empty;
        }

        private static string NormalizeResolvedImageUrl(string candidateUrl, string baseUrl)
        {
            if (string.IsNullOrWhiteSpace(candidateUrl))
            {
                return string.Empty;
            }

            candidateUrl = WebUtility.HtmlDecode(candidateUrl).Trim();

            if (candidateUrl.StartsWith("//"))
            {
                return "https:" + candidateUrl;
            }

            if (candidateUrl.StartsWith("/"))
            {
                try
                {
                    var baseUri = new Uri(baseUrl);
                    return new Uri(baseUri, candidateUrl).ToString();
                }
                catch (Exception ex)
                {
                    // Safe to ignore: best-effort relative-URL resolution against baseUrl; falls back
                    // to returning the candidate URL as-is.
                    LogUtility.LogDebug($"{nameof(NormalizeResolvedImageUrl)}: failed to resolve relative URL '{candidateUrl}' against base '{baseUrl}' - {ex.Message}");
                    return candidateUrl;
                }
            }

            return candidateUrl;
        }
    }
}
