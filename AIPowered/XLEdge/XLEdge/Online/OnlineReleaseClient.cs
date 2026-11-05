using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using XLEdge.Contracts;
using XLEdge.Shared;

namespace XLEdge.Online
{
    /// <summary>Where the Online reload gets XLEdge core releases (the server, or a test stand-in).</summary>
    internal interface IOnlineReleaseSource : IDisposable
    {
        /// <summary>The server's release index (manifest entries).</summary>
        Task<IReadOnlyList<VersionInfo>> GetReleasesAsync(CancellationToken ct);

        /// <summary>One release's manifest.json text and zip bytes.</summary>
        Task<(string ManifestJson, byte[] Zip)> DownloadAsync(string folderName, CancellationToken ct);
    }

    /// <summary>
    /// The ORBIT server's XLEdge add-in release endpoints (orbitbootapp XLEdgeAddinController - spec
    /// OISR-22460), called with the user's login: {LoginUrl}/rest/secure/orbit-addons/manifests?type=xledge,
    /// .../versions/{folderName}/manifest?type=xledge and .../zip?type=xledge, "Authorization: Bearer {LoginToken}". The shell's own HttpClient: the core's
    /// API helper lives in another AppDomain.
    /// </summary>
    internal sealed class OnlineReleaseClient : IOnlineReleaseSource
    {
        internal const string SessionExpired = "Your XLEdge session has expired. Log in again, then check for updates again.";
        internal const string NotOnServer = "That release is not on the server (any more).";
        internal const string NoUpdatesHere = "This XLEdge server does not offer XLEdge updates (yet).";
        internal const string Unreachable = "The XLEdge server could not be reached: ";
        private const string AddinTypeQuery = "?type=xledge";

        private readonly HttpClient _http;
        private readonly string _base;
        private readonly ILogger _logger;

        public OnlineReleaseClient(LoginInfo login, ILogger logger, HttpMessageHandler handler = null)
        {
            if (login == null || string.IsNullOrWhiteSpace(login.LoginUrl)) throw new ArgumentException("No login URL.", nameof(login));
            _logger = logger;
            _base = login.LoginUrl.Trim().TrimEnd('/') + "/rest/secure/orbit-addons";
            _http = handler == null ? new HttpClient(CreateHandler(), disposeHandler: true) : new HttpClient(handler, disposeHandler: false);
            _http.Timeout = TimeSpan.FromMinutes(5);
            if (!string.IsNullOrWhiteSpace(login.LoginToken))
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", login.LoginToken);
        }

        /// <summary>
        /// The core's HTTP settings for ORBIT calls (APIHelper: TLS 1.2/1.3, no proxy) - they are per AppDomain, so the shell sets
        /// them on its own handler. Certificates are validated as Windows does by default.
        /// </summary>
        internal static HttpClientHandler CreateHandler() => new HttpClientHandler
        {
            SslProtocols = System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13,
            UseProxy = false,
        };

        public async Task<IReadOnlyList<VersionInfo>> GetReleasesAsync(CancellationToken ct)
        {
            byte[] body = await GetAsync(_base + "/manifests" + AddinTypeQuery, ct, NoUpdatesHere).ConfigureAwait(false);
            VersionParseResult parsed = new VersionParser(_logger).ParseVersionJson(System.Text.Encoding.UTF8.GetString(body));
            if (!parsed.Success) throw new OnlineReleaseException("The server's release list could not be read: " + parsed.ErrorMessage, null);
            return parsed.AllVersions;
        }

        public async Task<(string ManifestJson, byte[] Zip)> DownloadAsync(string folderName, CancellationToken ct)
        {
            string release = _base + "/versions/" + Uri.EscapeDataString(folderName ?? string.Empty);
            byte[] manifest = await GetAsync(release + "/manifest" + AddinTypeQuery, ct).ConfigureAwait(false);
            byte[] zip = await GetAsync(release + "/zip" + AddinTypeQuery, ct).ConfigureAwait(false);
            return (System.Text.Encoding.UTF8.GetString(manifest), zip);
        }

        private async Task<byte[]> GetAsync(string url, CancellationToken ct, string notFound = NotOnServer)
        {
            HttpResponseMessage response;
            try { response = await _http.GetAsync(url, ct).ConfigureAwait(false); }
            catch (OperationCanceledException) when (ct.IsCancellationRequested) { throw; }
            catch (Exception ex)
            {
                string reason = ex is TaskCanceledException ? "no answer within 5 minutes" : (ex.InnerException ?? ex).Message;
                _logger?.LogWarn($"OnlineReleaseClient: {url} - {reason}");
                throw new OnlineReleaseException(Unreachable + reason, null);
            }
            using (response)
            {
                int status = (int)response.StatusCode;
                if (response.IsSuccessStatusCode) return await response.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                _logger?.LogWarn($"OnlineReleaseClient: {url} - HTTP {status}");
                if (status == 401) throw new OnlineReleaseException(SessionExpired, status);
                if (status == 404) throw new OnlineReleaseException(notFound, status);
                throw new OnlineReleaseException(Unreachable + "HTTP " + status, status);
            }
        }

        public void Dispose() => _http.Dispose();
    }
}
