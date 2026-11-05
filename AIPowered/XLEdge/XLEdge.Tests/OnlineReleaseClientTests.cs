using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Contracts;
using XLEdge.Online;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// Stage 6: the shell's online release client - the endpoints a future orbitbootapp XLEdgeAddinController serves
    /// ({LoginUrl}/rest/secure/orbit-addons/manifests?type=xledge, .../versions/{folder}/manifest?type=xledge, .../zip?type=xledge), called with the login's Bearer token.
    /// </summary>
    [TestClass]
    public class OnlineReleaseClientTests
    {
        /// <summary>Answers by path; records each request's "METHOD path" and Authorization header.</summary>
        internal sealed class FakeServer : HttpMessageHandler
        {
            public readonly List<string> Requests = new List<string>();
            public readonly List<string> Authorization = new List<string>();
            public HttpStatusCode Status = HttpStatusCode.OK;
            public readonly Dictionary<string, byte[]> Bodies = new Dictionary<string, byte[]>(StringComparer.Ordinal);

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                string path = request.RequestUri.GetComponents(UriComponents.PathAndQuery, UriFormat.UriEscaped);
                lock (Requests)
                {
                    Requests.Add(request.Method + " " + request.RequestUri.GetComponents(UriComponents.PathAndQuery, UriFormat.UriEscaped));
                    Authorization.Add(request.Headers.Authorization?.ToString());
                }
                var response = new HttpResponseMessage(Status);
                if (Status == HttpStatusCode.OK)
                    response.Content = new ByteArrayContent(Bodies.TryGetValue(path, out byte[] body) ? body : new byte[0]);
                return Task.FromResult(response);
            }
        }

        internal const string Base = "/reporting/rest/secure/orbit-addons";
        internal const string AddinTypeQuery = "?type=xledge";
        internal static LoginInfo Login(string url = "https://host/reporting") => new LoginInfo { LoginUrl = url, LoginToken = "tok", LoginUserName = "u", IsLoggedIn = true };

        internal static string Entry(string version, string date, string folder) =>
            "{\"version\":\"" + version + "\",\"releaseDate\":\"" + date + "\",\"folderName\":\"" + folder + "\",\"fileName\":\"" + folder +
            ".zip\",\"checksum\":\"ab\",\"notes\":\"n\",\"mandatory\":false,\"contractVersion\":5}";

        [TestMethod]
        public void Client_ListsTheServersReleases_WithTheBearerToken()
        {
            var server = new FakeServer();
            server.Bodies[Base + "/manifests" + AddinTypeQuery] = System.Text.Encoding.UTF8.GetBytes("[" + Entry("11.1.3.0", "2026-11-03T10:00:00", "vA") + "," + Entry("11.1.4.0", "2026-11-10T10:00:00", "vB") + "]");
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), server))
            {
                var releases = client.GetReleasesAsync(CancellationToken.None).GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "vA", "vB" }, releases.Select(r => r.FolderName).ToArray());
                CollectionAssert.AreEqual(new[] { "11.1.3.0", "11.1.4.0" }, releases.Select(r => r.Version).ToArray());
            }
            CollectionAssert.AreEqual(new[] { "GET " + Base + "/manifests" + AddinTypeQuery }, server.Requests);
            CollectionAssert.AreEqual(new[] { "Bearer tok" }, server.Authorization);
        }

        [TestMethod]
        public void Client_ListsOnlyReleasesMarkedForRelease()
        {
            // "release": true / "TRUE" / blank / missing is shown; false, "false" or anything else is hidden.
            string With(string folder, string release) => Entry("11.1.3", "2026-11-03T10:00:00", folder).TrimEnd('}') + ",\"release\":" + release + "}";
            var server = new FakeServer();
            server.Bodies[Base + "/manifests" + AddinTypeQuery] = System.Text.Encoding.UTF8.GetBytes("[" + string.Join(",",
                With("vFalse", "false"), With("vTrue", "true"), Entry("11.1.3", "2026-11-03T10:00:00", "vMissing"), With("vUpper", "\"TRUE\""),
                With("vFalseText", "\"false\""), With("vNull", "null"), With("vBlank", "\"\""), With("vYes", "\"yes\""), With("vZero", "0")) + "]");
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), server))
            {
                var releases = client.GetReleasesAsync(CancellationToken.None).GetAwaiter().GetResult();
                CollectionAssert.AreEqual(new[] { "vTrue", "vMissing", "vUpper", "vNull", "vBlank" }, releases.Select(r => r.FolderName).ToArray());
            }
        }

        [TestMethod]
        public void Client_DownloadsManifestAndZip_OfTheEscapedFolder()
        {
            var server = new FakeServer();
            string folder = "v11.1.3.0_2026-11-03T10-00-00";
            server.Bodies[Base + "/versions/" + folder + "/manifest" + AddinTypeQuery] = System.Text.Encoding.UTF8.GetBytes("[" + Entry("11.1.3.0", "2026-11-03T10:00:00", folder) + "]");
            server.Bodies[Base + "/versions/" + folder + "/zip" + AddinTypeQuery] = new byte[] { 1, 2, 3 };
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), server))
            {
                var (manifest, zip) = client.DownloadAsync(folder, CancellationToken.None).GetAwaiter().GetResult();
                StringAssert.Contains(manifest, folder);
                CollectionAssert.AreEqual(new byte[] { 1, 2, 3 }, zip);
            }
            CollectionAssert.AreEqual(new[] { "GET " + Base + "/versions/" + folder + "/manifest" + AddinTypeQuery, "GET " + Base + "/versions/" + folder + "/zip" + AddinTypeQuery }, server.Requests);

            var odd = new FakeServer();
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), odd))
                try { client.DownloadAsync("a b/../c", CancellationToken.None).GetAwaiter().GetResult(); } catch (OnlineReleaseException) { }
            StringAssert.Contains(odd.Requests[0], "/versions/a%20b%2F..%2Fc/manifest", "the folder name is one escaped path segment");
        }

        private static OnlineReleaseException Fails(HttpStatusCode status)
        {
            var server = new FakeServer { Status = status };
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), server))
                return Assert.ThrowsExactly<OnlineReleaseException>(() => client.GetReleasesAsync(CancellationToken.None).GetAwaiter().GetResult());
        }

        [TestMethod]
        public void Client_Unauthorized_SaysTheSessionExpired()
        {
            var ex = Fails(HttpStatusCode.Unauthorized);
            Assert.AreEqual("Your XLEdge session has expired. Log in again, then check for updates again.", ex.Message);
            Assert.AreEqual(401, ex.StatusCode);
        }

        [TestMethod]
        public void Client_NotFound_SaysNotOnTheServer()
        {
            Assert.AreEqual("This XLEdge server does not offer XLEdge updates (yet).", Fails(HttpStatusCode.NotFound).Message,
                "the list itself missing: the server has no XLEdge release endpoints");
            var server = new FakeServer { Status = HttpStatusCode.NotFound };
            using (var client = new OnlineReleaseClient(Login(), new RecordingLogger(), server))
            {
                var ex = Assert.ThrowsExactly<OnlineReleaseException>(() => client.DownloadAsync("vA", CancellationToken.None).GetAwaiter().GetResult());
                Assert.AreEqual("That release is not on the server (any more).", ex.Message);
            }
        }

        [TestMethod]
        public void Client_UsesTheCoresTlsAndProxySettings()
        {
            // Review fix: the core's API calls force TLS 1.2/1.3 and no proxy (APIHelper); the shell's AppDomain does not inherit them.
            using (var handler = OnlineReleaseClient.CreateHandler())
            {
                Assert.AreEqual(System.Security.Authentication.SslProtocols.Tls12 | System.Security.Authentication.SslProtocols.Tls13, handler.SslProtocols);
                Assert.IsFalse(handler.UseProxy);
            }
        }

        [TestMethod]
        public void Client_ServerDown_SaysItCouldNotBeReached()
        {
            var listener = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0);
            listener.Start();
            int port = ((IPEndPoint)listener.LocalEndpoint).Port;
            listener.Stop();
            using (var client = new OnlineReleaseClient(Login("http://127.0.0.1:" + port + "/reporting"), new RecordingLogger()))
            {
                var ex = Assert.ThrowsExactly<OnlineReleaseException>(() => client.GetReleasesAsync(CancellationToken.None).GetAwaiter().GetResult());
                StringAssert.StartsWith(ex.Message, "The XLEdge server could not be reached: ");
                Assert.IsNull(ex.StatusCode);
            }
        }

        [TestMethod]
        public void Client_LoginUrlWithTrailingSlash_IsTrimmed()
        {
            var server = new FakeServer();
            server.Bodies[Base + "/manifests" + AddinTypeQuery] = System.Text.Encoding.UTF8.GetBytes("[" + Entry("11.1.3.0", "2026-11-03T10:00:00", "vA") + "]");
            using (var client = new OnlineReleaseClient(Login("https://host/reporting/"), new RecordingLogger(), server))
                client.GetReleasesAsync(CancellationToken.None).GetAwaiter().GetResult();
            CollectionAssert.AreEqual(new[] { "GET " + Base + "/manifests" + AddinTypeQuery }, server.Requests);
        }
    }

    [TestClass]
    public class OnlineReleaseClassifierTests
    {
        private static VersionInfo V(string version, string date, string folder) => new VersionInfo { Version = version, ReleaseDate = date, FolderName = folder, FileName = folder + ".zip", ContractVersion = 1 };

        [TestMethod]
        public void Classify_MarksLoadedDownloadedAndNew_NewestFirst_NewestNewTicked()
        {
            var server = new[] { V("11.1.3.0", "2026-11-03T10:00:00", "vA"), V("11.1.5.0", "2026-11-20T10:00:00", "vC"), V("11.1.4.0", "2026-11-10T10:00:00", "vB"), null, V("", "x", "vX") };
            var local = new List<ReleaseEntry> { new ReleaseEntry { Version = "11.1.4.0", ReleaseDate = "2026-11-10T10:00:00", FolderName = "vB" } };
            var rows = OnlineReleaseClassifier.Classify(server, local, "11.1.3.0", "2026-11-03T10:00:00");
            CollectionAssert.AreEqual(new[] { "vC", "vB", "vA" }, rows.Select(r => r.FolderName).ToArray(), "newest first; null/blank skipped");
            Assert.AreEqual("New", rows[0].StatusText);
            Assert.IsTrue(rows[0].IsSelectable && rows[0].IsChecked, "the newest new release is ticked");
            Assert.AreEqual("Already downloaded", rows[1].StatusText);
            Assert.IsFalse(rows[1].IsSelectable || rows[1].IsChecked);
            Assert.AreEqual("Currently loaded", rows[2].StatusText);
            Assert.IsFalse(rows[2].IsSelectable || rows[2].IsChecked);
        }

        [TestMethod]
        public void Classify_AReleaseThisShellCannotLoad_IsShownNotSelectable()
        {
            var tooNew = V("11.2.0.0", "2026-12-01T10:00:00", "vNext");
            tooNew.ContractVersion = ContractInfo.Version + 1;
            var rows = OnlineReleaseClassifier.Classify(new[] { tooNew, V("11.1.4.0", "2026-11-10T10:00:00", "v4") }, new List<ReleaseEntry>(), "11.1.3.0", "2026-11-03T10:00:00");
            Assert.AreEqual("vNext", rows[0].FolderName);
            Assert.AreEqual("Needs a newer XLEdge install", rows[0].StatusText);
            Assert.IsFalse(rows[0].IsSelectable || rows[0].IsChecked);
            Assert.IsTrue(rows[1].IsChecked, "the newest release this shell can load is ticked");
        }

        [TestMethod]
        public void Row_OnlySelectableRowsCanBeTicked()
        {
            var row = new OnlineReleaseRow { IsAlreadyDownloaded = true };
            row.IsChecked = true;
            Assert.IsFalse(row.IsChecked);
        }
    }
}
