using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>
    /// The ported ApiHelper in the core's AppDomain: ServicePointManager is per AppDomain, so XLEdge 11.1.2's settings (made in
    /// ApiHelper's static constructor) must take effect in the core's domain; and a request reaches the server with the login
    /// token as a Bearer header, as today.
    /// </summary>
    [TestClass]
    public class CoreApiTests
    {
        /// <summary>A one-request-at-a-time HTTP server on localhost (TcpListener: needs no URL reservation).</summary>
        private sealed class TinyServer : IDisposable
        {
            private readonly TcpListener _listener = new TcpListener(IPAddress.Loopback, 0);
            private readonly int _status;
            private readonly string _body;
            public readonly List<string> Requests = new List<string>();

            public TinyServer(int status, string body)
            {
                _status = status;
                _body = body;
                _listener.Start();
                Task.Run(Serve);
            }

            public string Url => $"http://localhost:{((IPEndPoint)_listener.LocalEndpoint).Port}/rest/secure/test";

            private void Serve()
            {
                try
                {
                    while (true)
                    {
                        using (TcpClient client = _listener.AcceptTcpClient())
                        using (NetworkStream stream = client.GetStream())
                        {
                            var reader = new StreamReader(stream, Encoding.ASCII, false, 1024, leaveOpen: true);
                            var head = new StringBuilder();
                            string line;
                            int length = 0;
                            while (!string.IsNullOrEmpty(line = reader.ReadLine()))
                            {
                                head.AppendLine(line);
                                if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase)) length = int.Parse(line.Substring(15).Trim());
                            }
                            var body = new char[length];
                            int read = 0;
                            while (read < length) read += reader.Read(body, read, length - read);
                            lock (Requests) Requests.Add(head + "\n" + new string(body));

                            byte[] payload = Encoding.UTF8.GetBytes(_body);
                            string reply = $"HTTP/1.1 {_status} X\r\nContent-Type: application/json\r\nContent-Length: {payload.Length}\r\nConnection: close\r\n\r\n";
                            byte[] headBytes = Encoding.ASCII.GetBytes(reply);
                            stream.Write(headBytes, 0, headBytes.Length);
                            stream.Write(payload, 0, payload.Length);
                        }
                    }
                }
                catch (Exception)
                {
                    // Stopped.
                }
            }

            public void Dispose() => _listener.Stop();
        }

        private static (AppDomain, CoreProbe, TempFolder) Core()
        {
            var t = new TempFolder();
            var paths = new PathProvider(t.Combine("inst"), t.Combine("data"));
            paths.Ensure();
            var (domain, probe) = CoreProbe.Create("CoreApi");
            probe.CallStatic("XLEdge.Addin.Core.Infrastructure.CoreServices", "Initialize", new FakeContext(paths) { Logger = new RecordingLogger() });
            return (domain, probe, t);
        }

        [TestMethod]
        public void ApiHelper_MakesTodaysConnectionSettings_InTheCoresDomain()
        {
            var (domain, probe, t) = Core();
            try
            {
                Assert.AreEqual("limit=100 expect100=False idle=5000 protocols=Tls12, Tls13", probe.ServicePointSettingsAfterApiHelper());
            }
            finally
            {
                AppDomain.Unload(domain);
                t.Dispose();
            }
        }

        [TestMethod]
        public void ServerApi_SendsTheBodyWithTheBearerToken_AndReturnsTheResponse()
        {
            var (domain, probe, t) = Core();
            using (var server = new TinyServer(200, "{\"status\":\"ok\",\"rows\":3}"))
            {
                try
                {
                    string result = probe.CallServerApi(server.Url, "{\"reportId\":\"42\"}", "tok-123");
                    Assert.AreEqual("ok:{\"status\":\"ok\",\"rows\":3}", result);
                    string request;
                    lock (server.Requests) request = server.Requests[0];
                    StringAssert.StartsWith(request, "POST /rest/secure/test HTTP/1.1");
                    StringAssert.Contains(request, "Authorization: Bearer tok-123");
                    StringAssert.Contains(request, "{\"reportId\":\"42\"}");
                }
                finally
                {
                    AppDomain.Unload(domain);
                    t.Dispose();
                }
            }
        }

        [TestMethod]
        public void ServerApi_HttpError_IsAnApiRequestException()
        {
            var (domain, probe, t) = Core();
            using (var server = new TinyServer(404, "{\"message\":\"Report not found\"}"))
            {
                try
                {
                    string result = probe.CallServerApi(server.Url, "{}", "tok");
                    StringAssert.StartsWith(result, "throws XLEdge.Helpers.ApiRequestException:");
                    StringAssert.Contains(result, "Report not found");
                }
                finally
                {
                    AppDomain.Unload(domain);
                    t.Dispose();
                }
            }
        }
    }
}
