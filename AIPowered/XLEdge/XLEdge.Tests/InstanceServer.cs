using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace XLEdge.Tests
{
    /// <summary>
    /// A stand-in XLEdge instance on localhost for the pane (TcpListener: needs no URL reservation): its pages, the requests it got.
    /// Every page sets document.title to "ready" once loaded, then runs <see cref="AfterReady"/> (script) - how a test drives the
    /// title commands and the login. /web/secure/get-broadcast-msg answers "[]"; /web/secure/applogout a plain page.
    /// </summary>
    internal sealed class InstanceServer : IDisposable
    {
        private readonly TcpListener _listener = new TcpListener(IPAddress.Loopback, 0);
        private readonly List<string> _requests = new List<string>();

        public InstanceServer()
        {
            _listener.Start();
            Task.Run(Serve);
        }

        /// <summary>The instance's address, as a login URL (no trailing slash).</summary>
        public string Url => $"http://localhost:{((IPEndPoint)_listener.LocalEndpoint).Port}";

        /// <summary>Script the page runs after its "ready" title (null: nothing).</summary>
        public string AfterReady { get; set; }

        /// <summary>A Set-Cookie value sent with the page (null: none).</summary>
        public string Cookie { get; set; }

        /// <summary>Each request's first line and headers.</summary>
        public string[] Requests { get { lock (_requests) return _requests.ToArray(); } }

        private void Serve()
        {
            try
            {
                while (true)
                {
                    TcpClient client = _listener.AcceptTcpClient();
                    Task.Run(() => Answer(client));
                }
            }
            catch (Exception)
            {
                // Stopped.
            }
        }

        private void Answer(TcpClient client)
        {
            try
            {
                using (client)
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
                    string request = head.ToString();
                    lock (_requests) _requests.Add(request);

                    string path = request.Split(' ').Length > 1 ? request.Split(' ')[1] : "/";
                    string contentType = "text/html", payload, extra = "";
                    if (path.StartsWith("/web/secure/get-broadcast-msg", StringComparison.Ordinal))
                    {
                        contentType = "application/json";
                        payload = "[]";
                    }
                    else if (path.StartsWith("/favicon", StringComparison.Ordinal))
                    {
                        payload = "";
                    }
                    else
                    {
                        payload = "<!doctype html><html><head><title>loading</title></head><body>instance<script>" +
                                  "document.title='ready';" + (AfterReady == null ? "" : "setTimeout(function(){" + AfterReady + "},100);") +
                                  "</script></body></html>";
                        if (Cookie != null) extra = "Set-Cookie: " + Cookie + "\r\n";
                    }
                    byte[] bytes = Encoding.UTF8.GetBytes(payload);
                    string reply = $"HTTP/1.1 200 OK\r\nContent-Type: {contentType}; charset=utf-8\r\nContent-Length: {bytes.Length}\r\n{extra}Cache-Control: no-store\r\nConnection: close\r\n\r\n";
                    byte[] headBytes = Encoding.ASCII.GetBytes(reply);
                    stream.Write(headBytes, 0, headBytes.Length);
                    stream.Write(bytes, 0, bytes.Length);
                }
            }
            catch (Exception)
            {
                // A client that went away.
            }
        }

        public void Dispose() => _listener.Stop();
    }
}
