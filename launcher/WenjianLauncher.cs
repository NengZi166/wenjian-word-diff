using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Reflection;
using System.Text;
using System.Threading;

[assembly: AssemblyTitle("Wenjian Launcher")]
[assembly: AssemblyDescription("Local-only launcher for Wenjian Word document comparison")]
[assembly: AssemblyCompany("Wenjian contributors")]
[assembly: AssemblyProduct("Wenjian")]
[assembly: AssemblyCopyright("Copyright Wenjian contributors")]
[assembly: AssemblyVersion("0.1.2.0")]
[assembly: AssemblyFileVersion("0.1.2.0")]

internal sealed class ClientContext
{
    public ClientContext(TcpClient client, string siteRoot, string indexPath)
    {
        Client = client;
        SiteRoot = siteRoot;
        IndexPath = indexPath;
    }

    public TcpClient Client { get; private set; }
    public string SiteRoot { get; private set; }
    public string IndexPath { get; private set; }
}

internal static class Program
{
    private static readonly Dictionary<string, string> MimeTypes =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".html", "text/html; charset=utf-8" },
            { ".js", "text/javascript; charset=utf-8" },
            { ".css", "text/css; charset=utf-8" },
            { ".json", "application/json; charset=utf-8" },
            { ".wasm", "application/wasm" },
            { ".dll", "application/octet-stream" },
            { ".dat", "application/octet-stream" },
            { ".blat", "application/octet-stream" },
            { ".br", "application/octet-stream" },
            { ".woff", "font/woff" },
            { ".woff2", "font/woff2" },
            { ".ttf", "font/ttf" },
            { ".svg", "image/svg+xml" },
            { ".png", "image/png" },
            { ".ico", "image/x-icon" }
        };

    private static volatile bool running = true;
    private static TcpListener listener;

    private static int Main(string[] args)
    {
        try
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            Console.Title = "Wenjian - Word document comparison";
        }
        catch
        {
        }

        string siteRoot = Path.GetFullPath(AppDomain.CurrentDomain.BaseDirectory);
        string indexPath = Path.Combine(siteRoot, "index.html");

        if (!File.Exists(indexPath))
        {
            ShowFatal("Offline site files are incomplete. Extract the whole ZIP and keep all files together.");
            return 1;
        }

        int port;
        listener = StartListener(out port);
        if (listener == null)
        {
            ShowFatal("Could not use a local port from 8765 to 8775. Close another Wenjian window and retry.");
            return 2;
        }

        Console.CancelKeyPress += delegate(object sender, ConsoleCancelEventArgs eventArgs)
        {
            eventArgs.Cancel = true;
            running = false;
            try { listener.Stop(); } catch { }
        };

        string url = "http://127.0.0.1:" + port + "/";
        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine("Wenjian is running.");
        Console.ResetColor();
        Console.WriteLine("Browser address: " + url);
        Console.WriteLine("Keep this window open. Closing it stops Wenjian.");
        Console.WriteLine();

        bool noBrowser = HasArgument(args, "--no-browser");
        if (!noBrowser)
        {
            TryOpenBrowser(url);
        }

        while (running)
        {
            try
            {
                TcpClient client = listener.AcceptTcpClient();
                ThreadPool.QueueUserWorkItem(HandleClient, new ClientContext(client, siteRoot, indexPath));
            }
            catch (SocketException)
            {
                if (running)
                {
                    Console.WriteLine("The local server stopped unexpectedly.");
                    return 3;
                }
            }
            catch (ObjectDisposedException)
            {
                break;
            }
        }

        return 0;
    }

    private static TcpListener StartListener(out int selectedPort)
    {
        for (int port = 8765; port <= 8775; port++)
        {
            TcpListener candidate = null;
            try
            {
                candidate = new TcpListener(IPAddress.Loopback, port);
                candidate.Start();
                selectedPort = port;
                return candidate;
            }
            catch
            {
                if (candidate != null)
                {
                    try { candidate.Stop(); } catch { }
                }
            }
        }

        selectedPort = 0;
        return null;
    }

    private static void HandleClient(object state)
    {
        ClientContext context = (ClientContext)state;

        using (TcpClient client = context.Client)
        {
            try
            {
                client.ReceiveTimeout = 10000;
                client.SendTimeout = 30000;

                using (NetworkStream stream = client.GetStream())
                {
                    byte[] requestBuffer = new byte[32768];
                    int requestLength = ReadHeaders(stream, requestBuffer);
                    if (requestLength <= 0)
                    {
                        return;
                    }

                    string request = Encoding.ASCII.GetString(requestBuffer, 0, requestLength);
                    int firstLineEnd = request.IndexOf("\r\n", StringComparison.Ordinal);
                    string requestLine = firstLineEnd >= 0 ? request.Substring(0, firstLineEnd) : request;
                    string[] requestParts = requestLine.Split(' ');

                    if (requestParts.Length < 2)
                    {
                        WriteTextResponse(stream, 400, "Bad Request", "Bad request.", false);
                        return;
                    }

                    string method = requestParts[0].ToUpperInvariant();
                    bool headOnly = method == "HEAD";
                    if (method != "GET" && !headOnly)
                    {
                        WriteTextResponse(stream, 405, "Method Not Allowed", "Only GET and HEAD are supported.", headOnly);
                        return;
                    }

                    int statusCode;
                    string filePath = ResolveFile(context.SiteRoot, context.IndexPath, requestParts[1], out statusCode);
                    if (filePath == null)
                    {
                        if (statusCode == 403)
                        {
                            WriteTextResponse(stream, 403, "Forbidden", "Forbidden.", headOnly);
                        }
                        else if (statusCode == 400)
                        {
                            WriteTextResponse(stream, 400, "Bad Request", "Bad request.", headOnly);
                        }
                        else
                        {
                            WriteTextResponse(stream, 404, "Not Found", "Not found.", headOnly);
                        }
                        return;
                    }

                    byte[] body = File.ReadAllBytes(filePath);
                    string extension = Path.GetExtension(filePath);
                    string contentType;
                    if (!MimeTypes.TryGetValue(extension, out contentType))
                    {
                        contentType = "application/octet-stream";
                    }

                    WriteResponse(stream, 200, "OK", contentType, body, headOnly);
                }
            }
            catch
            {
            }
        }
    }

    private static int ReadHeaders(NetworkStream stream, byte[] buffer)
    {
        int total = 0;
        while (total < buffer.Length)
        {
            int read = stream.Read(buffer, total, buffer.Length - total);
            if (read <= 0)
            {
                return total;
            }

            total += read;
            for (int i = Math.Max(3, total - read - 3); i < total; i++)
            {
                if (i >= 3 && buffer[i - 3] == 13 && buffer[i - 2] == 10 && buffer[i - 1] == 13 && buffer[i] == 10)
                {
                    return total;
                }
            }
        }

        return total;
    }

    private static string ResolveFile(string siteRoot, string indexPath, string requestTarget, out int statusCode)
    {
        statusCode = 200;
        try
        {
            Uri absoluteUri;
            if (Uri.TryCreate(requestTarget, UriKind.Absolute, out absoluteUri))
            {
                requestTarget = absoluteUri.PathAndQuery;
            }

            int queryIndex = requestTarget.IndexOf('?');
            if (queryIndex >= 0)
            {
                requestTarget = requestTarget.Substring(0, queryIndex);
            }

            string relativePath = Uri.UnescapeDataString(requestTarget)
                .Replace('/', Path.DirectorySeparatorChar)
                .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            if (relativePath.IndexOf('\0') >= 0)
            {
                statusCode = 400;
                return null;
            }

            if (relativePath.Length == 0)
            {
                return indexPath;
            }

            string candidate = Path.GetFullPath(Path.Combine(siteRoot, relativePath));
            string rootedPrefix = siteRoot.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? siteRoot
                : siteRoot + Path.DirectorySeparatorChar;

            if (!candidate.StartsWith(rootedPrefix, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(candidate, siteRoot.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
            {
                statusCode = 403;
                return null;
            }

            if (Directory.Exists(candidate))
            {
                candidate = Path.Combine(candidate, "index.html");
            }

            if (File.Exists(candidate))
            {
                return candidate;
            }

            if (Path.GetExtension(relativePath).Length == 0)
            {
                return indexPath;
            }

            statusCode = 404;
            return null;
        }
        catch
        {
            statusCode = 400;
            return null;
        }
    }

    private static void WriteTextResponse(NetworkStream stream, int statusCode, string statusText, string message, bool headOnly)
    {
        byte[] body = Encoding.UTF8.GetBytes(message);
        WriteResponse(stream, statusCode, statusText, "text/plain; charset=utf-8", body, headOnly);
    }

    private static void WriteResponse(NetworkStream stream, int statusCode, string statusText, string contentType, byte[] body, bool headOnly)
    {
        string header = string.Format(
            "HTTP/1.1 {0} {1}\r\nContent-Type: {2}\r\nContent-Length: {3}\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nConnection: close\r\n\r\n",
            statusCode,
            statusText,
            contentType,
            body.Length);

        byte[] headerBytes = Encoding.ASCII.GetBytes(header);
        stream.Write(headerBytes, 0, headerBytes.Length);
        if (!headOnly && body.Length > 0)
        {
            stream.Write(body, 0, body.Length);
        }
        stream.Flush();
    }

    private static bool HasArgument(string[] args, string expected)
    {
        foreach (string argument in args)
        {
            if (string.Equals(argument, expected, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        return false;
    }

    private static void TryOpenBrowser(string url)
    {
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = url;
            startInfo.UseShellExecute = true;
            Process.Start(startInfo);
        }
        catch
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("The browser could not open automatically.");
            Console.WriteLine("Open this address manually: " + url);
            Console.ResetColor();
            Console.WriteLine();
        }
    }

    private static void ShowFatal(string message)
    {
        try { Console.ForegroundColor = ConsoleColor.Red; } catch { }
        Console.WriteLine(message);
        try { Console.ResetColor(); } catch { }
        Console.WriteLine("Press any key to close.");
        try { Console.ReadKey(true); } catch { }
    }
}
