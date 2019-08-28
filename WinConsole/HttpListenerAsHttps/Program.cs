using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace HttpListenerAsHttps
{
    // C# - HttpListener를 이용한 HTTPS 통신 방법
    // http://www.sysnet.pe.kr/2/0/12012

    class Program
    {
        // netsh http add urlacl url=http://+:80/MyTemp/ user=testuser
        // netsh http delete urlacl url=http://+:80/MyTemp/

        // netsh http add urlacl url=https://+:50443/MyTemp/ user=testuser
        // netsh http delete urlacl url=https://+:50443/MyTemp/

        // makecert -n "CN=localhost" -r -sky exchange -sv mycert.pvk mycert.cer
        // pvk2pfx -pvk mycert.pvk -spc mycert.cer -pfx mycert.pfx
        // Register mycert.pfx at Local Computer / Trusted Root Certification Authorities / Certificates
        // Register mycert.pfx at Local Computer / Personal / Certificates

        // netsh http add sslcert ipport=0.0.0.0:50443 certhash=[...[cert_thumbprint]...] appid={...[arbitrary_guid]...}

        static void Main(string[] args)
        {
            HttpListener httpServer = new HttpListener();
            httpServer.Prefixes.Add("http://+:80/MyTemp/");
            httpServer.Prefixes.Add("https://+:50443/MyTemp/");

            httpServer.AuthenticationSchemes = AuthenticationSchemes.Anonymous;
            httpServer.Start();

            httpServer.BeginGetContext(ProcessRequest, httpServer);

            Console.ReadLine();
        }

        static void ProcessRequest(IAsyncResult ar)
        {
            HttpListener listener = ar.AsyncState as HttpListener;
            HttpListenerContext ctx = listener.EndGetContext(ar);

            HttpListenerRequest req = ctx.Request;
            HttpListenerResponse resp = ctx.Response;

            StringBuilder sb = new StringBuilder();
            sb.Append("<html><body><h1>" + "test" + "</h1>");
            sb.Append("</body></html>");

            byte[] buf = Encoding.UTF8.GetBytes(sb.ToString());
            resp.OutputStream.Write(buf, 0, buf.Length);
            resp.OutputStream.Close();

            listener.BeginGetContext(ProcessRequest, listener);
        }
    }
}
