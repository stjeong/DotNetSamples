using System;
using System.Threading;
using System.Net.Sockets;
using System.IO;
using System.Runtime.InteropServices;

class Program
{
    static string _socketName = "testd.sock";
    static string _linuxTempPath = "/var/tmp";
    static string _winTempPath = "%TEMP%";

    static string _socketPath;

    // C# - .NET Core Unix Domain Socket 사용 예제
    // http://www.sysnet.pe.kr/2/0/11963

    // C# - .NET Core 2.0 이하의 Unix Domain Socket 사용 시 System.IndexOutOfRangeException 오류
    // http://www.sysnet.pe.kr/2/0/11999


    static void Main(string[] args)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            string tempPath = Path.Combine(_winTempPath, _socketName);
            _socketPath = Environment.ExpandEnvironmentVariables(tempPath);
        }
        else
        {
            _socketPath = Path.Combine(_linuxTempPath, _socketName);
        }

        Thread t = new Thread(threadFunc);
        t.IsBackground = true;
        t.Start();

        Thread.Sleep(2000);

        while (true)
        {
            ConnectAsClient();
            string txt = Console.ReadLine();
            if (txt == "q")
            {
                return;
            }
        }
    }

    private static void ConnectAsClient()
    {
        var socket = new Socket(AddressFamily.Unix, SocketType.Stream, ProtocolType.IP);
        var unixEp = new UnixDomainSocketEndPoint(_socketPath);

        // Under .NET Core 2.1, you can use user-defined UnixEndPoint type
        // var unixEp = new UnixEndPoint(_socketPath);

        socket.Connect(unixEp);
        Console.WriteLine("[Client] Conencted");
        socket.Close();
        Console.WriteLine("[Client] Closed");
    }
    static void threadFunc()
    {
        Console.WriteLine("[Server] Thread is started.");

        if (File.Exists(_socketPath) == true)
        {
            File.Delete(_socketPath);
        }

        try
        {
            using (var socket = new Socket(AddressFamily.Unix, SocketType.Stream, ProtocolType.IP))
            {
                var unixEp = new UnixDomainSocketEndPoint(_socketPath);

                // Under .NET Core 2.1, you can use user-defined UnixEndPoint type
                // var unixEp = new UnixEndPoint(_socketPath);

                socket.Bind(unixEp);
                socket.Listen(5);
                while (true)
                {
                    using (Socket clntSocket = socket.Accept())
                    {
                        Console.WriteLine("[Server] ClientConencted");
                    }
                    Console.WriteLine("[Server] ClientClosed");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
    }
}
