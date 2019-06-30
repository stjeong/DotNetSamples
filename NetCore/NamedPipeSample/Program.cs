using System;
using System.Threading;
using System.IO.Pipes;
using System.Text;

namespace NamedPipeSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Thread t = new Thread(threadFunc);
            t.IsBackground = true;
            t.Start();

            Thread.Sleep(2000);

            while (true)
            {
                using (NamedPipeClientStream pipe = new NamedPipeClientStream("pipe_test1"))
                {
                    pipe.Connect(3000);
                    Console.WriteLine("[Client] Connected");
                    WriteString(pipe, "test2:test");
                    Console.WriteLine("[Client] Read: " + ReadString(pipe));
                }

                Console.WriteLine("Press any key to connect again... (q to quit)");
                if (Console.ReadKey().Key == ConsoleKey.Q)
                {
                    break;
                }

                Console.WriteLine();
            }
        }

        static void threadFunc()
        {
            Console.WriteLine("[Server] Thread is started.");

            using (NamedPipeServerStream pipe = new NamedPipeServerStream("pipe_test1",
                PipeDirection.InOut, NamedPipeServerStream.MaxAllowedServerInstances, PipeTransmissionMode.Byte, PipeOptions.None))
            {
                while (true)
                {
                    Console.WriteLine("[Server] Waiting for client connection...");
                    pipe.WaitForConnection();
                    Console.WriteLine("[Server] Client Connected.");

                    string txt = ReadString(pipe);
                    Console.WriteLine("[Server] Read: " + txt);
                    WriteString(pipe, "Echo: " + txt);
                    pipe.Disconnect();
                }
            }
        }

        static void WriteString(PipeStream pipe, string text)
        {
            byte[] buf = Encoding.UTF8.GetBytes(text);

            WriteInt32(pipe, buf.Length);
            pipe.Write(buf, 0, buf.Length);
        }

        static string ReadString(PipeStream pipe)
        {
            int sentBytes = ReadInt32(pipe);
            byte[] buf = new byte[sentBytes];

            pipe.Read(buf, 0, buf.Length);
            return Encoding.UTF8.GetString(buf, 0, buf.Length);
        }

        static void WriteInt32(PipeStream pipe, int value)
        {
            byte[] buf = BitConverter.GetBytes(value);
            pipe.Write(buf, 0, buf.Length);
        }

        static int ReadInt32(PipeStream pipe)
        {
            byte[] intBytes = new byte[4];
            pipe.Read(intBytes, 0, 4);

            return BitConverter.ToInt32(intBytes, 0);
        }
    }
}
