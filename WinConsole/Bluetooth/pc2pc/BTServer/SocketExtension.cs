using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;

static class SocketExtension
{
    public static void WriteString(this Socket socket, string text)
    {
        byte[] buf = Encoding.UTF8.GetBytes(text);

        WriteInt32(socket, buf.Length);
        socket.Send(buf, buf.Length, SocketFlags.None);
    }

    public static string ReadString(this Socket socket)
    {
        int sentBytes = ReadInt32(socket);
        byte[] buf = new byte[sentBytes];

        socket.Receive(buf, buf.Length, SocketFlags.None);
        return Encoding.UTF8.GetString(buf, 0, buf.Length);
    }

    public static void WriteInt32(this Socket socket, int value)
    {
        byte[] buf = BitConverter.GetBytes(value);
        socket.Send(buf, buf.Length, SocketFlags.None);
    }

    public static int ReadInt32(this Socket socket)
    {
        byte[] buf = new byte[4];
        socket.Receive(buf, buf.Length, SocketFlags.None);

        return BitConverter.ToInt32(buf, 0);
    }
}
