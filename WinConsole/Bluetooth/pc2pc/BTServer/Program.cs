using InTheHand.Net;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;
using System;
using System.Net.Sockets;

namespace BTServer
{
    class Program
    {
        static void Main(string[] args)
        {
            Guid serviceUUID = BluetoothService.TcpProtocol; // new Guid("00000004-0000-1000-8000-00805F9B34FB");

            BluetoothEndPoint bep = new BluetoothEndPoint(BluetoothAddress.None, BluetoothService.TcpProtocol);
            BluetoothListener blueListener = new BluetoothListener(bep);

            blueListener.Start();

            Console.WriteLine("accepting...");

            while (true)
            {
                Socket socket = blueListener.AcceptSocket();

                string text = socket.ReadString();
                Console.WriteLine("socket accepted: " + text);
                socket.WriteString("Echo: " + text);
                socket.Close();
                Console.WriteLine("Closed");
            }
        }
    }
}
