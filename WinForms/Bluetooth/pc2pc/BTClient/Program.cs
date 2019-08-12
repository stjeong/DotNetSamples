using InTheHand.Net;
using InTheHand.Net.Bluetooth;
using InTheHand.Net.Sockets;
using System;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Text;

namespace ConsoleApp1
{
    class Program
    {
        static Dictionary<string, BluetoothDeviceInfo> _btList = new Dictionary<string, BluetoothDeviceInfo>();
        static string _targetMachineName = "";

        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                _targetMachineName = args[0];
            }

            // http://blogs.microsoft.co.il/shair/2009/06/21/working-with-bluetooth-devices-using-c-part-1/
            //{
            //    BluetoothClient bc = new BluetoothClient();
            //    foreach (BluetoothDeviceInfo bdi in bc.DiscoverDevicesInRange())
            //    {
            //        Console.WriteLine($"{bdi.DeviceName}");
            //    }

            //    Console.WriteLine();

            //    foreach (BluetoothDeviceInfo bdi in bc.DiscoverDevices())
            //    {
            //        Console.WriteLine($"{bdi.DeviceName}");
            //    }
            //}

            // https://stackoverflow.com/questions/42701793/how-to-programatically-pair-a-bluetooth-device
            {
                BluetoothClient bluetoothClient;
                BluetoothComponent bluetoothComponent;

                bluetoothClient = new BluetoothClient();
                bluetoothComponent = new BluetoothComponent(bluetoothClient);

                bluetoothComponent.DiscoverDevicesProgress += bluetoothComponent_DiscoverDevicesProgress;
                bluetoothComponent.DiscoverDevicesComplete += bluetoothComponent_DiscoverDevicesComplete;

                bluetoothComponent.DiscoverDevicesAsync(255, false, true, true, false, null);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();

            // https://stackoverflow.com/questions/22346166/error-10049-while-connecting-bluetooth-device-with-32feet?rq=1
            // if (false)
            {
                //{
                //    byte[] thisMacBytes = GetThisBTMacAddress().GetAddressBytes();
                //    BluetoothAddress ba = new BluetoothAddress(thisMacBytes);
                //}

                // BluetoothAddress targetMacAddress = BluetoothAddress.Parse("BEA5C3EF3F7D");
            }
        }

        private static void bluetoothComponent_DiscoverDevicesComplete(object sender, DiscoverDevicesEventArgs e)
        {
            Console.WriteLine("_bluetoothComponent_DiscoverDevicesComplete");
        }

        private static void bluetoothComponent_DiscoverDevicesProgress(object sender, DiscoverDevicesEventArgs e)
        {
            foreach (var item in e.Devices)
            {
                if (_btList.ContainsKey(item.DeviceName) == true)
                {
                    continue;
                }

                _btList.Add(item.DeviceName, item);

                Console.WriteLine(item.DeviceName);
                Console.WriteLine($"\tConnected {item.Connected}");
                Console.WriteLine($"\tAuthenticated {item.Authenticated}");
                Console.WriteLine($"\tDeviceAddress {item.DeviceAddress}");

                //foreach (var service in item.InstalledServices)
                //{
                //    Console.WriteLine("\t\t" + service);
                //}

                if (string.IsNullOrEmpty(_targetMachineName) == true)
                {
                    continue;
                }

                if (string.Equals(item.DeviceName, _targetMachineName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    ConnectTo(item);
                }
            }
        }

        private static void ConnectTo(BluetoothDeviceInfo item)
        {
            Guid serviceUUID = BluetoothService.TcpProtocol; // new Guid("00000004-0000-1000-8000-00805F9B34FB");

            BluetoothClient bluetoothClient = new BluetoothClient();
            bluetoothClient.Connect(item.DeviceAddress, serviceUUID);

            // BluetoothEndPoint bep = new BluetoothEndPoint(item.DeviceAddress, serviceUUID);
            // bluetoothClient.Connect(bep);

            Console.WriteLine("Connected");
            bluetoothClient.Client.WriteString("Hello");
            string result = bluetoothClient.Client.ReadString();
            Console.WriteLine(result);

            bluetoothClient.Close();
            Console.WriteLine("Closed");
        }

        public static PhysicalAddress GetThisBTMacAddress()
        {
            foreach (NetworkInterface nic in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (nic.Name == "Bluetooth Network Connection")
                {
                    return nic.GetPhysicalAddress();
                }
            }

            return null;
        }
    }
}

// Adapter profile that is in connecting to the internet
// var profile = InTheHand.Networking.Connectivity.NetworkInformation.GetInternetConnectionProfile();
// Console.WriteLine(profile);

// Another method to enumerate paired bluetooth devices
//string selector = InTheHand.Devices.Bluetooth.BluetoothDevice.GetDeviceSelector();
//foreach (var item in InTheHand.Devices.Enumeration.DeviceInformation.FindAll(selector))
//{
//    Console.WriteLine($"{item.Id}, {item.Name}, {item.Pairing}");
//}
