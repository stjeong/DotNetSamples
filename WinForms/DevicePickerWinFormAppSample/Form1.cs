using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using Windows.Devices.Bluetooth;
using Windows.Devices.Enumeration;

namespace DevicePickerWinFormAppSample
{
    // WPF/WinForm에서 UWP의 기능을 이용해 Bluetooth 기기와 Pairing하는 방법
    // http://www.sysnet.pe.kr/2/0/12001

    // add reference: C:\Program Files (x86)\Reference Assemblies\Microsoft\Framework\.NETCore\v4.5\System.Runtime.WindowsRuntime.dll

    // add reference: C:\Program Files (x86)\Windows Kits\10\UnionMetadata\...[version]...\Windows.winmd
    //                C:\Program Files (x86)\Windows Kits\10\UnionMetadata\10.0.18362.0\Windows.winmd
    //                C:\Program Files (x86)\Windows Kits\10\UnionMetadata\Windows.winmd
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void Button1_Click(object sender, EventArgs e)
        {
            var picker = new DevicePicker();
            picker.DeviceSelected += Picker_DeviceSelected;
            picker.Filter.SupportedDeviceSelectors.Add(
                BluetoothDevice.GetDeviceSelectorFromPairingState(false));

            // sync
            // picker.Show(new Windows.Foundation.Rect(100, 100, 500, 500));

            // async
            DeviceInformation di = await picker.PickSingleDeviceAsync(new Windows.Foundation.Rect(0, 0, 0, 0), Windows.UI.Popups.Placement.Default);
            if (di == null)
            {
                return;
            }

            if (di.Pairing.IsPaired == false)
            {
                di.Pairing.Custom.PairingRequested += Custom_PairingRequested;

                if (di.Pairing.CanPair == true)
                {
                    var pairResult = await di.Pairing.PairAsync();
                    if (pairResult.Status == DevicePairingResultStatus.Paired)
                    {
                        // do something
                    }
                }
                else
                {
                    var pairResult = await di.Pairing.Custom.PairAsync(DevicePairingKinds.ProvidePin);
                    if (pairResult.Status == DevicePairingResultStatus.Paired)
                    {
                        // do something
                    }
                }
            }
        }

        private void Custom_PairingRequested(DeviceInformationCustomPairing sender, DevicePairingRequestedEventArgs args)
        {
            if (args.PairingKind == DevicePairingKinds.ProvidePin)
            {
                // add reference: Microsoft.VisualBasic.dll
                string result = Microsoft.VisualBasic.Interaction.InputBox("Pin?", "Pairing...", "");
                if (string.IsNullOrEmpty(result) == false)
                {
                    args.Accept(result);
                }
            }
            else
            {
                args.Accept();
            }
        }

        private void Picker_DeviceSelected(DevicePicker sender, DeviceSelectedEventArgs args)
        {
            var di = args.SelectedDevice;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(di.Name);
            sb.AppendLine(di.Id);
            sb.AppendLine(di.Kind.ToString());
            sb.AppendLine(di.Pairing.IsPaired.ToString());
            sb.AppendLine(di.Pairing.CanPair.ToString());
            sb.AppendLine("");
            foreach (var item in di.Properties.Keys)
            {
                sb.Append(item);

                string[] props = di.Properties[item] as string[];
                if (props != null)
                {
                    sb.AppendLine(":");
                    foreach (var prop in props)
                    {
                        sb.AppendLine("    " + prop);
                    }
                }
                else
                {
                    sb.AppendLine(": " + di.Properties[item]);
                }
            }

            Trace.WriteLine(sb.ToString());
        }
    }
}
