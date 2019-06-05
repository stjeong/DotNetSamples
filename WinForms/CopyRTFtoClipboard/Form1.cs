using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CopyRTFtoClipboard
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
//            string txt = @"{\rtf1\ansi\ansicpg949\deff0\nouicompat\deflang1033\deflangfe1042{\fonttbl{\f0\fnil\fcharset0 Malgun Gothic;}{\f1\fnil\fcharset129 Malgun Gothic;}{\f2\fnil\fcharset129 \'b8\'bc\'c0\'ba \'b0\'ed\'b5\'f1;}}
//{\colortbl ;\red0\green0\blue0;\red0\green100\blue0;\red0\green0\blue255;}
//{\*\generator Riched20 10.0.17763}\viewkind4\uc1 
//\pard\cf1\f0\fs20\lang1042\par
//\cf2 #\f1\'c5\'d7\'bd\'ba\'c6\'ae\'c0\'d4\'b4\'cf\'b4\'d92.\cf1\par
//\par

//\pard\sa200\sl276\slmult1\cf3 Get-Item\cf0\f2\lang18\par}
//";

            string txt = File.ReadAllText("data.txt");

            DataObject dataObject = new DataObject();
            dataObject.SetData(DataFormats.Rtf, txt);
            Clipboard.SetDataObject(dataObject);
        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            foreach (string txt in Clipboard.GetDataObject().GetFormats())
            {
                System.Diagnostics.Trace.WriteLine(txt);

                object objValue = Clipboard.GetData(txt);
                if (objValue != null)
                {
                    System.Diagnostics.Trace.WriteLine(objValue);

                    WriteText(objValue as MemoryStream);
                }
            }
        }

        private void WriteText(MemoryStream ms)
        {
            if (ms == null)
            {
                return;
            }

            byte[] buf = new byte[ms.Length];
            ms.Read(buf, 0, buf.Length);

            string txt = Encoding.UTF8.GetString(buf);
            System.Diagnostics.Trace.WriteLine(txt);
        }
    }
}
