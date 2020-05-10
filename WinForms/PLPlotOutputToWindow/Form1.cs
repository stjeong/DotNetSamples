using PLplot;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PLPlotOutputToWindow
{
    // C# - PLplot 출력을 파일이 아닌 Window 화면으로 변경
    // http://www.sysnet.pe.kr/2/0/11935

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Draw();
        }

        void Draw()
        {
            int width = 400;
            int height = 400;
            const int RGBsize = 3;

            byte[] buffer = new byte[width * height * RGBsize];
            for (int i = 0; i < buffer.Length; i++)
            {
                buffer[i] = 0xff;
            }

            GCHandle handle = GCHandle.Alloc(buffer, GCHandleType.Pinned);

            try
            {
                IntPtr pBuf = handle.AddrOfPinnedObject();

                using (var pl = new PLStream())
                {
                    pl.sdev("mem");
                    pl.smem(width, height, pBuf);

                    pl.init();

                    pl.env(0, 10.0, 0, 10.0, 0, 0);
                    pl.join(1.0, 2.0, 7.0, 8.0);
                }

                Bitmap bmp = new Bitmap(width, height, 3 * width, PixelFormat.Format24bppRgb, pBuf);
                this.pictureBox1.Width = width;
                this.pictureBox1.Height = height;
                this.pictureBox1.Image = bmp;

                // bmp.Save("test.bmp");
            }
            finally
            {
                handle.Free();
            }
        }

        private void drawPLplot_Click(object sender, EventArgs e)
        {
            Draw();
        }
    }
}
