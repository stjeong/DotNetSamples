using System;
using Cairo;

namespace CairoSharp_x64
{
    // C# - CairoSharp/GtkSharp 사용을 위한 프로젝트 구성 방법
    // http://www.sysnet.pe.kr/2/0/11931

    class Program
    {
        static void Main(string[] args)
        {
            int width = 400;
            int height = 400;

            using (Surface surface = new ImageSurface(Cairo.Format.Rgb24, width, height))
            {
                surface.WriteToPng("plplot1.png");
            }
        }
    }
}
