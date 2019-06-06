using System;
using Cairo;

namespace CairoSharp_x64
{
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
