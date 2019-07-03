using System;
using OpenCvSharp;

namespace src
{
    class Program
    {
        static void Main(string[] args)
        {
            Mat src = new Mat("lenna.png", ImreadModes.Grayscale);
            Mat dst = new Mat();

            Cv2.Canny(src, dst, 50, 200);
            using (new Window("src image", src))
            using (new Window("dst image", dst))
            {
                Cv2.WaitKey();
            }
        }
    }
}
