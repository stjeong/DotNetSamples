using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace NetDbgExt
{
    public static class NativeMethods
    {
        [DllImport("kernel32.dll")]
        public static extern void OutputDebugString(string lpOutputString);
    }
}
