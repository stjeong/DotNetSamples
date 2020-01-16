using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDbgExt
{
    public static class HelperExtension
    {
        public unsafe static void WriteValue<T>(this IntPtr ptr, T value) where T : unmanaged
        {
            T* pValue = (T*)ptr.ToPointer();
            *pValue = value;
        }
    }
}
