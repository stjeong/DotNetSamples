using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KernelStructOffset
{
    public static class HelperExtension
    {
        public static unsafe IntPtr ReadPtr(this IntPtr ptr)
        {
            if (IntPtr.Size == 4)
            {
                int* ptrInt = (int*)ptr.ToPointer();
                return new IntPtr(*ptrInt);
            }
            else
            {
                long* ptrLong = (long*)ptr.ToPointer();
                return new IntPtr(*ptrLong);
            }
        }
    }
}
