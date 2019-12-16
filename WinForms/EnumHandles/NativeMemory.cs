using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace EnumHandles
{
    public unsafe ref struct NativeMemory<T> where T : unmanaged
    {
        int _size;
        IntPtr _ptr;

        public NativeMemory(int size, IntPtr ptr)
        {
            _size = size;
            _ptr = ptr;
        }

        public Span<T> GetView()
        {
            return new Span<T>(_ptr.ToPointer(), _size);
        }

        // C# 8.0에서만 using과 함께 사용 가능
        public void Dispose()
        {
            if (_ptr == IntPtr.Zero)
            {
                return;
            }

            Marshal.FreeHGlobal(_ptr);
            _ptr = IntPtr.Zero;
        }
    }
}
