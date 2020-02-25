using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc
{
    public sealed class MachineCodeGen<T> : IDisposable where T: Delegate
    {
        IntPtr _codePointer;
        public IntPtr CodePointer => _codePointer;

        public IntPtr Alloc(int length)
        {
            _codePointer = NativeMethods.VirtualAlloc(IntPtr.Zero, new UIntPtr((uint)length),
                   AllocationType.COMMIT | AllocationType.RESERVE,
                   MemoryProtection.EXECUTE_READWRITE
               );

            return _codePointer;
        }

        public T GetFunc(byte[] codeBytes)
        {
            Marshal.Copy(codeBytes, 0, _codePointer, codeBytes.Length);
            return (T)Marshal.GetDelegateForFunctionPointer(_codePointer, typeof(T));
        }

        public void Dispose()
        {
            if (_codePointer != IntPtr.Zero)
            {
                NativeMethods.VirtualFree(_codePointer, 0, MemFreeType.MEM_RELEASE);
                _codePointer = IntPtr.Zero;
            }
        }
    }
}
