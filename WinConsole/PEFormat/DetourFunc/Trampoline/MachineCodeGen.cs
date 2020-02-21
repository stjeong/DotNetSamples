using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc
{
    public class MachineCodeGen<T> : IDisposable where T: Delegate
    {
        IntPtr _codePointer;

        public T GetFunc(byte[] codeBytes)
        {
            _codePointer = NativeMethods.VirtualAlloc(IntPtr.Zero, new UIntPtr((uint)codeBytes.Length),
                   AllocationType.COMMIT | AllocationType.RESERVE,
                   MemoryProtection.EXECUTE_READWRITE
               );

            Marshal.Copy(codeBytes, 0, _codePointer, codeBytes.Length);
            return (T)Marshal.GetDelegateForFunctionPointer(_codePointer, typeof(T));
        }

        public void Dispose()
        {
            if (_codePointer != IntPtr.Zero)
            {
                NativeMethods.VirtualFree(_codePointer, 0, 0x8000);
                _codePointer = IntPtr.Zero;
            }
        }
    }
}
