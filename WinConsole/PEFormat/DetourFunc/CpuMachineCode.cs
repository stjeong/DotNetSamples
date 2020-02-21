using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace DetourFunc
{
    public static class CpuMachineCode
    {
        [UnmanagedFunctionPointerAttribute(CallingConvention.Cdecl)]
        public delegate void CpuIdDelegate(byte[] buffer);

        public static byte[] GetCpuId()
        {
            if (IntPtr.Size != 4)
            {
                return null;
            }

            byte[] x86CpuIdBytes =
            {
                0x55,
                0x8B, 0xEC,
                0x53,
                0x57,

                0x33, 0xDB,
                0x33, 0xC9,
                0x33, 0xD2,
                0xB8, 0x00, 0x00, 0x00, 0x00,

                0x0F, 0xA2, // cpuid

                0x8B, 0x7D, 0x08,

                0x89, 0x07,
                0x89, 0x5F, 0x04,
                0x89, 0x4F, 0x08,
                0x89, 0x57, 0x0C,

                0x5F,
                0x5B,

                0x5D,
                0xC3, // ret
            };

            using (var item = new MachineCodeGen<CpuIdDelegate>())
            {
                CpuIdDelegate func = item.GetFunc(x86CpuIdBytes);
                if (func != null)
                {
                    byte[] cpudIdBytes = new byte[4 * 4];
                    GCHandle handle = GCHandle.Alloc(cpudIdBytes, GCHandleType.Pinned);

                    try
                    {
                        func(cpudIdBytes);
                        return cpudIdBytes;
                    }
                    finally
                    {
                        if (handle.IsAllocated)
                        {
                            handle.Free();
                        }
                    }
                }

                return null;
            }
        }
    }
}
