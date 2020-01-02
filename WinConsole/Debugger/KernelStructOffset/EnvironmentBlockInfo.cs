using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KernelStructOffset
{
    public class EnvironmentBlockInfo
    {
        private readonly static byte[] _x64TebBytes =
        {
            0x40, 0x57, // push rdi
        
            0x65, 0x48, 0x8B, 0x04, 0x25, 0x30, 0x00, 0x00, 0x00, // mov rax, qword ptr gs:[30h]

            0x5F, // pop rdi
            0xC3, // ret
        };

        static IntPtr _codePointer;
        static GetTebDelegate _getTebDelg;

        [UnmanagedFunctionPointerAttribute(CallingConvention.Cdecl)]
        private delegate long GetTebDelegate();

        public const string PROCESS_ENVIRONMENT_BLOCK = "ProcessEnvironmentBlock";

        static EnvironmentBlockInfo()
        {
            if (IntPtr.Size != 8)
            {
                return;
            }

            byte[] codeBytes = _x64TebBytes;

            _codePointer = NativeMethods.VirtualAlloc(IntPtr.Zero, new UIntPtr((uint)codeBytes.Length),
                AllocationType.COMMIT | AllocationType.RESERVE,
                MemoryProtection.EXECUTE_READWRITE
            );

            Marshal.Copy(codeBytes, 0, _codePointer, codeBytes.Length);

            _getTebDelg = (GetTebDelegate)Marshal.GetDelegateForFunctionPointer(
                _codePointer, typeof(GetTebDelegate));
        }

        public static IntPtr GetTebAddress()
        {
            if (IntPtr.Size == 8)
            {
                if (_getTebDelg == null)
                {
                    throw new ObjectDisposedException("GetTebAddress");
                }

                return new IntPtr(_getTebDelg());
            }
            else
            {
                return NativeMethods.NtCurrentTeb();
            }
        }

        public static IntPtr GetPebAddress(out IntPtr tebAddress)
        {
            tebAddress = GetTebAddress();
            var dict = GetTebOffset();

            if (dict.ContainsKey(PROCESS_ENVIRONMENT_BLOCK) == false)
            {
                // Retry one more for x86 on x64 of Windows Server 2008 R2
                dict = DbgOffset.Get("ole32!_TEB", "ole32.dll", "c:\\windows\\syswow64\\notepad.exe");
            }

            if (dict.ContainsKey(PROCESS_ENVIRONMENT_BLOCK) == false)
            {
                return IntPtr.Zero;
            }

            int offset = dict[PROCESS_ENVIRONMENT_BLOCK];
            return Marshal.ReadIntPtr(tebAddress + offset);
        }

        private static Dictionary<string, int> GetTebOffset()
        {
            string typeInfo = GetTebOwner(out string moduleName);
            return DbgOffset.Get(typeInfo, moduleName);
        }

        private static string GetTebOwner(out string moduleName)
        {
            if (IntPtr.Size == 4 && Environment.Is64BitOperatingSystem == true)
            {
                moduleName = "ntdll.dll";
                return "_TEB32";
            }

            moduleName = "ntdll.dll";
            return "_TEB";
        }
    }
}
