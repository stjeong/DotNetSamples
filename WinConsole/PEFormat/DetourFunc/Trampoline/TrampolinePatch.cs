using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using WindowsPE;

namespace DetourFunc
{
    /*
스택 값 하나를 빼는 용도
// 58 pop, rax
    */
    public sealed class TrampolinePatch<T> : IDisposable where T : Delegate
    {
        byte[] _oldCode;
        IntPtr _funcAddress;

        public void JumpPatch(IntPtr codeAddress, IntPtr valueAddress)
        {
            byte[] code = GetJumpToCode(codeAddress, valueAddress);
            _oldCode = OverwriteCode(codeAddress, code);

            if (_oldCode.Length == code.Length)
            {
                _funcAddress = codeAddress;
            }
        }

        byte[] OverwriteCode(IntPtr codeAddress, byte[] code)
        {
            byte[] oldCode = new byte[code.Length];

            for (int i = 0; i < code.Length; i++)
            {
                oldCode[i] = codeAddress.ReadByte(i);
            }

            ProcessAccessRights rights = ProcessAccessRights.PROCESS_VM_OPERATION | ProcessAccessRights.PROCESS_VM_READ | ProcessAccessRights.PROCESS_VM_WRITE;
            PageAccessRights dwOldProtect = PageAccessRights.NONE;
            IntPtr hHandle = IntPtr.Zero;

            try
            {
                int pid = Process.GetCurrentProcess().Id;

                hHandle = NativeMethods.OpenProcess(rights, false, pid);
                if (hHandle == IntPtr.Zero)
                {
                    return null;
                }

#pragma warning disable IDE0059 // Unnecessary assignment of a value
                if (NativeMethods.VirtualProtectEx(hHandle, codeAddress, new UIntPtr((uint)IntPtr.Size), PageAccessRights.PAGE_EXECUTE_READWRITE, out dwOldProtect) == false)
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                {
                    return null;
                }

                codeAddress.WriteBytes(code);

                NativeMethods.FlushInstructionCache(hHandle, codeAddress, new UIntPtr((uint)code.Length));
                return oldCode;
            }
            finally
            {
                if (dwOldProtect != PageAccessRights.NONE)
                {
                    NativeMethods.VirtualProtectEx(hHandle, codeAddress, new UIntPtr((uint)IntPtr.Size), dwOldProtect, out PageAccessRights _);
                }

                if (hHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(hHandle);
                }
            }
        }

        byte[] GetJumpToCode(IntPtr codeAddress, IntPtr valueAddress)
        {
            // E9 0B 08 00 00 jmp 0000080b

            // 48 B8 FF FF FF FF FF FF FF 7F mov rax,7FFFFFFFFFFFFFFFh
            // FF E0 jmp rax 
            byte[] _longJumpToBytes = new byte[]
            {
                0x48, 0xB8, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                0xFF, 0xE0
            };

            byte[] _shortJumpToBytes = new byte[]
            {
                0xE9, 0x00, 0x00, 0x00, 0x00
            };

            if (IntPtr.Size == 8)
            {
                long offset = valueAddress.ToInt64() - codeAddress.ToInt64();

                if (Math.Abs(offset) <= Int32.MaxValue)
                {
                    byte[] buf = BitConverter.GetBytes((int)offset - _shortJumpToBytes.Length);
                    Array.Copy(buf, 0, _shortJumpToBytes, 1, 4);
                    return _shortJumpToBytes;
                }
                else
                {
                    byte[] buf = BitConverter.GetBytes(valueAddress.ToInt64());
                    Array.Copy(buf, 0, _longJumpToBytes, 2, IntPtr.Size);
                    return _longJumpToBytes;
                }
            }
            else
            {
                long offset = valueAddress.ToInt64() - codeAddress.ToInt64();

                byte[] buf = BitConverter.GetBytes(offset - _shortJumpToBytes.Length);
                Array.Copy(buf, 0, _shortJumpToBytes, 1, IntPtr.Size);
                return _shortJumpToBytes;
            }
        }

        public void Dispose()
        {
            if (_funcAddress != IntPtr.Zero && _oldCode.Length != 0)
            {
                OverwriteCode(_funcAddress, _oldCode);
                _funcAddress = IntPtr.Zero;
                _oldCode = null;
            }
        }
    }
}
