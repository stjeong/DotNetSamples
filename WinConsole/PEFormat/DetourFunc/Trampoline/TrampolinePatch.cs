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
        IntPtr _fromMethodAddress;
        MachineCodeGen<T> _originalCode;
        T _originalMethod;

        readonly byte[] _longJumpTemplate = new byte[]
        {
            // 48 B8 FF FF FF FF FF FF FF 7F mov rax,7FFFFFFFFFFFFFFFh
            // FF E0 jmp rax 
            0x48, 0xB8, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
            0xFF, 0xE0
        };

        readonly byte[] _shortJumpTemplate = new byte[]
        {
            // E9 0B 08 00 00 jmp 0000080b
            0xE9, 0x00, 0x00, 0x00, 0x00
        };

        public bool JumpPatch(IntPtr fromMethodAddress, IntPtr toMethodAddress)
        {
            byte[] code = GetJumpToCode(fromMethodAddress, 0, toMethodAddress);
            byte[] oldCode = GetOldCode(fromMethodAddress, code.Length);

            OverwriteCode(fromMethodAddress, code);

            if (oldCode.Length >= code.Length)
            {
                _oldCode = oldCode;
                _fromMethodAddress = fromMethodAddress;

                return true;
            }

            return false;
        }

        private byte[] GetOldCode(IntPtr codeAddress, int maxBytes)
        {
            SharpDisasm.ArchitectureMode mode = (IntPtr.Size == 8) ? SharpDisasm.ArchitectureMode.x86_64 : SharpDisasm.ArchitectureMode.x86_32;
            List<byte> entranceCodes = new List<byte>();

            int totalLen = 0;
            using (var disasm = new SharpDisasm.Disassembler(codeAddress, maxBytes + NativeMethods.MaxLengthOpCode, mode))
            {
                foreach (var insn in disasm.Disassemble())
                {
                    for (int i = 0; i < insn.Length; i++)
                    {
                        entranceCodes.Add(codeAddress.ReadByte(totalLen + i));
                    }

                    totalLen += insn.Length;

                    if (totalLen >= maxBytes)
                    {
                        return entranceCodes.ToArray();
                    }
                }
            }

            return null;
        }

        byte[] OverwriteCode(IntPtr codeAddress, byte[] code)
        {
            byte[] oldCode = codeAddress.ReadBytes(code.Length);

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

        public T GetOriginalFunc()
        {
            if (_originalMethod == null)
            {
                List<byte> newJump = new List<byte>();
                newJump.AddRange(_oldCode);

                _originalCode = new MachineCodeGen<T>();
                IntPtr fromAddress = _originalCode.Alloc(newJump.Count + NativeMethods.MaxLengthOpCode * 2);

                byte[] jumpCode = GetJumpToCode(fromAddress, newJump.Count, _fromMethodAddress + _oldCode.Length);
                newJump.AddRange(jumpCode);

                _originalMethod = _originalCode.GetFunc(newJump.ToArray()) as T;
            }

            return _originalMethod;
        }

        byte[] GetJumpToCode(IntPtr fromAddress, int prologueLengthOnFromAddress, IntPtr toAddress)
        {
            long offset = toAddress.ToInt64() - fromAddress.ToInt64();

            if (IntPtr.Size == 8)
            {
                if (Math.Abs(offset) > (Int32.MaxValue - NativeMethods.MaxLengthOpCode * 10))
                {
                    byte[] longJumpToBytes = _longJumpTemplate.ToArray();
                    byte[] buf8 = BitConverter.GetBytes(toAddress.ToInt64());
                    Array.Copy(buf8, 0, longJumpToBytes, 2, IntPtr.Size);
                    return longJumpToBytes;
                }
            }

            byte[] shortJumpToBytes = _shortJumpTemplate.ToArray();
            byte[] buf4 = BitConverter.GetBytes(offset - (prologueLengthOnFromAddress + shortJumpToBytes.Length));
            Array.Copy(buf4, 0, shortJumpToBytes, 1, 4);
            return shortJumpToBytes;
        }

        public void Dispose()
        {
            if (_fromMethodAddress != IntPtr.Zero && _oldCode.Length != 0)
            {
                OverwriteCode(_fromMethodAddress, _oldCode);
                _fromMethodAddress = IntPtr.Zero;
                _oldCode = null;
            }
        }
    }
}
