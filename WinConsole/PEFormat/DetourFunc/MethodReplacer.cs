using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc
{
    // https://www.sysnet.pe.kr/2/0/12132
    public class MethodReplacer : IDisposable
    {
        IntPtr _addressOrgFuncAddr;
        IntPtr _orgFuncAddr;

        private MethodReplacer() { }

        public static MethodReplacer Win32FuncWithManagedFunc<T>(T dllImportMethod, T funcToReplace, out T saveOrgFunc) where T : Delegate
        {
            saveOrgFunc = null;

            foreach (var attr in dllImportMethod.Method.GetCustomAttributes(false))
            {
                if (attr is DllImportAttribute dllImportAttr)
                {
                    string dllName = dllImportAttr.Value;
                    string funcName = dllImportAttr.EntryPoint;

                    if (string.IsNullOrEmpty(dllName) == false && string.IsNullOrEmpty(funcName) == false)
                    {
                        return Win32FuncWithManagedFunc(dllName, funcName, funcToReplace, out saveOrgFunc);
                    }
                }
            }

            return null;
        }

        public static MethodReplacer Win32FuncWithManagedFunc<T>(string win32Dll, string win32FuncName, T funcToReplace, out T saveOrgFunc) where T : Delegate
        {
            saveOrgFunc = null;

            IntPtr orgFuncAddr = GetEATFunctionAddress(win32Dll, win32FuncName, out IntPtr addressOrgFuncAddr);
            if (orgFuncAddr == IntPtr.Zero || addressOrgFuncAddr == IntPtr.Zero)
            {
                return null;
            }

            if (typeof(T).IsGenericType == true)
            {
                return null;
            }

            IntPtr proxyFuncAddr = Marshal.GetFunctionPointerForDelegate(funcToReplace);
            if (proxyFuncAddr == IntPtr.Zero)
            {
                return null;
            }

            MethodReplacer mr = new MethodReplacer();
            if (mr.Win32FuncWith<T>(addressOrgFuncAddr, orgFuncAddr, proxyFuncAddr, out saveOrgFunc) == false)
            {
                return null;
            }

            return mr;
        }

        public static MethodReplacer Win32FuncWithExportFunc<T>(string win32Dll, string win32FuncName, T funcToReplace, out T saveOrgFunc) where T : Delegate
        {
            saveOrgFunc = null;

            IntPtr orgFuncAddr = GetEATFunctionAddress(win32Dll, win32FuncName, out IntPtr addressOrgFuncAddr);
            if (orgFuncAddr == IntPtr.Zero || addressOrgFuncAddr == IntPtr.Zero)
            {
                return null;
            }

            string moduleName = Path.GetFileName(funcToReplace.GetType().Assembly.Location);
            string funcName = funcToReplace.Method.Name;

            IntPtr proxyFuncAddr = GetEATFunctionAddress(moduleName, funcName, out IntPtr addressToProxyFuncAddr);
            if (proxyFuncAddr == IntPtr.Zero || addressToProxyFuncAddr == IntPtr.Zero)
            {
                return null;
            }

            MethodReplacer mr = new MethodReplacer();
            if (mr.Win32FuncWith<T>(addressOrgFuncAddr, orgFuncAddr, proxyFuncAddr, out saveOrgFunc) == false)
            {
                return null;
            }

            return mr;
        }

        private bool Win32FuncWith<T>(IntPtr addressOrgFuncAddr, IntPtr orgFuncAddr, IntPtr proxyFuncAddr, out T saveOrgFunc) where T : Delegate
        {
            saveOrgFunc = null;

            T orgFuncDelegate = Marshal.GetDelegateForFunctionPointer(orgFuncAddr, typeof(T)) as T;

            if (orgFuncDelegate == null)
            {
                return false;
            }

            if (WriteProtectedAddress(addressOrgFuncAddr, proxyFuncAddr) == false)
            {
                return false;
            }

            _addressOrgFuncAddr = addressOrgFuncAddr;
            _orgFuncAddr = orgFuncAddr;

            saveOrgFunc = orgFuncDelegate;
            return true;
        }

        private static bool WriteProtectedAddress(IntPtr address, IntPtr value)
        {
            ProcessAccessRights rights = ProcessAccessRights.PROCESS_VM_OPERATION | ProcessAccessRights.PROCESS_VM_READ | ProcessAccessRights.PROCESS_VM_WRITE;

            int pid = Process.GetCurrentProcess().Id;
            IntPtr hHandle = IntPtr.Zero;
            PageAccessRights dwOldProtect = PageAccessRights.NONE;

            try
            {
                hHandle = NativeMethods.OpenProcess(rights, false, pid);
                if (hHandle == IntPtr.Zero)
                {
                    return false;
                }

#pragma warning disable IDE0059 // Unnecessary assignment of a value
                if (NativeMethods.VirtualProtectEx(hHandle, address, new UIntPtr((uint)IntPtr.Size), PageAccessRights.PAGE_EXECUTE_READWRITE, out dwOldProtect) == false)
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                {
                    return false;
                }

                if (IntPtr.Size == 4)
                {
                    address.WriteInt32(value.ToInt32());
                }
                else
                {
                    address.WriteInt64(value.ToInt64());
                }

                NativeMethods.FlushInstructionCache(hHandle, address, new UIntPtr((uint)IntPtr.Size));

                return true;
            }
            finally
            {
                if (dwOldProtect != PageAccessRights.NONE)
                {
                    NativeMethods.VirtualProtectEx(hHandle, address, new UIntPtr((uint)IntPtr.Size), dwOldProtect, out PageAccessRights _);
                }

                if (hHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(hHandle);
                }
            }
        }

        private static IntPtr GetEATFunctionAddress(string moduleName, string functionName, out IntPtr ptrContainingFuncAddress)
        {
            ptrContainingFuncAddress = IntPtr.Zero;

            SharpDisasm.ArchitectureMode mode = (IntPtr.Size == 8) ? SharpDisasm.ArchitectureMode.x86_64 : SharpDisasm.ArchitectureMode.x86_32;
            SharpDisasm.Disassembler.Translator.IncludeAddress = true;
            SharpDisasm.Disassembler.Translator.IncludeBinary = true;

            PEImage img = PEImage.FromLoadedModule(moduleName);
            if (img == null)
            {
                return IntPtr.Zero;
            }

            foreach (var efi in img.EnumerateExportFunctions())
            {
                if (efi.Name == functionName)
                {
                    IntPtr funcAddr = img.BaseAddress + (int)efi.RvaAddress;

                    byte[] code = img.ReadBytes(efi.RvaAddress, IntPtr.Size * 10);

                    var disasm = new SharpDisasm.Disassembler(code, mode, (ulong)funcAddr.ToInt64(), true);

                    foreach (var insn in disasm.Disassemble().Take(5))
                    {
                        switch (insn.Mnemonic)
                        {
                            case SharpDisasm.Udis86.ud_mnemonic_code.UD_Imov:
                                if (insn.Operands.Length == 2)
                                {
                                    long value = insn.Operands[1].Value;
                                    IntPtr jumpPtr = new IntPtr(value);
                                    ptrContainingFuncAddress = jumpPtr;
                                    return new IntPtr(IntPtr.Size == 8 ? jumpPtr.ReadInt64() : jumpPtr.ReadInt32());
                                }
                                break;

                            case SharpDisasm.Udis86.ud_mnemonic_code.UD_Ijmp:
                                if (insn.Operands.Length == 1)
                                {
                                    long value = insn.Operands[0].Value;
                                    IntPtr jumpPtr = new IntPtr(value);

                                    if (IntPtr.Size == 8)
                                    {
                                        jumpPtr = new IntPtr((long)insn.PC + value);
                                    }

                                    ptrContainingFuncAddress = jumpPtr;
                                    return new IntPtr(IntPtr.Size == 8 ? jumpPtr.ReadInt64() : jumpPtr.ReadInt32());
                                }
                                break;
                        }

                    }

                    break;
                }
            }

            return IntPtr.Zero;
        }

        bool _disposed = false;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed == false)
            {
                if (disposing == true)
                {
                    if (_addressOrgFuncAddr != IntPtr.Zero && _orgFuncAddr != IntPtr.Zero)
                    {
                        WriteProtectedAddress(_addressOrgFuncAddr, _orgFuncAddr);
                    }
                }

                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
