using KernelStructOffset;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsPE
{
    // mridgers/pdbdump.c
    // https://gist.github.com/mridgers/2968595
    public class PdbDump : IDisposable
    {
        IntPtr _hProcess = IntPtr.Zero;
        string _pdbFilePath;
        IntPtr _baseAddress;
        int _memorySize;

        bool _symInitialized = false;

        public PdbDump(string pdbFilePath, IntPtr baseAddress, int memorySize)
        {
            uint options = NativeMethods.SymGetOptions();
            Console.WriteLine($"SymGetOptions: {options}");

            options &= ~(uint)SymOpt.SYMOPT_DEFERRED_LOADS;
            options |= (uint)SymOpt.SYMOPT_LOAD_LINES;
            options |= (uint)SymOpt.SYMOPT_IGNORE_NT_SYMPATH;
#if ENABLE_DEBUG_OUTPUT
            options |= (uint)SymOpt.SYMOPT_DEBUG;
#endif
            options |= (uint)SymOpt.SYMOPT_UNDNAME;

            NativeMethods.SymSetOptions(options);

            int pid = Process.GetCurrentProcess().Id;
            IntPtr processHandle = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_QUERY_INFORMATION | ProcessAccessRights.PROCESS_VM_READ, false, pid);

            if (NativeMethods.SymInitialize(processHandle, null, false) == false)
            {
                return;
            }

            _symInitialized = true;

            _hProcess = processHandle;
            _pdbFilePath = pdbFilePath;
            _memorySize = memorySize;
            _baseAddress = LoadPdbModule(_hProcess, _pdbFilePath, baseAddress, (uint)_memorySize);
        }

        public static PdbStore CreateSymbolStore(string pdbFilePath, IntPtr baseAddress, int memorySize)
        {
            PdbStore store = null;

            using (PdbDump pdbDumper = new PdbDump(pdbFilePath, baseAddress, memorySize))
            {
                store = pdbDumper.GetStore((pStore) =>
                {
                    NativeMethods.SymEnumSymbols(pdbDumper._hProcess, (ulong)baseAddress.ToInt64(), "*", enum_proc, pStore);
                });
            }

            return store;
        }

        public static PdbStore CreateTypeStore(string pdbFilePath, IntPtr baseAddress, int memorySize)
        {
            PdbStore store = null;

            using (PdbDump pdbDumper = new PdbDump(pdbFilePath, baseAddress, memorySize))
            {
                store = pdbDumper.GetStore((pStore) =>
                {
                    NativeMethods.SymEnumTypes(pdbDumper._hProcess, (ulong)baseAddress.ToInt64(), enum_proc, pStore);
                });
            }

            return store;
        }

        public void Dispose()
        {
            if (_hProcess != IntPtr.Zero)
            {
                NativeMethods.CloseHandle(_hProcess);
                _hProcess = IntPtr.Zero;
            }

            if (_symInitialized == true)
            {
                NativeMethods.SymCleanup(_hProcess);
                _symInitialized = false;
            }
        }

        public IEnumerable<SYMBOL_INFO> EnumerateTypes()
        {
            return Enumerate((pStore) =>
            {
                NativeMethods.SymEnumTypes(_hProcess, (ulong)_baseAddress.ToInt64(), enum_proc, pStore);
            });
        }

        public IEnumerable<SYMBOL_INFO> EnumerateSymbols()
        {
            return Enumerate((pStore) =>
            {
                NativeMethods.SymEnumSymbols(_hProcess, (ulong)_baseAddress.ToInt64(), "*", enum_proc, pStore);
            });
        }

        private IEnumerable<SYMBOL_INFO> Enumerate(Action<IntPtr> action)
        {
            PdbStore store = GetStore(action);
            foreach (SYMBOL_INFO si in store.Enumerate())
            {
                yield return si;
            }
        }

        private PdbStore GetStore(Action<IntPtr> action)
        {
            PdbStore store = new PdbStore();

            IntPtr pStore = Marshal.AllocHGlobal(16); // 16 == sizeof(VARIANT)
            Marshal.GetNativeVariantForObject(store, pStore);

            action(pStore);

            return store;
        }

        private static unsafe bool enum_proc(IntPtr pinfo, uint size, IntPtr pUserContext)
        {
            SYMBOL_INFO info = SYMBOL_INFO.Create(pinfo);

            PdbStore pdbStore = (PdbStore)Marshal.GetObjectForNativeVariant(pUserContext);
            pdbStore.Add(info);

            return true;
        }

        // It's possible even if processHandle is not real Handle. (For example processHandle == 0x493)
        // Also, base_addr and moduleSize can be arbitrary.
        private static unsafe IntPtr LoadPdbModule(IntPtr processHandle, string pdbFilePath)
        {
            IntPtr base_addr = new IntPtr(0x400000);
            return LoadPdbModule(processHandle, pdbFilePath, base_addr, 0x7fffffff);
        }

        private static unsafe IntPtr LoadPdbModule(IntPtr processHandle, string pdbFilePath, IntPtr baseAddress, uint moduleSize)
        {
            return new IntPtr((long)NativeMethods.SymLoadModuleEx(processHandle,
                IntPtr.Zero, pdbFilePath, null, baseAddress.ToInt64(), moduleSize, null, 0));
        }
    }
}
