using KernelStructOffset;
using System;
using System.Linq;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace HideModule
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(Process.GetCurrentProcess().Id);

            _PEB peb = EnvironmentBlockInfo.GetPeb();
            _PEB_LDR_DATA ldrData = _PEB_LDR_DATA.Create(peb.Ldr);

            string moduleName = "ole32.dll";
            FindModules(moduleName);

            DllOrderLink hiddenModuleLink = null;

            try
            {
                hiddenModuleLink = ldrData.HideDLL(moduleName);
                FindModules(moduleName);

                Console.ReadLine();
            }
            finally
            {
                if (hiddenModuleLink != null)
                {
                    ldrData.UnhideDLL(hiddenModuleLink);
                }
            }

            FindModules(moduleName);

            Console.WriteLine(Process.GetCurrentProcess().Id);
            Console.ReadLine();

            Console.WriteLine("[Load Order]");
            foreach (var item in ldrData.EnumerateLoadOrderModules())
            {
                Console.WriteLine("\t" + item.FullDllName.GetText());
            }

            Console.WriteLine();
            Console.WriteLine("[Memory Order]");
            foreach (var item in ldrData.EnumerateMemoryOrderModules())
            {
                Console.WriteLine("\t" + item.FullDllName.GetText());
            }
        }

        private static void FindModules(string moduleName)
        {
            bool found = false;
            foreach (ProcessModule pm in Process.GetCurrentProcess().Modules)
            {
                if (pm.FileName.EndsWith(moduleName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    found = true;
                    break;
                }
            }

            Console.WriteLine($"{moduleName}: " + found);
        }
    }
}
