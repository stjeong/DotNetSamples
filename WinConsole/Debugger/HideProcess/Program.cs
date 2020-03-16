using KernelStructOffset;
using System;
using System.Diagnostics;

namespace HideProcess
{
    // Prerequisite:
    //  Register and start "KernelMemoryIO" kernel driver
    //  https://github.com/stjeong/KernelMemoryIO/tree/master/KernelMemoryIO
    //
    // sc create "KernelMemoryIO" binPath= "D:\Debug\KernelMemoryIO.sys" type= kernel start= demand
    // sc delete "KernelMemoryIO"
    // net start KernelMemoryIO
    // net stop KernelMemoryIO

    class Program
    {
        static readonly DbgOffset _ethreadOffset = DbgOffset.Get("_ETHREAD");
        static readonly DbgOffset _kthreadOffset = DbgOffset.Get("_KTHREAD");
        static readonly DbgOffset _eprocessOffset = DbgOffset.Get("_EPROCESS");

        static void Main(string[] _)
        {
            int processId = Process.GetCurrentProcess().Id;
            Console.WriteLine($"ThisPID: {processId}");

            IntPtr ethreadPtr = GetEThread(processId);
            if (ethreadPtr == IntPtr.Zero)
            {
                Console.WriteLine("THREAD handle not found");
                return;
            }

            Console.WriteLine($"_ETHREAD address: {ethreadPtr.ToInt64():x}");
            Console.WriteLine();

            using (KernelMemoryIO memoryIO = new KernelMemoryIO())
            {
                if (memoryIO.IsInitialized == false)
                {
                    Console.WriteLine("Failed to open device");
                    return;
                }

                {
                    // +0x648 Cid : _CLIENT_ID
                    IntPtr clientIdPtr = _ethreadOffset.GetPointer(ethreadPtr, "Cid");
                    _CLIENT_ID cid = memoryIO.ReadMemory<_CLIENT_ID>(clientIdPtr);

                    Console.WriteLine($"PID: {cid.Pid} ({cid.Pid:x})");
                    Console.WriteLine($"TID: {cid.Tid} ({cid.Tid:x})");

                    if (cid.Pid != processId)
                    {
                        return;
                    }
                }

                {
                    // +0x220 Process : Ptr64 _KPROCESS
                    IntPtr processPtr = _kthreadOffset.GetPointer(ethreadPtr, "Process");
                    IntPtr eprocessPtr = memoryIO.ReadMemory<IntPtr>(processPtr);
                    IntPtr activeProcessLinksPtr = _eprocessOffset.GetPointer(eprocessPtr, "ActiveProcessLinks");

                    // _LIST_ENTRY entry = memoryIO.ReadMemory<_LIST_ENTRY>(activeProcessLinksPtr);

                    Console.WriteLine("Press any key to hide this process from Task Manager");
                    Console.ReadLine();
                    IntPtr deletedEntry = IntPtr.Zero;

                    try
                    {
                        deletedEntry = Unlink(memoryIO, activeProcessLinksPtr);

                        Console.WriteLine();

                        Console.WriteLine("Press any key to unhide this process");
                        Console.ReadLine();
                    }
                    finally
                    {
                        RestoreLink(memoryIO, deletedEntry);
                    }

                    Console.WriteLine("Check this process appeared again in Task Manager");
                    Console.WriteLine("Press any key to exit");
                    Console.ReadLine();
                }
            }
        }

        private unsafe static void RestoreLink(KernelMemoryIO memoryIO, IntPtr deletedLink)
        {
            if (deletedLink == IntPtr.Zero)
            {
                return;
            }

            IntPtr baseLink = GetActiveProcessLinksFromAnotherProcess(memoryIO);
            if (baseLink == IntPtr.Zero)
            {
                Console.WriteLine("Can't find an appropriate ActiveProcessLinks");
                return;
            }

            Console.WriteLine($"Restore {deletedLink.ToInt64():x} to {baseLink.ToInt64():x}");
            _LIST_ENTRY baseEntry = memoryIO.ReadMemory<_LIST_ENTRY>(baseLink);

            IntPtr nextItem = baseEntry.Flink;

            memoryIO.WriteMemory<IntPtr>(deletedLink + 0              /* deletedLink.Flink */, nextItem);
            memoryIO.WriteMemory<IntPtr>(deletedLink + sizeof(IntPtr) /* deletedLink.Blink */, baseLink);

            memoryIO.WriteMemory<IntPtr>(baseLink + 0               /* baseLink.Flink */, deletedLink);
            memoryIO.WriteMemory<IntPtr>(nextItem + sizeof(IntPtr)  /* nextItem.Blink */, deletedLink);
        }

        private static IntPtr GetActiveProcessLinksFromAnotherProcess(KernelMemoryIO memoryIO)
        {
            foreach (Process process in Process.GetProcesses())
            {
                Console.WriteLine("Try-Target process name: " + process.ProcessName + ", " + process.Id);
                IntPtr ethreadPtr = GetEThread(process.Id);

                if (ethreadPtr == IntPtr.Zero)
                {
                    continue;
                }

                {
                    // +0x648 Cid : _CLIENT_ID
                    IntPtr clientIdPtr = _ethreadOffset.GetPointer(ethreadPtr, "Cid");
                    _CLIENT_ID cid = memoryIO.ReadMemory<_CLIENT_ID>(clientIdPtr);

                    if (cid.Pid != process.Id) // Check that ethreadPtr is valid
                    {
                        Console.WriteLine("NOT Match with Cid");
                        Console.WriteLine();
                        continue;
                    }
                }

                IntPtr processPtr = _kthreadOffset.GetPointer(ethreadPtr, "Process");
                IntPtr eprocessPtr = memoryIO.ReadMemory<IntPtr>(processPtr);
                IntPtr activeProcessLinksPtr = _eprocessOffset.GetPointer(eprocessPtr, "ActiveProcessLinks");

                return activeProcessLinksPtr;
            }

            return IntPtr.Zero;
        }

        private unsafe static IntPtr Unlink(KernelMemoryIO memoryIO, IntPtr linkPtr)
        {
            _LIST_ENTRY entry = memoryIO.ReadMemory<_LIST_ENTRY>(linkPtr);

            IntPtr pNext = entry.Flink;
            IntPtr pPrev = entry.Blink;

            memoryIO.WriteMemory<IntPtr>(pNext + sizeof(IntPtr) /* pNext.Blink */, pPrev);
            memoryIO.WriteMemory<IntPtr>(pPrev + 0              /* pPrev.Flink */, pNext);

            memoryIO.WriteMemory<IntPtr>(linkPtr + sizeof(IntPtr) /* linkPtr.Blink */, linkPtr);
            memoryIO.WriteMemory<IntPtr>(linkPtr + 0              /* linkPtr.Flink */, linkPtr);

            return linkPtr;
        }

        private static IntPtr GetEThread(int processId)
        {
            using (WindowsHandleInfo whi = new WindowsHandleInfo())
            {
                for (int i = 0; i < whi.HandleCount; i++)
                {
                    var she = whi[i];

                    if (she.OwnerPid != processId)
                    {
                        continue;
                    }

                    string _ = she.GetName(out string handleTypeName);

                    if (handleTypeName == "Thread")
                    {
                        return she.ObjectPointer;
                    }
                }
            }

            return IntPtr.Zero;
        }
    }
}
