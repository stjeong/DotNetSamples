using System;
using System.Diagnostics;
using System.Threading;

class Program
{
    private static PerformanceCounter freeMemory;
    private static PerformanceCounter modifiedMemory;

    static void Main(string[] args)
    {
        freeMemory = new PerformanceCounter("Memory", "Free & Zero Page List Bytes", true);
        modifiedMemory = new PerformanceCounter("Memory", "Modified Page List Bytes", true);

        while (true)
        {
            PERFORMANCE_INFORMATION pi = new PERFORMANCE_INFORMATION();
            pi.Initialize();
            SafeNativeMethods.GetPerformanceInfo(out pi, pi.cb);

            Console.WriteLine("[Resource Monitor]");

            SafeNativeMethods.GetPhysicallyInstalledSystemMemory(out ulong installedMemory);

            double reserved = (installedMemory - (pi.Total / 1024.0));
            ulong modified = (ulong)modifiedMemory.RawValue;
            ulong inuse = pi.Total - pi.Available - modified;

            long reservedMB = (long)Math.Round(reserved / 1024.0);
            Console.WriteLine($"Hardware Reserved: {reservedMB} MB");

            Console.WriteLine($"In Use: {inuse / 1024 / 1024} MB");
            Console.WriteLine($"Modified: {modified / 1024 / 1024} MB");

            ulong free = (ulong)freeMemory.RawValue;
            ulong standby = pi.Available - free;
            Console.WriteLine($"Standby: {standby / 1024 / 1024} MB");
            Console.WriteLine($"Free: {free / 1024 / 1024} MB");
            Console.WriteLine();
            Console.WriteLine($"Available: {pi.Available.MB()} MB");
            Console.WriteLine($"Cached: {(standby + modified).MB()} MB");
            Console.WriteLine($"Total: {pi.Total.MB()} MB");
            Console.WriteLine($"Installed: {installedMemory / 1024} MB");

            //IntPtr queryResult = SafeNativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemFullMemoryInformation, 0);
            //if (queryResult == IntPtr.Zero)
            //{
            //    Console.WriteLine(queryResult.ToInt64());
            //}

            MEMORYSTATUSEX globalMemoryStatus = new MEMORYSTATUSEX();
            globalMemoryStatus.Initialize();
            SafeNativeMethods.GlobalMemoryStatusEx(ref globalMemoryStatus);

            Console.WriteLine();
            Console.WriteLine("[Task Manager]");

            Console.WriteLine($"Memory: {installedMemory / 1024.0 / 1024.0} GB");
            Console.WriteLine($"Memory usage: {pi.Total / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine();
            Console.WriteLine($"In use: {inuse / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine($"Available: {pi.Available / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine($"Committed: {pi.Commit / 1024.0 / 1024.0 / 1024.0:#.0} / {globalMemoryStatus.ullTotalPageFile / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine($"Cached: {(standby + modified) / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine($"Paged pool: {pi.KernelPage / 1024.0 / 1024.0 / 1024.0:#.0} GB");
            Console.WriteLine($"Non-paged pool: {pi.KernelNonPage / 1024.0 / 1024.0:#} MB");

            Console.WriteLine();
            Thread.Sleep(1000);
        }
    }
}

static class Unit
{
    public static string MB(this ulong size)
    {
        return $"{(size / 1024 / 1024)}";
    }
}
