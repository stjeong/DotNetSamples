using Microsoft.Diagnostics.Runtime.Interop;
using System;
using System.Runtime.InteropServices;
using WindowsPE;

#pragma warning disable IDE1006 // Naming Styles

namespace NetDbgExt
{
    // DbgShell/ClrMemDiag/Debugger/
    // https://github.com/microsoft/DbgShell/tree/master/ClrMemDiag/Debugger

    // .load C:\...\DotNetSamples\WinConsole\Debugger\NetDbgExt\bin\Debug\x64\NetDbgExt.dll
    // .unload C:\...\DotNetSamples\WinConsole\Debugger\NetDbgExt\bin\Debug\x64\NetDbgExt.dll
    public static class UnmanagedMain
    {
        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Unicode)]
        static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32", CharSet = CharSet.Ansi)]
        static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static uint DebugExtensionInitialize(IntPtr version, IntPtr flags)
        {
            version.WriteValue(DEBUG_EXTENSION_VERSION(1, 0));
            flags.WriteValue(0);

            NativeMethods.OutputDebugString("NetDbgExt.DebugExtensionInitialize");
            return 0;
        }

        public static uint DEBUG_EXTENSION_VERSION(int Major, int Minor)
        {
            return (uint)((((Major) & 0xffff) << 16) | ((Minor) & 0xffff));
        }

        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static void DebugExtensionNotify(DebugNotifySession notify, long argument)
        {
            switch (notify)
            {
                case DebugNotifySession.Active:
                    NativeMethods.OutputDebugString($"DEBUG_NOTIFY_SESSION_ACTIVE: {argument}");
                    break;
                case DebugNotifySession.Inactive:
                    NativeMethods.OutputDebugString($"DEBUG_NOTIFY_SESSION_INACTIVE: {argument}");
                    break;
                case DebugNotifySession.Accessible:
                    NativeMethods.OutputDebugString($"DEBUG_NOTIFY_SESSION_ACCESSIBLE: {argument}");
                    break;
                case DebugNotifySession.InAccessible:
                    NativeMethods.OutputDebugString($"DEBUG_NOTIFY_SESSION_INACCESSIBLE: {argument}");
                    break;
            }
        }

        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static void DebugExtensionUninitialize()
        {
            NativeMethods.OutputDebugString("NetDbgExt.DebugExtensionUninitialize");
        }

        // http://www.sysnet.pe.kr/2/0/12439
        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static uint ubp(IDebugClient pDebugClient, [MarshalAs(UnmanagedType.LPStr)] string args)
        {
            if (!(pDebugClient is IDebugControl dbgControl))
            {
                return 0;
            }

            string[] arg = args.Split('!');
            if (arg.Length != 2)
            {
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"Invalid argument\n");
                return 0;
            }

            string dllPath = arg[0];
            string apiName = arg[1];

            IntPtr ptrDllAddress = LoadLibrary(dllPath);
            if (ptrDllAddress == IntPtr.Zero)
            {
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"DLL not found\n");
                return 0;
            }

            IntPtr ptrApiAddress = GetProcAddress(ptrDllAddress, apiName);
            if (ptrApiAddress == IntPtr.Zero)
            {
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"API not found\n");
                return 0;
            }

            string text = (IntPtr.Size == 4) ? ptrApiAddress.ToInt32().ToString("x") : ptrApiAddress.ToInt64().ToString("x");
            dbgControl.Execute(DEBUG_OUTCTL.THIS_CLIENT, $"bp {text}", DEBUG_EXECUTE.DEFAULT);

            return 0;
        }

        // http://www.sysnet.pe.kr/2/0/11313
        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static uint printdt(IDebugClient pDebugClient, [MarshalAs(UnmanagedType.LPStr)] string args)
        {
            if (!(pDebugClient is IDebugControl dbgControl))
            {
                return 0;
            }

            int result = dbgControl.Evaluate(args, DEBUG_VALUE_TYPE.INT64, out DEBUG_VALUE dbgValue, out _);
            if (result != 0)
            {
                return 0;
            }

            ulong u64Value = dbgValue.I64;
            DateTime dt = Util.ToDateTime((long)u64Value);

            dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"0x{u64Value:x}, 0n{u64Value} ==> {dt}\n");

            return 0;
        }

        // http://www.sysnet.pe.kr/2/0/1198
        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static uint kt(IDebugClient pDebugClient, [MarshalAs(UnmanagedType.LPStr)] string args)
        {
            if (!(pDebugClient is IDebugControl dbgControl))
            {
                return 0;
            }

            if (Util.CheckActiveTarget(dbgControl) == false)
            {
                return 0;
            }

            int result = dbgControl.Evaluate(args, DEBUG_VALUE_TYPE.INT32, out DEBUG_VALUE dbgValue, out _);
            if (result != 0)
            {
                return 0;
            }

            uint threadId = dbgValue.I32;
            IntPtr threadHandle = NativeMethods.OpenThread(ThreadAccessRights.THREAD_ALL_ACCESS, false, (int)threadId);

            if (threadHandle == IntPtr.Zero)
            {
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"OpenThread({threadId}) failed.\n");
                return 0;
            }

            try
            {
                if (NativeMethods.TerminateThread(threadHandle, NativeMethods.DBG_TERMINATE_THREAD) == true)
                {
                    dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"Terminated. Run with 'g' then the thread will be killed\n");
                }
                else
                {
                    dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"TerminateThread: failed.\n");
                }
            }
            finally
            {
                NativeMethods.CloseHandle(threadHandle);
            }

            return 0;
        }

        [DllExport(CallingConvention = CallingConvention.StdCall)]
        public static uint help(IDebugClient pDebugClient, [MarshalAs(UnmanagedType.LPStr)] string _)
        {
            if (!(pDebugClient is IDebugControl dbgControl))
            {
                return 0;
            }

            dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"usage:\n");
            dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"\tkt [threadid]\n");
            dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"\tprintdt [value of DateTime]\n");
            dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"\tubp [dll_path]![api_name]\n");

            return 0;
        }
    }
}
