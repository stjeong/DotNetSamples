using KernelStructOffset;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace KernelStructOffset
{
    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEM_HANDLE_ENTRY
    {
        public int OwnerPid;
        public byte ObjectType;
        public byte HandleFlags;
        public short HandleValue;
        public IntPtr ObjectPointer;
        public int AccessMask;

        private const int STANDARD_RIGHTS_ALL = 0x001F0000;
        private static Dictionary<string, string> _deviceMap;
        private const int MAX_PATH = 260;
        private const string networkDevicePrefix = "\\Device\\LanmanRedirector\\";

        public override string ToString()
        {
            return $"0x{HandleValue.ToString("x")}(0x{ObjectPointer.ToString("x")})";
        }

        public string GetName(out string handleTypeName)
        {
            IntPtr handle = new IntPtr(HandleValue);
            IntPtr dupHandle = IntPtr.Zero;
            handleTypeName = "";

            try
            {
                dupHandle = DuplicateHandle(handle, out bool isRemoteHandle);

                if (isRemoteHandle == true && dupHandle == IntPtr.Zero)
                {
                    return "";
                }

                if (dupHandle != IntPtr.Zero)
                {
                    handle = dupHandle;
                }

                handleTypeName = GetHandleType(handle);
                string devicePath = "";

                switch (handleTypeName)
                {
                    case "Process":
                        {
                            string processName = GetProcessName(dupHandle);
                            int processId = NativeMethods.GetProcessId(dupHandle);

                            return $"{processName}({processId})";
                        }

                    case "Thread":
                        {
                            string processName = GetProcessName(this.OwnerPid);
                            int threadId = NativeMethods.GetThreadId(dupHandle);

                            return $"{processName}({this.OwnerPid}): {threadId}";
                        }

                    case "Directory":
                    case "ALPC Port":
                    case "Desktop":
                    case "Event":
                    case "Key":
                    case "Mutant":
                    case "Section":
                    case "Semaphore":
                    case "Token":
                    case "WindowStation":
                    case "File":
                        devicePath = GetObjectNameFromHandle(handle);

                        if (string.IsNullOrEmpty(devicePath) == true)
                        {
                            return "";
                        }

                        string dosPath;
                        if (ConvertDevicePathToDosPath(devicePath, out dosPath))
                        {
                            return dosPath;
                        }

                        return devicePath;
                }
            }
            finally
            {
                if (dupHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(dupHandle);
                }
            }

            return "";
        }

        private string GetProcessName(int ownerPid)
        {
            IntPtr processHandle = IntPtr.Zero;

            try
            {
                processHandle = NativeMethods.OpenProcess(
                    ProcessAccessRights.PROCESS_QUERY_INFORMATION | ProcessAccessRights.PROCESS_VM_READ, false, OwnerPid);

                if (processHandle == IntPtr.Zero)
                {
                    return "";
                }

                return GetProcessName(processHandle);
            }
            finally
            {
                if (processHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(processHandle);
                }
            }
        }

        private string GetProcessName(IntPtr processHandle)
        {
            StringBuilder sb = new StringBuilder(4096);
            NativeMethods.GetModuleFileNameEx(processHandle, IntPtr.Zero, sb, 4096);

            try
            {
                return Path.GetFileName(sb.ToString());
            }
            catch (System.ArgumentException)
            {
                return sb.ToString();
            }
        }

        private static bool ConvertDevicePathToDosPath(string devicePath, out string dosPath)
        {
            EnsureDeviceMap();
            int i = devicePath.Length;

            while (i > 0 && (i = devicePath.LastIndexOf('\\', i - 1)) != -1)
            {
                string drive;
                if (_deviceMap.TryGetValue(devicePath.Substring(0, i), out drive))
                {
                    dosPath = string.Concat(drive, devicePath.Substring(i));
                    return dosPath.Length != 0;
                }
            }

            dosPath = string.Empty;
            return false;
        }

        private static void EnsureDeviceMap()
        {
            if (_deviceMap == null)
            {
                Dictionary<string, string> localDeviceMap = BuildDeviceMap();
                Interlocked.CompareExchange<Dictionary<string, string>>(ref _deviceMap, localDeviceMap, null);
            }
        }

        private static Dictionary<string, string> BuildDeviceMap()
        {
            string[] logicalDrives = Environment.GetLogicalDrives();

            Dictionary<string, string> localDeviceMap = new Dictionary<string, string>(logicalDrives.Length);
            StringBuilder lpTargetPath = new StringBuilder(MAX_PATH);

            foreach (string drive in logicalDrives)
            {
                string lpDeviceName = drive.Substring(0, 2);
                NativeMethods.QueryDosDevice(lpDeviceName, lpTargetPath, MAX_PATH);
                localDeviceMap.Add(NormalizeDeviceName(lpTargetPath.ToString()), lpDeviceName);
            }

            localDeviceMap.Add(networkDevicePrefix.Substring(0, networkDevicePrefix.Length - 1), "\\");
            return localDeviceMap;
        }

        private static string NormalizeDeviceName(string deviceName)
        {
            if (string.Compare(deviceName, 0, networkDevicePrefix, 0, networkDevicePrefix.Length, StringComparison.InvariantCulture) == 0)
            {
                string shareName = deviceName.Substring(deviceName.IndexOf('\\', networkDevicePrefix.Length) + 1);
                return string.Concat(networkDevicePrefix, shareName);
            }
            return deviceName;
        }

        private string GetObjectNameFromHandle(IntPtr handle)
        {
            using (FileNameFromHandleState f = new FileNameFromHandleState(handle))
            {
                ThreadPool.QueueUserWorkItem(GetObjectNameFromHandle, f);
                f.WaitOne(10);
                return f.FileName;
            }
        }

        private void GetObjectNameFromHandle(object obj)
        {
            FileNameFromHandleState state = obj as FileNameFromHandleState;

            int guessSize = 1024;
            int requiredSize = 0;
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            try
            {
                while (true)
                {
                    ret = NativeMethods.NtQueryObject(state.Handle,
                        OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, guessSize, out requiredSize);

                    if (ret == NT_STATUS.STATUS_INFO_LENGTH_MISMATCH)
                    {
                        Marshal.FreeHGlobal(ptr);
                        guessSize = requiredSize;
                        ptr = Marshal.AllocHGlobal(guessSize);
                        continue;
                    }

                    if (ret == NT_STATUS.STATUS_SUCCESS)
                    {
                        OBJECT_NAME_INFORMATION oti = (OBJECT_NAME_INFORMATION)Marshal.PtrToStructure(ptr, typeof(OBJECT_NAME_INFORMATION));
                        state.FileName = oti.Name.GetText();
                        return;
                    }

                    break;
                }
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(ptr);
                }
            }
        }

        private string GetHandleType(IntPtr handle)
        {
            int guessSize = 1024;
            int requiredSize = 0;
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            try
            {
                while (true)
                {
                    ret = NativeMethods.NtQueryObject(handle,
                        OBJECT_INFORMATION_CLASS.ObjectTypeInformation, ptr, guessSize, out requiredSize);

                    if (ret == NT_STATUS.STATUS_INFO_LENGTH_MISMATCH)
                    {
                        Marshal.FreeHGlobal(ptr);
                        guessSize = requiredSize;
                        ptr = Marshal.AllocHGlobal(guessSize);
                        continue;
                    }

                    if (ret == NT_STATUS.STATUS_SUCCESS)
                    {
                        OBJECT_TYPE_INFORMATION oti = (OBJECT_TYPE_INFORMATION)Marshal.PtrToStructure(ptr, typeof(OBJECT_TYPE_INFORMATION));
                        return oti.Name.GetText();
                    }

                    break;
                }
            }
            finally
            {
                if (ptr != IntPtr.Zero)
                {
                    Marshal.FreeHGlobal(ptr);
                }
            }

            return "(unknown)";
        }

        private IntPtr DuplicateHandle(IntPtr handle, out bool isRemoteHandle)
        {
            IntPtr currentProcess = NativeMethods.GetCurrentProcess();
            int processId = NativeMethods.GetCurrentProcessId();
            isRemoteHandle = (OwnerPid != processId);

            IntPtr processHandle = IntPtr.Zero;
            IntPtr objectHandle = IntPtr.Zero;

            try
            {
                if (isRemoteHandle == true)
                {
                    processHandle = NativeMethods.OpenProcess(
                        ProcessAccessRights.PROCESS_DUP_HANDLE, false, OwnerPid);
                    if (processHandle == IntPtr.Zero)
                    {
                        return IntPtr.Zero;
                    }

                    bool queryResult = NativeMethods.DuplicateHandle(processHandle, handle, currentProcess,
                        out objectHandle, STANDARD_RIGHTS_ALL, false, DuplicateHandleOptions.DUPLICATE_SAME_ACCESS);
                    if (queryResult == true)
                    {
                        return objectHandle;
                    }
                }

                return IntPtr.Zero;
            }
            finally
            {
                if (processHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(processHandle);
                }
            }
        }
    }
}
