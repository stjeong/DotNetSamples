using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

#pragma warning disable IDE1006 // Naming Styles

#if _KSOBUILD
namespace KernelStructOffset
#else
namespace WindowsPE
#endif
{
    [StructLayout(LayoutKind.Sequential)]
    public struct StructFieldInfo
    {
        public readonly string Name;
        public readonly int Offset;
        public readonly string Type;

        public StructFieldInfo(int offset, string name, string type)
        {
            Name = name;
            Offset = offset;
            Type = type;
        }
    }

    public class DllOrderLink
    {
        public IntPtr LoadOrderLink;
        public IntPtr MemoryOrderLink;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CodeViewRSDS
    {
        public uint CvSignature;
        public Guid Signature;
        public uint Age;
        public string PdbFileName;

        public string PdbLocalPath
        {
            get
            {
                string fileName = Path.GetFileName(PdbFileName);

                string uid = Signature.ToString("N") + Age;
                return Path.Combine(fileName, uid, fileName);
            }
        }

        public string PdbUriPath
        {
            get
            {
                string fileName = Path.GetFileName(PdbFileName);

                string uid = Signature.ToString("N") + Age;
                return $"{fileName}/{uid}/{fileName}";
            }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_TABLE_ENTRY_INFO
    {
        public short UniqueProcessId;
        public short CreatorBackTraceIndex;
        public byte ObjectType;
        public byte HandleFlags;
        public short HandleValue;
        public IntPtr ObjectPointer;
        public int AccessMask;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX
    {
        public IntPtr ObjectPointer;
        public IntPtr UniqueProcessId;
        public IntPtr HandleValue;
        public uint GrantedAccess;
        public ushort CreatorBackTraceIndex;
        public ushort ObjectTypeIndex;
        public uint HandleAttributes;
        public uint Reserved;

        public int OwnerPid
        {
            get { return UniqueProcessId.ToInt32(); }
        }

        private static Dictionary<string, string> _deviceMap;
        private const int MAX_PATH = 260;
        private const string networkDevicePrefix = "\\Device\\LanmanRedirector\\";

        public override string ToString()
        {
            return $"0x{HandleValue.ToString("x")}(0x{ObjectPointer.ToString("x")})";
        }

        public string GetName(out string handleTypeName)
        {
            IntPtr handle = HandleValue;
            IntPtr dupHandle = IntPtr.Zero;
            handleTypeName = "";
            int ownerPid = UniqueProcessId.ToInt32();

            try
            {
                int addAccessRights = 0;
                dupHandle = DuplicateHandle(handle, addAccessRights);

                if (dupHandle == IntPtr.Zero)
                {
                    return "";
                }

                handleTypeName = GetHandleType(dupHandle);

                switch (handleTypeName)
                {
                    case "EtwRegistration":
                        return "";

                    case "Process":
                        addAccessRights = (int)(ProcessAccessRights.PROCESS_VM_READ | ProcessAccessRights.PROCESS_QUERY_INFORMATION);
                        NativeMethods.CloseHandle(dupHandle);
                        dupHandle = DuplicateHandle(handle, addAccessRights);
                        break;

                    default:
                        break;
                }

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
                            string processName = GetProcessName(ownerPid);
                            int threadId = NativeMethods.GetThreadId(dupHandle);

                            return $"{processName}({ownerPid}): {threadId}";
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
                        devicePath = GetObjectNameFromHandle(dupHandle);

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
                    ProcessAccessRights.PROCESS_QUERY_INFORMATION | ProcessAccessRights.PROCESS_VM_READ, false, ownerPid);

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
            if (processHandle == IntPtr.Zero)
            {
                return "";
            }

            StringBuilder sb = new StringBuilder(4096);
            uint getResult = NativeMethods.GetModuleFileNameEx(processHandle, IntPtr.Zero, sb, sb.Capacity);
            if (getResult == 0)
            {
                return "";
            }

            try
            {
                return Path.GetFileName(sb.ToString());
            }
            catch (System.ArgumentException)
            {
                return "";
            }
        }

        private static bool ConvertDevicePathToDosPath(string devicePath, out string dosPath)
        {
            EnsureDeviceMap();
            int i = devicePath.Length;

            while (i > 0 && (i = devicePath.LastIndexOf('\\', i - 1)) != -1)
            {
                if (_deviceMap.TryGetValue(devicePath.Substring(0, i), out string drive))
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
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            try
            {
                while (true)
                {
                    ret = NativeMethods.NtQueryObject(state.Handle,
                        OBJECT_INFORMATION_CLASS.ObjectNameInformation, ptr, guessSize, out int requiredSize);

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
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            try
            {
                while (true)
                {
                    ret = NativeMethods.NtQueryObject(handle,
                        OBJECT_INFORMATION_CLASS.ObjectTypeInformation, ptr, guessSize, out int requiredSize);

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

        private IntPtr DuplicateHandle(IntPtr targetHandle, int addAccessRights)
        {
            IntPtr currentProcess = NativeMethods.GetCurrentProcess();
            int ownerPid = UniqueProcessId.ToInt32();

            IntPtr targetProcessHandle = IntPtr.Zero;
            IntPtr duplicatedHandle = IntPtr.Zero;

            try
            {
                targetProcessHandle = NativeMethods.OpenProcess(ProcessAccessRights.PROCESS_DUP_HANDLE, false, ownerPid);
                if (targetProcessHandle == IntPtr.Zero)
                {
                    return IntPtr.Zero;
                }

                bool dupResult = NativeMethods.DuplicateHandle(targetProcessHandle, targetHandle, currentProcess,
                    out duplicatedHandle, (int)addAccessRights, false,
                     (addAccessRights == 0) ? DuplicateHandleOptions.DUPLICATE_SAME_ACCESS : 0);
                if (dupResult == true)
                {
                    return duplicatedHandle;
                }

                return IntPtr.Zero;
            }
            finally
            {
                if (targetProcessHandle != IntPtr.Zero)
                {
                    NativeMethods.CloseHandle(targetProcessHandle);
                }
            }
        }

        class FileNameFromHandleState : IDisposable
        {
            private ManualResetEvent _mr;
            private readonly IntPtr _handle;
            private string _fileName;
            private bool _retValue;

            public IntPtr Handle
            {
                get
                {
                    return _handle;
                }
            }

            public string FileName
            {
                get
                {
                    return _fileName;
                }
                set
                {
                    _fileName = value;
                }

            }

            public bool RetValue
            {
                get
                {
                    return _retValue;
                }
                set
                {
                    _retValue = value;
                }
            }

            public FileNameFromHandleState(IntPtr handle)
            {
                _mr = new ManualResetEvent(false);
                this._handle = handle;
            }

            public bool WaitOne(int wait)
            {
                return _mr.WaitOne(wait, false);
            }

            public void Set()
            {
                if (_mr == null)
                {
                    return;
                }

                _mr.Set();
            }

            public void Dispose()
            {
                if (_mr != null)
                {
                    _mr.Close();
                    _mr = null;
                }
            }
        }
    }
}
