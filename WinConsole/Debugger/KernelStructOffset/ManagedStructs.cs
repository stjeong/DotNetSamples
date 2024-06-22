using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
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

        public static bool operator ==(StructFieldInfo t1, StructFieldInfo t2)
        {
            return t1.Equals(t2);
        }

        public static bool operator !=(StructFieldInfo t1, StructFieldInfo t2)
        {
            return !t1.Equals(t2);
        }

        public override bool Equals(object obj)
        {
            StructFieldInfo target = (StructFieldInfo)obj;

            if (target.Name == this.Name)
            {
                return true;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return this.Name?.GetHashCode() ?? 0;
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

        public static bool operator ==(CodeViewRSDS t1, CodeViewRSDS t2)
        {
            return t1.Equals(t2);
        }

        public static bool operator !=(CodeViewRSDS t1, CodeViewRSDS t2)
        {
            return !t1.Equals(t2);
        }

        public override bool Equals(object obj)
        {
            CodeViewRSDS target = (CodeViewRSDS)obj;

            if (target.Signature == this.Signature && target.Age == this.Age)
            {
                return true;
            }

            return false;
        }

        public override int GetHashCode()
        {
            return Signature.GetHashCode() + Age.GetHashCode();
        }

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
    public struct _PROCESS_HANDLE_TABLE_ENTRY_INFO
    {
        public IntPtr HandleValue;
        public UIntPtr HandleCount;
        public UIntPtr PointerCount;
        public uint GrantedAccess;
        public uint ObjectTypeIndex;
        public uint HandleAttributes;
        public uint Reserved;

        public string GetName(int ownerPid, out string handleTypeName)
        {
            return _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX.GetName(HandleValue, ownerPid, out handleTypeName);
        }
    }

    public struct _PROCESS_HANDLE_SNAPSHOT_INFORMATION
    {
        public UIntPtr HandleCount;
        public UIntPtr Reserved;
        public _PROCESS_HANDLE_TABLE_ENTRY_INFO Handles; /* Handles[0] */

        public int NumberOfHandles
        {
            get { return (int)HandleCount.ToUInt32(); }
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

        public static bool operator ==(_SYSTEM_HANDLE_TABLE_ENTRY_INFO t1, _SYSTEM_HANDLE_TABLE_ENTRY_INFO t2)
        {
            return t1.Equals(t2);
        }

        public static bool operator !=(_SYSTEM_HANDLE_TABLE_ENTRY_INFO t1, _SYSTEM_HANDLE_TABLE_ENTRY_INFO t2)
        {
            return !t1.Equals(t2);
        }

        public override int GetHashCode()
        {
            return this.ObjectPointer.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            _SYSTEM_HANDLE_TABLE_ENTRY_INFO target = (_SYSTEM_HANDLE_TABLE_ENTRY_INFO)obj;

            if (target.ObjectPointer == this.ObjectPointer)
            {
                return true;
            }

            return false;
        }
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
            return GetName(HandleValue, UniqueProcessId.ToInt32(), out handleTypeName);
        }

        public static string GetName(IntPtr handleValue, int ownerPid, out string handleTypeName)
        {
            IntPtr handle = handleValue;
            IntPtr dupHandle = IntPtr.Zero;
            handleTypeName = "";

            try
            {
                int addAccessRights = 0;
                dupHandle = DuplicateHandle(ownerPid, handle, addAccessRights);

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
                        dupHandle = DuplicateHandle(ownerPid, handle, addAccessRights);
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

        private static string GetProcessName(int ownerPid)
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

        private static string GetProcessName(IntPtr processHandle)
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

        private static string GetObjectNameFromHandle(IntPtr handle)
        {
            using (FileNameFromHandleState f = new FileNameFromHandleState(handle))
            {
                ThreadPool.QueueUserWorkItem(GetObjectNameFromHandleFunc, f);
                f.WaitOne(16);
                return f.FileName;
            }
        }

        private static void GetObjectNameFromHandleFunc(object obj)
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

                state.Set();
            }
        }

        private static string GetHandleType(IntPtr handle)
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

        private static IntPtr DuplicateHandle(int ownerPid, IntPtr targetHandle, int addAccessRights)
        {
            IntPtr currentProcess = NativeMethods.GetCurrentProcess();

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

        public static bool operator ==(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX t1, _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX t2)
        {
            return t1.Equals(t2);
        }

        public static bool operator !=(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX t1, _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX t2)
        {
            return !t1.Equals(t2);
        }

        public override int GetHashCode()
        {
            return this.ObjectPointer.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX target = (_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX)obj;

            if (target.ObjectPointer == this.ObjectPointer)
            {
                return true;
            }

            return false;
        }
    }

    public class ImageResourceEntryId
    {
        public string Name { get; set; } = "";
        public ushort Id { get; set; } = 0;

        public static bool operator == (ImageResourceEntryId t1, ResourceTypeId t2)
        {
            return string.IsNullOrEmpty(t1.Name) && t1.Id == (ushort)t2;
        }

        public static bool operator !=(ImageResourceEntryId t1, ResourceTypeId t2)
        {
            return !t1.Equals(t2);
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(Name) == false)
            {
                return Name;
            }

            return $"{Id}";
        }
    }

    public class _IMAGE_RESOURCE_DATA_ENTRY
    {
        IMAGE_RESOURCE_DATA_ENTRY entry;

        public uint CodePage => entry.CodePage;
        public uint OffsetToData => entry.OffsetToData;
        public uint Size => entry.Size;

        public _IMAGE_RESOURCE_DATA_ENTRY(IMAGE_RESOURCE_DATA_ENTRY entry)
        {
            this.entry = entry;
        }

        public override string ToString()
        {
            return $"{entry}";
        }
    }

    public class _IMAGE_RESOURCE_DIRECTORY_ENTRY
    {
        ImageResourceEntryId _id;
        public ImageResourceEntryId Id => _id;

        public _IMAGE_RESOURCE_DIRECTORY Next;
        public _IMAGE_RESOURCE_DATA_ENTRY Data;

        public _IMAGE_RESOURCE_DIRECTORY_ENTRY(IMAGE_RESOURCE_DIRECTORY_ENTRY entry, IntPtr directoryPtr)
        {
            _id = new ImageResourceEntryId();

            if (entry.Name.NameIsString == true)
            {
                IntPtr namePos = IntPtr.Add(directoryPtr, (int)entry.Name.NameOffset);
                int nameLength = Marshal.ReadInt16(namePos);
                namePos = IntPtr.Add(namePos, sizeof(short));
                unsafe
                {
                    _id.Name = new string((char*)namePos.ToPointer(), 0, nameLength);
                }
            }
            else
            {
                _id.Id = entry.Name.Id;
            }
        }

        public override string ToString()
        {
            if (Next != null)
            {
                return $"[dir]: {_id}, LinkTo: {Next}";
            }
            else
            {
                return $"[entry]: Id={_id}, {Data}";
            }
        }
    }

    public class _IMAGE_RESOURCE_DIRECTORY
    {
        _IMAGE_RESOURCE_DIRECTORY_ENTRY[] entries;
        public _IMAGE_RESOURCE_DIRECTORY_ENTRY[] Entries => entries;

        public _IMAGE_RESOURCE_DIRECTORY(int numberOfEntries)
        {
            this.entries = new _IMAGE_RESOURCE_DIRECTORY_ENTRY[numberOfEntries];
        }

        public override string ToString()
        {
            return $"Entries: {this.entries.Length}";
        }

        public _IMAGE_RESOURCE_DATA_ENTRY FindLcidEntry(int lcid)
        {
            foreach (var item in Entries)
            {
                if (item.Data != null && item.Data.CodePage == lcid)
                {
                    return item.Data;
                }
            }

            if (this.Entries.Length > 0)
            {
                return this.Entries[0].Data;
            }

            return null;
        }
    }

    public class VS_VERSION_INFO
    {
        public ushort wLength;
        public ushort wValueLength;
        public ushort wType;
        public string szKey;
        public VS_FIXEDFILEINFO FileInfo;
        public VarFileInfo VarFileInfo;
        public StringFileInfo StringFileInfo;

        public static unsafe VS_VERSION_INFO Parse(IntPtr ptr, uint length)
        {
            VS_VERSION_INFO item = new VS_VERSION_INFO();

            item.wLength = ptr.ReadUInt16(0);
            item.wValueLength = ptr.ReadUInt16(2);
            item.wType = ptr.ReadUInt16(4);
            item.szKey = ptr.ReadString(6);

            return item;
        }
    }

    public class StringFileInfo
    {
        public string szKey;
        public StringTable[] Children;
    }

    public class StringTable
    {
        public string szKey;
        public StringTableString[] Children;
    }

    public class StringTableString
    {
        public string szKey;
        public string Value;
    }

    public class VarFileInfo
    {
        public string szKey;
        public VarFileInfoVar[] Children;
    }

    public class VarFileInfoVar
    {
        public string szKey;
        public ushort[] Value;
    }
}
