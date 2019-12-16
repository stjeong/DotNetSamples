using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace EnumHandles
{
    internal enum NT_STATUS
    {
        STATUS_SUCCESS = 0x00000000,
        STATUS_BUFFER_OVERFLOW = unchecked((int)0x80000005L),
        STATUS_INFO_LENGTH_MISMATCH = unchecked((int)0xC0000004L),
        STATUS_INVALID_HANDLE = unchecked((int)0xC0000008L),
        STATUS_INVALID_PARAMETER = unchecked((int)0xC000000DL),
        STATUS_ACCESS_DENIED = unchecked((int)0xC0000022L),
        STATUS_BUFFER_TOO_SMALL = unchecked((int)0xC0000023L),
        STATUS_NOT_SUPPORTED = unchecked((int)0xC00000BBL),
    }

    internal enum SYSTEM_INFORMATION_CLASS
    {
        SystemBasicInformation = 0,
        SystemPerformanceInformation = 2,
        SystemTimeOfDayInformation = 3,
        SystemProcessInformation = 5,
        SystemProcessorPerformanceInformation = 8,
        SystemHandleInformation = 16,
        SystemInterruptInformation = 23,
        SystemExceptionInformation = 33,
        SystemRegistryQuotaInformation = 37,
        SystemLookasideInformation = 45
    }

    internal enum OBJECT_INFORMATION_CLASS
    {
        ObjectBasicInformation = 0,
        ObjectNameInformation = 1,
        ObjectTypeInformation = 2,
        ObjectAllTypesInformation = 3,
        ObjectHandleFlagInformation = 4,
        ObjectSessionInformation = 5,
    }

    [Flags]
    internal enum ProcessAccessRights
    {
        PROCESS_VM_READ = 0x10,
        PROCESS_DUP_HANDLE = 0x00000040,
        PROCESS_QUERY_INFORMATION = 0x0400,
    }

    [Flags]
    internal enum ThreadAccessRights
    {
        THREAD_QUERY_INFORMATION = 0x00000040,
    }

    [Flags]
    internal enum DuplicateHandleOptions
    {
        DUPLICATE_CLOSE_SOURCE = 0x1,
        DUPLICATE_SAME_ACCESS = 0x2
    }

    public struct UNICODE_STRING
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;

        public string GetText()
        {
            if (Buffer == IntPtr.Zero)
            {
                return "";
            }

            return Marshal.PtrToStringAuto(Buffer, Length / 2);
        }
    }

    public struct GENERIC_MAPPING
    {
        public uint GenericRead;
        public uint GenericWrite;
        public uint GenericExecute;
        public uint GenericAll;
    }

    public struct OBJECT_NAME_INFORMATION
    {
        public UNICODE_STRING Name;
    }

    public struct OBJECT_TYPE_INFORMATION
    {
        public UNICODE_STRING Name;
        public uint TotalNumberOfObjects;
        public uint TotalNumberOfHandles;
        public uint TotalPagedPoolUsage;
        public uint TotalNonPagedPoolUsage;
        public uint TotalNamePoolUsage;
        public uint TotalHandleTableUsage;
        public uint HighWaterNumberOfObjects;
        public uint HighWaterNumberOfHandles;
        public uint HighWaterPagedPoolUsage;
        public uint HighWaterNonPagedPoolUsage;
        public uint HighWaterNamePoolUsage;
        public uint HighWaterHandleTableUsage;
        public uint InvalidAttributes;
        public GENERIC_MAPPING GenericMapping;
        public uint ValidAccess;
        public byte SecurityRequired;
        public byte MaintainHandleCount;
        public ushort MaintainTypeList;

        /*
enum _POOL_TYPE
{
	NonPagedPool,
	PagedPool,
	NonPagedPoolMustSucceed,
	DontUseThisType,
	NonPagedPoolCacheAligned,
	PagedPoolCacheAligned,
	NonPagedPoolCacheAlignedMustS
}
        */

        public int PoolType;
        public uint PagedPoolUsage;
        public uint NonPagedPoolUsage;
    }

    public struct SYSTEM_HANDLE_INFORMATION
    {
        public int HandleCount;
        public SYSTEM_HANDLE_ENTRY Handles; /* Handles[0] */
    }

    internal static class NativeMethods
    {
        [DllImport("psapi.dll")]
        internal static extern uint GetModuleFileNameEx(IntPtr hProcess, 
            IntPtr hModule, 
            [Out] StringBuilder lpBaseName, 
            [In] [MarshalAs(UnmanagedType.U4)] 
            int nSize);

        [DllImport("ntdll.dll")]
        internal static extern NT_STATUS NtQuerySystemInformation(
            [In] SYSTEM_INFORMATION_CLASS SystemInformationClass,
            [In] IntPtr SystemInformation,
            [In] int SystemInformationLength,
            [Out] out int ReturnLength);

        [DllImport("ntdll.dll")]
        internal static extern NT_STATUS NtQueryObject(
            [In] IntPtr Handle,
            [In] OBJECT_INFORMATION_CLASS ObjectInformationClass,
            [In] IntPtr ObjectInformation,
            [In] int ObjectInformationLength,
            [Out] out int ReturnLength);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int GetThreadId(IntPtr threadHandle);


        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr OpenProcess(
            [In] ProcessAccessRights dwDesiredAccess,
            [In, MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            [In] int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DuplicateHandle(
            [In] IntPtr hSourceProcessHandle,
            [In] IntPtr hSourceHandle,
            [In] IntPtr hTargetProcessHandle,
            [Out] out IntPtr lpTargetHandle,
            [In] int dwDesiredAccess,
            [In, MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            [In] DuplicateHandleOptions dwOptions);

        [DllImport("kernel32.dll")]
        internal static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll")]
        internal static extern int GetCurrentProcessId();

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int GetProcessId(
            [In] IntPtr Process);

        [ReliabilityContract(Consistency.WillNotCorruptState, Cer.Success)]
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(
            [In] IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int QueryDosDevice(
            [In] string lpDeviceName,
            [Out] StringBuilder lpTargetPath,
            [In] int ucchMax);
    }
}
