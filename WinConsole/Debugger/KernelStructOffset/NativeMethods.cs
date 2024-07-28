using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.ConstrainedExecution;
using System.Runtime.InteropServices;
using System.Text;

#if _KSOBUILD
namespace KernelStructOffset
#else
namespace WindowsPE
#endif
{
    public enum DebugNotifySession
    {
        Active = 0x0,
        Inactive = 0x01,
        Accessible = 0x02,
        InAccessible = 0x03,
    }

    public enum SYM_TYPE
    {
        SymNone = 0,
        SymCoff,
        SymCv,
        SymPdb,
        SymExport,
        SymDeferred,
        SymSym,       // .sym file
        SymDia,
        SymVirtual,
        NumSymTypes
    }

    [Flags]
    public enum SymOpt : uint
    {
        UNDNAME = 0x2,
        DEFERRED_LOADS = 0x4,
        LOAD_LINES = 0x10,
        IGNORE_NT_SYMPATH = 0x1000,
        DEBUG = 0x80000000,
    }

    // Memory Protection Constants
    // https://docs.microsoft.com/en-us/windows/win32/memory/memory-protection-constants
    [Flags]
    public enum PageAccessRights : uint
    {
        NONE = 0x0,
        PAGE_NOACCESS = 0x01,
        PAGE_READONLY = 0x02,
        PAGE_READWRITE = 0x04,
        PAGE_EXECUTE = 0x10,
        PAGE_EXECUTE_READ = 0x20,
        PAGE_EXECUTE_READWRITE = 0x40,
    }

    [Flags]
    public enum MemFreeType : uint
    {
        NONE = 0x0,
        MEM_COALESCE_PLACEHOLDERS = 0x01,
        MEM_PRESERVE_PLACEHOLDER = 0x02,
        MEM_DECOMMIT = 0x00004000,
        MEM_RELEASE = 0x00008000,
    }

    [Flags]
    public enum ProcessAccessRights : uint
    {
        PROCESS_VM_OPERATION = 0x0008,
        PROCESS_VM_READ = 0x10,
        PROCESS_VM_WRITE = 0x0020,
        PROCESS_DUP_HANDLE = 0x00000040,
        PROCESS_QUERY_INFORMATION = 0x0400,
    }

    [Flags]
#pragma warning disable CA2217 // Do not mark enums with FlagsAttribute
    public enum ThreadAccessRights : uint
#pragma warning restore CA2217 // Do not mark enums with FlagsAttribute
    {
        THREAD_QUERY_INFORMATION = 0x00000040,

        // Vista or later
        THREAD_ALL_ACCESS = (uint)0xFFFF | NativeFileAccess.SYNCHRONIZE | NativeFileAccess.STANDARD_RIGHTS_REQUIRED,
        // XP or below
        // THREAD_ALL_ACCESS = (uint)0x3FF | NativeFileAccess.SYNCHRONIZE | NativeFileAccess.STANDARD_RIGHTS_REQUIRED,
    }

    public enum CodeViewSignature : uint
    {
        RSDS = 0x53445352, // SDSR
    }

    public enum DebugDirectoryType : uint
    {
        IMAGE_DEBUG_TYPE_UNKNOWN = 0,
        IMAGE_DEBUG_TYPE_COFF = 1,
        IMAGE_DEBUG_TYPE_CODEVIEW = 2,
        IMAGE_DEBUG_TYPE_FPO = 3,
        IMAGE_DEBUG_TYPE_MISC = 4,
        IMAGE_DEBUG_TYPE_EXCEPTION = 5,
        IMAGE_DEBUG_TYPE_FIXUP = 6,
        IMAGE_DEBUG_TYPE_OMAP_TO_SRC = 7,
        IMAGE_DEBUG_TYPE_OMAP_FROM_SRC = 8,
        IMAGE_DEBUG_TYPE_BORLAND = 9,
        IMAGE_DEBUG_TYPE_RESERVED10 = 10,
        IMAGE_DEBUG_TYPE_CLSID = 11,
        IMAGE_DEBUG_TYPE_VC_FEATURE = 12,
        IMAGE_DEBUG_TYPE_POGO = 13,
        IMAGE_DEBUG_TYPE_ILTCG = 14,
        IMAGE_DEBUG_TYPE_MPX = 15,
        IMAGE_DEBUG_TYPE_REPRO = 16,
    }

    public enum MachineType : ushort
    {
        IMAGE_FILE_MACHINE_I386 = 0x014C,
        IMAGE_FILE_MACHINE_IA64 = 0x0200,
        IMAGE_FILE_MACHINE_AMD64 = 0x8664,
    }

    [Flags]
#pragma warning disable CA2217 // Do not mark enums with FlagsAttribute
    public enum NativeFileAccess : uint
#pragma warning restore CA2217 // Do not mark enums with FlagsAttribute
    {
        FILE_SPECIAL = 0,
        FILE_APPEND_DATA = (0x0004), // file
        FILE_READ_DATA = (0x0001), // file & pipe
        FILE_WRITE_DATA = (0x0002), // file & pipe
        FILE_READ_EA = (0x0008), // file & directory
        FILE_WRITE_EA = (0x0010), // file & directory
        FILE_READ_ATTRIBUTES = (0x0080), // all
        FILE_WRITE_ATTRIBUTES = (0x0100), // all
        DELETE = 0x00010000,
        READ_CONTROL = (0x00020000),
        WRITE_DAC = (0x00040000),
        WRITE_OWNER = (0x00080000),
        SYNCHRONIZE = (0x00100000),
        STANDARD_RIGHTS_REQUIRED = (0x000F0000),
        STANDARD_RIGHTS_READ = (READ_CONTROL),
        STANDARD_RIGHTS_WRITE = (READ_CONTROL),
        STANDARD_RIGHTS_EXECUTE = (READ_CONTROL),
        STANDARD_RIGHTS_ALL = (0x001F0000),
        SPECIFIC_RIGHTS_ALL = (0x0000FFFF),
        FILE_GENERIC_READ = (STANDARD_RIGHTS_READ | FILE_READ_DATA | FILE_READ_ATTRIBUTES | FILE_READ_EA | SYNCHRONIZE),
        FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE | FILE_WRITE_DATA | FILE_WRITE_ATTRIBUTES | FILE_WRITE_EA | FILE_APPEND_DATA | SYNCHRONIZE),
        SPECIAL = 0
    }

    public enum MagicType : ushort
    {
        IMAGE_NT_OPTIONAL_HDR32_MAGIC = 0x10b,
        IMAGE_NT_OPTIONAL_HDR64_MAGIC = 0x20b
    }

    public enum SubSystemType : ushort
    {
        IMAGE_SUBSYSTEM_UNKNOWN = 0,
        IMAGE_SUBSYSTEM_NATIVE = 1,
        IMAGE_SUBSYSTEM_WINDOWS_GUI = 2,
        IMAGE_SUBSYSTEM_WINDOWS_CUI = 3,
        IMAGE_SUBSYSTEM_OS2_CUI = 5,
        IMAGE_SUBSYSTEM_POSIX_CUI = 7,
        IMAGE_SUBSYSTEM_WINDOWS_CE_GUI = 9,
        IMAGE_SUBSYSTEM_EFI_APPLICATION = 10,
        IMAGE_SUBSYSTEM_EFI_BOOT_SERVICE_DRIVER = 11,
        IMAGE_SUBSYSTEM_EFI_RUNTIME_DRIVER = 12,
        IMAGE_SUBSYSTEM_EFI_ROM = 13,
        IMAGE_SUBSYSTEM_XBOX = 14,
        IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION = 16,
    }

    [Flags]
    public enum DllCharacteristicsType : ushort
    {
        RES_0 = 0x0001,
        RES_1 = 0x0002,
        RES_2 = 0x0004,
        RES_3 = 0x0008,
        IMAGE_DLLCHARACTERISTICS_HIGH_ENTROPY_VA = 0x0020,
        IMAGE_DLLCHARACTERISTICS_DYNAMIC_BASE = 0x0040,
        IMAGE_DLLCHARACTERISTICS_FORCE_INTEGRITY = 0x0080,
        IMAGE_DLLCHARACTERISTICS_NX_COMPAT = 0x0100,
        IMAGE_DLLCHARACTERISTICS_NO_ISOLATION = 0x0200,
        IMAGE_DLLCHARACTERISTICS_NO_SEH = 0x0400,
        IMAGE_DLLCHARACTERISTICS_NO_BIND = 0x0800,
        IMAGE_DLLCHARACTERISTICS_APPCONTAINER = 0x1000,
        IMAGE_DLLCHARACTERISTICS_WDM_DRIVER = 0x2000,
        IMAGE_DLLCHARACTERISTICS_TERMINAL_SERVER_AWARE = 0x8000
    }

    public enum NativeFileMode : uint
    {
        CREATE_NEW = 1,
        CREATE_ALWAYS = 2,
        OPEN_EXISTING = 3,
        OPEN_ALWAYS = 4,
        TRUNCATE_EXISTING = 5,
    }

    [Flags]
    public enum NativeFileShare : uint
    {
        NONE = 0,
        FILE_SHARE_READ = 0x00000001,
        FILE_SHARE_WRITE = 0x00000002,
        FILE_SHARE_DEELETE = 0x00000004,
    }

    [Flags]
    public enum NativeFileFlag : uint
    {
        FILE_ATTRIBUTE_READONLY = 0x00000001,
        FILE_ATTRIBUTE_HIDDEN = 0x00000002,
        FILE_ATTRIBUTE_SYSTEM = 0x00000004,
        FILE_ATTRIBUTE_DIRECTORY = 0x00000010,
        FILE_ATTRIBUTE_ARCHIVE = 0x00000020,
        FILE_ATTRIBUTE_DEVICE = 0x00000040,
        FILE_ATTRIBUTE_NORMAL = 0x00000080,
        FILE_ATTRIBUTE_TEMPORARY = 0x00000100,
        FILE_ATTRIBUTE_SPARSE_FILE = 0x00000200,
        FILE_ATTRIBUTE_REPARSE_POINT = 0x00000400,
        FILE_ATTRIBUTE_COMPRESSED = 0x00000800,
        FILE_ATTRIBUTE_OFFLINE = 0x00001000,
        FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 0x00002000,
        FILE_ATTRIBUTE_ENCRYPTED = 0x00004000,
        FILE_FLAG_WRITE_THROUGH = 0x80000000,
        FILE_FLAG_OVERLAPPED = 0x40000000,
        FILE_FLAG_NO_BUFFERING = 0x20000000,
        FILE_FLAG_RANDOM_ACCESS = 0x10000000,
        FILE_FLAG_SEQUENTIAL_SCAN = 0x08000000,
        FILE_FLAG_DELETE_ON_CLOSE = 0x04000000,
        FILE_FLAG_BACKUP_SEMANTICS = 0x02000000,
        FILE_FLAG_POSIX_SEMANTICS = 0x01000000,
        FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000,
        FILE_FLAG_OPEN_NO_RECALL = 0x00100000,
        FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000,
    }

    [Flags()]
    public enum AllocationType : uint
    {
        COMMIT = 0x1000,
        RESERVE = 0x2000,
        RESET = 0x80000,
        LARGE_PAGES = 0x20000000,
        PHYSICAL = 0x400000,
        TOP_DOWN = 0x100000,
        WRITE_WATCH = 0x200000
    }

    public enum NT_STATUS
    {
        STATUS_SUCCESS = 0x00000000,
        STATUS_BUFFER_OVERFLOW = unchecked((int)0x80000005L),
        STATUS_INVALID_INFO_CLASS = unchecked((int)0x80000003L),
        STATUS_INFO_LENGTH_MISMATCH = unchecked((int)0xC0000004L),
        STATUS_INVALID_HANDLE = unchecked((int)0xC0000008L),
        STATUS_INVALID_PARAMETER = unchecked((int)0xC000000DL),
        STATUS_ACCESS_DENIED = unchecked((int)0xC0000022L),
        STATUS_BUFFER_TOO_SMALL = unchecked((int)0xC0000023L),
        STATUS_NOT_SUPPORTED = unchecked((int)0xC00000BBL),
    }

    public enum SYSTEM_INFORMATION_CLASS
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
        SystemLookasideInformation = 45,
        SystemExtendedHandleInformation = 64
    }

    public enum PROCESS_INFORMATION_CLASS
    {
        ProcessHandleInformation = 51,
    };

    public enum OBJECT_INFORMATION_CLASS
    {
        ObjectBasicInformation = 0,
        ObjectNameInformation = 1,
        ObjectTypeInformation = 2,
        ObjectAllTypesInformation = 3,
        ObjectHandleFlagInformation = 4,
        ObjectSessionInformation = 5,
    }

    [Flags]
    public enum DuplicateHandleOptions
    {
        DUPLICATE_CLOSE_SOURCE = 0x1,
        DUPLICATE_SAME_ACCESS = 0x2
    }

    [Flags()]
    public enum MemoryProtection : uint
    {
        EXECUTE = 0x10,
        EXECUTE_READ = 0x20,
        EXECUTE_READWRITE = 0x40,
        EXECUTE_WRITECOPY = 0x80,
        NOACCESS = 0x01,
        READONLY = 0x02,
        READWRITE = 0x04,
        WRITECOPY = 0x08,
        GUARD_Modifierflag = 0x100,
        NOCACHE_Modifierflag = 0x200,
        WRITECOMBINE_Modifierflag = 0x400
    }

    public static class NativeMethods
    {
        internal const int DBG_TERMINATE_THREAD = 0x40010003;
        private const int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;
        private static readonly StringBuilder m_sbSysMsg = new StringBuilder(1024);

        public static string FormatSystemMessage(int dwMessageId)
        {
            lock (m_sbSysMsg)
            {
                int ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM,
                                        IntPtr.Zero,
                                        dwMessageId,
                                        CultureInfo.CurrentCulture.LCID,
                                        m_sbSysMsg,
                                        m_sbSysMsg.Capacity - 1,
                                        IntPtr.Zero);

                if (ret > 0) return m_sbSysMsg.ToString(0, ret);
                return null;
            }
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern SafeFileHandle CreateFile(string fileName,
                                                       [MarshalAs(UnmanagedType.U4)] NativeFileAccess fileAccess,
                                                       [MarshalAs(UnmanagedType.U4)] NativeFileShare fileShare,
                                                       IntPtr securityAttributes,
                                                       [MarshalAs(UnmanagedType.U4)] NativeFileMode creationDisposition,
                                                       NativeFileFlag flags,
                                                       IntPtr template);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        internal static extern bool DeviceIoControl(SafeHandle hDevice,
                                                  uint IoControlCode,
                                                  [MarshalAs(UnmanagedType.AsAny), In] object InBuffer,
                                                  int nInBufferSize,
                                                  IntPtr OutBuffer,
                                                  int nOutBufferSize,
                                                  out int pBytesReturned,
                                                  IntPtr pOverlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        internal static extern bool DeviceIoControl(SafeHandle hDevice,
                                                  uint IoControlCode,
                                                  [MarshalAs(UnmanagedType.AsAny), In] object InBuffer,
                                                  int nInBufferSize,
                                                  [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 6), Out] byte[] OutBuffer,
                                                  int nOutBufferSize,
                                                  out int pBytesReturned,
                                                  IntPtr Overlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi, EntryPoint = "DeviceIoControl")]
        internal static extern bool DeviceIoControlAsObject(SafeHandle hDevice,
                                                          uint IoControlCode,
                                                          [MarshalAs(UnmanagedType.AsAny), In] object InBuffer,
                                                          int nInBufferSize,
                                                          [MarshalAs(UnmanagedType.AsAny), Out] object OutBuffer,
                                                          int nOutBufferSize,
                                                          ref int pBytesReturned,
                                                          IntPtr Overlapped);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int FormatMessage(int dwFlags,
                                                IntPtr lpSource,
                                                int dwMessageId,
                                                int dwLanguageId,
                                                [Out] StringBuilder lpBuffer,
                                                int nSize,
                                                IntPtr lpArguments);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Ansi)]
        internal static extern bool GetOverlappedResult(SafeHandle hDevice, IntPtr lpOverlapped, out int lpNumberOfBytesTransferred, bool bWait);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr VirtualAlloc(IntPtr lpAddress, UIntPtr dwSize, AllocationType flAllocationType,
            MemoryProtection flProtect);

        [DllImport("kernel32")]
        internal static extern bool VirtualFree(IntPtr lpAddress, UInt32 dwSize, UInt32 dwFreeType);

        [DllImport("ntdll.dll")]
        internal static extern IntPtr NtCurrentTeb();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWow64Process([In] IntPtr hProcess, [Out] out bool lpSystemInfo);

        [DllImport("psapi.dll", SetLastError = true, CharSet = CharSet.Unicode)]
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
        internal static extern NT_STATUS NtQueryInformationProcess(
            [In] IntPtr ProcessHandle,
            [In] PROCESS_INFORMATION_CLASS ProcessInformationClass,
            [In] IntPtr ProcessInformation,
            [In] int ProcessInformationLength,
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

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr OpenProcess(
            [In] ProcessAccessRights dwDesiredAccess,
            [In, MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            [In] int dwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr OpenThread(
            [In] ThreadAccessRights dwDesiredAccess,
            [In, MarshalAs(UnmanagedType.Bool)] bool bInheritHandle,
            [In] int dwProcessId);

        [DllImport("kernel32.dll")]
        internal static extern bool TerminateThread(IntPtr hThread, uint dwExitCode);

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
        internal static extern bool CloseHandle([In] IntPtr hObject);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern int QueryDosDevice(
            [In] string lpDeviceName,
            [Out] StringBuilder lpTargetPath,
            [In] int ucchMax);

        public const int MAX_PATH = 260;

        [DllImport("kernel32.dll", SetLastError = true)]

        internal static extern void GetSystemInfo(ref SYSTEM_INFO Info);

        [DllImport("Dbghelp.dll", SetLastError = true)]
        internal static extern uint SymGetOptions();

        [DllImport("Dbghelp.dll", SetLastError = true)]
        internal static extern uint SymSetOptions(uint SymOptions);

        [DllImport("Dbghelp.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static extern bool SymInitialize(IntPtr hProcess, string UserSearchPath, bool fInvadeProcess);

        [DllImport("Dbghelp.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        internal static unsafe extern ulong SymLoadModuleEx(IntPtr hProcess,
          IntPtr hFile,
          string ImageName,
          string ModuleName,
          long BaseOfDll,
          uint DllSize,
          _MODLOAD_DATA* Data,
          uint Flags);

        [DllImport("kernel32", SetLastError = true, CharSet = CharSet.Ansi, BestFitMapping = false)]
        internal static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("Dbghelp.dll", SetLastError = true)]
        internal static unsafe extern bool SymGetModuleInfo64(IntPtr hProcess, ulong qwAddr, ref _IMAGEHLP_MODULE64 ModuleInfo);

        [DllImport("dbghelp.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SymEnumSymbolsW")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SymEnumSymbols(IntPtr hProcess,
                                                    ulong modBase,
                                                    string mask,
                                                    native_SYM_ENUMERATESYMBOLS_CALLBACK callback,
                                                    IntPtr pUserContext);

        [DllImport("dbghelp.dll", SetLastError = true, CharSet = CharSet.Unicode, EntryPoint = "SymEnumTypesW")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SymEnumTypes(IntPtr hProcess, ulong BaseOfDll,
            native_SYM_ENUMERATESYMBOLS_CALLBACK callback, IntPtr UserContext);

        [return: MarshalAs(UnmanagedType.Bool)]
        internal delegate bool native_SYM_ENUMERATESYMBOLS_CALLBACK( /* SYMBOL_INFO* */ IntPtr symInfo,
                                                                                     uint symbolSize,
                                                                                     IntPtr pUserContext);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("dbghelp.dll", SetLastError = true, EntryPoint = "SymCleanup")]
        internal static extern bool SymCleanup(IntPtr hProcess);

        [DllImport("user32.dll")]
        internal static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, ref IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        internal static extern void OutputDebugString(string lpOutputString);

        [DllImport("kernel32.dll")]
        internal static extern bool VirtualProtectEx(IntPtr hProcess, IntPtr lpAddress, UIntPtr dwSize, PageAccessRights flNewProtect, out PageAccessRights lpflOldProtect);

        [DllImport("kernel32.dll")]
        internal static extern bool FlushInstructionCache(IntPtr hProcess, IntPtr lpBaseAddress, UIntPtr dwSize);
    }
}
