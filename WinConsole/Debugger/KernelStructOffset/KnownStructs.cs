using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

#if _KSOBUILD
namespace KernelStructOffset
#else
namespace WindowsPE
#endif
{
    public enum WM_MESSAGE
    {
        WM_COPYDATA = 0x004A,
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct COPYDATASTRUCT
    {
        public IntPtr dwData;    // Any value the sender chooses.  Perhaps its main window handle?
        public int cbData;       // The count of bytes in the message.
        public IntPtr lpData;    // The address of the message.
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct IMAGE_DOS_HEADER
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public char[] e_magic;       // Magic number
        public UInt16 e_cblp;    // Bytes on last page of file
        public UInt16 e_cp;      // Pages in file
        public UInt16 e_crlc;    // Relocations
        public UInt16 e_cparhdr;     // Size of header in paragraphs
        public UInt16 e_minalloc;    // Minimum extra paragraphs needed
        public UInt16 e_maxalloc;    // Maximum extra paragraphs needed
        public UInt16 e_ss;      // Initial (relative) SS value
        public UInt16 e_sp;      // Initial SP value
        public UInt16 e_csum;    // Checksum
        public UInt16 e_ip;      // Initial IP value
        public UInt16 e_cs;      // Initial (relative) CS value
        public UInt16 e_lfarlc;      // File address of relocation table
        public UInt16 e_ovno;    // Overlay number
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public UInt16[] e_res1;    // Reserved words
        public UInt16 e_oemid;       // OEM identifier (for e_oeminfo)
        public UInt16 e_oeminfo;     // OEM information; e_oemid specific
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 10)]
        public UInt16[] e_res2;    // Reserved words
        public Int32 e_lfanew;      // File address of new exe header

        private string _e_magic
        {
            get { return new string(e_magic); }
        }

        public bool isValid
        {
            get { return _e_magic == "MZ"; }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct IMAGE_FILE_HEADER
    {
        public UInt16 Machine;
        public UInt16 NumberOfSections;
        public UInt32 TimeDateStamp;
        public UInt32 PointerToSymbolTable;
        public UInt32 NumberOfSymbols;
        public UInt16 SizeOfOptionalHeader;
        public UInt16 Characteristics;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _CLIENT_ID
    {
        public IntPtr UniqueProcess;
        public IntPtr UniqueThread;

        public int Pid
        {
            get { return this.UniqueProcess.ToInt32(); }
        }

        public int Tid
        {
            get { return this.UniqueThread.ToInt32(); }
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _LIST_ENTRY
    {
        public IntPtr Flink;
        public IntPtr Blink;

        internal unsafe IntPtr Unlink()
        {
            _LIST_ENTRY* pNext = (_LIST_ENTRY*)Flink.ToPointer();
            _LIST_ENTRY* pPrev = (_LIST_ENTRY*)Blink.ToPointer();

            IntPtr thisLink = pNext->Blink;

            pNext->Blink = Blink;
            pPrev->Flink = Flink;

            _LIST_ENTRY* thisItem = (_LIST_ENTRY*)thisLink.ToPointer();
            
            thisItem->Blink = IntPtr.Zero;
            thisItem->Flink = IntPtr.Zero;

            return thisLink;
        }

        internal unsafe void LinkTo(IntPtr hiddenModuleLink)
        {
            if (hiddenModuleLink == IntPtr.Zero)
            {
                return;
            }

            _LIST_ENTRY* nextItem = (_LIST_ENTRY*)Flink.ToPointer();
            _LIST_ENTRY* thisItem = (_LIST_ENTRY*)nextItem->Blink.ToPointer();

            _LIST_ENTRY* newItem = (_LIST_ENTRY*)hiddenModuleLink.ToPointer();
            newItem->Flink = Flink;
            newItem->Blink = new IntPtr(thisItem);

            thisItem->Flink = hiddenModuleLink;
            nextItem->Blink = hiddenModuleLink;
        }
    }

#if _INCLUDE_MANAGED_STRUCTS
    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_INFORMATION
    {
        public int HandleCount;
        public _SYSTEM_HANDLE_TABLE_ENTRY_INFO Handles; /* Handles[0] */
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_INFORMATION_EX
    {
        public IntPtr HandleCount;
        public IntPtr Reserved;
        public _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX Handles; /* Handles[0] */

        public int NumberOfHandles
        {
            get { return HandleCount.ToInt32(); }
        }
    }
#endif

    [StructLayout(LayoutKind.Sequential)]
    public struct GENERIC_MAPPING
    {
        public uint GenericRead;
        public uint GenericWrite;
        public uint GenericExecute;
        public uint GenericAll;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct OBJECT_NAME_INFORMATION
    {
        public _UNICODE_STRING Name;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct OBJECT_TYPE_INFORMATION
    {
        public _UNICODE_STRING Name;
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

    [StructLayout(LayoutKind.Sequential)]
    public struct _UNICODE_STRING
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;

        public string GetText()
        {
            if (Buffer == IntPtr.Zero || MaximumLength == 0)
            {
                return "";
            }

            return Marshal.PtrToStringUni(Buffer, Length / 2);
        }
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-teb
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _TEB
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 12)]
        public IntPtr[] Reserved1;
        public IntPtr ProcessEnvironmentBlock;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 399)]
        public IntPtr[] Reserved2;
        public fixed byte Reserved3[1952];
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 64)]
        public IntPtr[] TlsSlots;
        public fixed byte Reserved4[8];
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 26)]
        public IntPtr[] Reserved5;
        public IntPtr ReservedForOle;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
        public IntPtr[] Reserved6;
        public IntPtr TlsExpansionSlots;

        public static _TEB Create(IntPtr tebAddress)
        {
            _TEB teb = (_TEB)Marshal.PtrToStructure(tebAddress, typeof(_TEB));
            return teb;
        }
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _PEB
    {
        public fixed byte Reserved1[2];
        public byte BeingDebugged;
        public fixed byte Reserved2[1];
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public IntPtr[] Reserved3;
        public IntPtr Ldr;
        public IntPtr ProcessParameters;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public IntPtr[] Reserved4;
        public IntPtr AtlThunkSListPtr;
        public IntPtr Reserved5;
        public uint Reserved6;
        public IntPtr Reserved7;
        public uint Reserved8;
        public uint AtlThunkSListPtr32;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 45)]
        public IntPtr[] Reserved9;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 96)]
        public byte[] Reserved10;
        public IntPtr PostProcessInitRoutine;
        public fixed byte Reserved11[128];
        public IntPtr Reserved12;
        public uint SessionId;

        public static _PEB Create(IntPtr pebAddress)
        {
            _PEB peb = (_PEB)Marshal.PtrToStructure(pebAddress, typeof(_PEB));
            return peb;
        }

#if _KSOBUILD
        public static IEnumerable<IntPtr> EnumerateHeaps(IntPtr pebAddress)
        {
            DbgOffset pebOffset = DbgOffset.Get("_PEB");

            IntPtr processHeapsPtr = pebOffset.GetPointer(pebAddress, "ProcessHeaps").ReadPtr();
            if (processHeapsPtr == IntPtr.Zero)
            {
                yield break;
            }

            if (pebOffset.TryRead<int>(pebAddress, "NumberOfHeaps", out int numberOfHeaps) == false)
            {
                yield break;
            }

            for (int i = 0; i < numberOfHeaps; i++)
            {
                IntPtr entryPtr = processHeapsPtr + (IntPtr.Size * i);
                yield return entryPtr.ReadPtr();
            }
        }
#endif
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb_ldr_data
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _PEB_LDR_DATA
    {
        public fixed byte Reserved1[8];
        public IntPtr Reserved2;
        public _LIST_ENTRY InLoadOrderModuleList;
        public _LIST_ENTRY InMemoryOrderModuleList;

        public static _PEB_LDR_DATA Create(IntPtr ldrAddress)
        {
            _PEB_LDR_DATA ldrData = (_PEB_LDR_DATA)Marshal.PtrToStructure(ldrAddress, typeof(_PEB_LDR_DATA));
            return ldrData;
        }

        public IEnumerable<_LDR_DATA_TABLE_ENTRY> EnumerateLoadOrderModules()
        {
            IntPtr startLink = InLoadOrderModuleList.Flink;
            _LDR_DATA_TABLE_ENTRY item = _LDR_DATA_TABLE_ENTRY.CreateFromLoadOrder(startLink);

            while (true)
            {
                if (item.DllBase != IntPtr.Zero)
                {
                    yield return item;
                }

                if (item.InLoadOrderLinks.Flink == startLink)
                {
                    break;
                }

                item = _LDR_DATA_TABLE_ENTRY.CreateFromLoadOrder(item.InLoadOrderLinks.Flink);
            }
        }

        public IEnumerable<_LDR_DATA_TABLE_ENTRY> EnumerateMemoryOrderModules()
        {
            IntPtr startLink = InMemoryOrderModuleList.Flink;
            _LDR_DATA_TABLE_ENTRY item = _LDR_DATA_TABLE_ENTRY.CreateFromMemoryOrder(startLink);

            while (true)
            {
                if (item.DllBase != IntPtr.Zero)
                {
                    yield return item;
                }

                if (item.InMemoryOrderLinks.Flink == startLink)
                {
                    break;
                }

                item = _LDR_DATA_TABLE_ENTRY.CreateFromMemoryOrder(item.InMemoryOrderLinks.Flink);
            }
        }

        public _LDR_DATA_TABLE_ENTRY Find(string dllFileName)
        {
            return Find(dllFileName, true);
        }

        public _LDR_DATA_TABLE_ENTRY Find(string dllFileName, bool memoryOrder)
        {
            foreach (_LDR_DATA_TABLE_ENTRY entry in 
                (memoryOrder == true) ? EnumerateMemoryOrderModules() : EnumerateLoadOrderModules())
            {
                if (entry.FullDllName.GetText().EndsWith(dllFileName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return entry;
                }
            }

            return default(_LDR_DATA_TABLE_ENTRY);
        }

#if _INCLUDE_MANAGED_STRUCTS
        public unsafe void UnhideDLL(DllOrderLink hiddenModuleLink)
        {
            _LDR_DATA_TABLE_ENTRY dllLink = EnumerateMemoryOrderModules().First();
            
            dllLink.InMemoryOrderLinks.LinkTo(hiddenModuleLink.MemoryOrderLink);
            dllLink.InLoadOrderLinks.LinkTo(hiddenModuleLink.LoadOrderLink);
        }

        public unsafe DllOrderLink HideDLL(string fileName)
        {
            _LDR_DATA_TABLE_ENTRY dllLink = Find(fileName);

            if (dllLink.DllBase == IntPtr.Zero)
            {
                return null;
            }

            DllOrderLink orderLink = new DllOrderLink();

            orderLink.MemoryOrderLink = dllLink.InMemoryOrderLinks.Unlink();
            orderLink.LoadOrderLink = dllLink.InLoadOrderLinks.Unlink();

            return orderLink;
        }
#endif
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb_ldr_data
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _LDR_DATA_TABLE_ENTRY
    {
        public _LIST_ENTRY InLoadOrderLinks;
        public _LIST_ENTRY InMemoryOrderLinks;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public IntPtr[] Reserved2;
        public IntPtr DllBase;
        public IntPtr EntryPoint;
        public IntPtr SizeOfImage;
        public _UNICODE_STRING FullDllName;
        public fixed byte Reserved4[8];
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public IntPtr[] Reserved5;

        /*
        union {
            ULONG CheckSum;
            IntPtr Reserved6;
        };
        */
        public IntPtr Reserved6;

        public uint TimeDateStamp;

        public static _LDR_DATA_TABLE_ENTRY CreateFromMemoryOrder(IntPtr memoryOrderLink)
        {
            IntPtr head = memoryOrderLink - Marshal.SizeOf(typeof(_LIST_ENTRY));

            _LDR_DATA_TABLE_ENTRY entry = (_LDR_DATA_TABLE_ENTRY)Marshal.PtrToStructure(
                head, typeof(_LDR_DATA_TABLE_ENTRY));

            return entry;
        }

        public static _LDR_DATA_TABLE_ENTRY CreateFromLoadOrder(IntPtr loadOrderLink)
        {
            IntPtr head = loadOrderLink;

            _LDR_DATA_TABLE_ENTRY entry = (_LDR_DATA_TABLE_ENTRY)Marshal.PtrToStructure(
                head, typeof(_LDR_DATA_TABLE_ENTRY));

            return entry;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct IMAGE_DATA_DIRECTORY
    {
        public UInt32 VirtualAddress;
        public UInt32 Size;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct IMAGE_OPTIONAL_HEADER32
    {
        [FieldOffset(0)]
        public MagicType Magic;

        [FieldOffset(2)]
        public byte MajorLinkerVersion;

        [FieldOffset(3)]
        public byte MinorLinkerVersion;

        [FieldOffset(4)]
        public uint SizeOfCode;

        [FieldOffset(8)]
        public uint SizeOfInitializedData;

        [FieldOffset(12)]
        public uint SizeOfUninitializedData;

        [FieldOffset(16)]
        public uint AddressOfEntryPoint;

        [FieldOffset(20)]
        public uint BaseOfCode;

        // PE32 contains this additional field
        [FieldOffset(24)]
        public uint BaseOfData;

        [FieldOffset(28)]
        public uint ImageBase;

        [FieldOffset(32)]
        public uint SectionAlignment;

        [FieldOffset(36)]
        public uint FileAlignment;

        [FieldOffset(40)]
        public ushort MajorOperatingSystemVersion;

        [FieldOffset(42)]
        public ushort MinorOperatingSystemVersion;

        [FieldOffset(44)]
        public ushort MajorImageVersion;

        [FieldOffset(46)]
        public ushort MinorImageVersion;

        [FieldOffset(48)]
        public ushort MajorSubsystemVersion;

        [FieldOffset(50)]
        public ushort MinorSubsystemVersion;

        [FieldOffset(52)]
        public uint Win32VersionValue;

        [FieldOffset(56)]
        public uint SizeOfImage;

        [FieldOffset(60)]
        public uint SizeOfHeaders;

        [FieldOffset(64)]
        public uint CheckSum;

        [FieldOffset(68)]
        public SubSystemType Subsystem;

        [FieldOffset(70)]
        public DllCharacteristicsType DllCharacteristics;

        [FieldOffset(72)]
        public uint SizeOfStackReserve;

        [FieldOffset(76)]
        public uint SizeOfStackCommit;

        [FieldOffset(80)]
        public uint SizeOfHeapReserve;

        [FieldOffset(84)]
        public uint SizeOfHeapCommit;

        [FieldOffset(88)]
        public uint LoaderFlags;

        [FieldOffset(92)]
        public uint NumberOfRvaAndSizes;

        [FieldOffset(96)]
        public IMAGE_DATA_DIRECTORY ExportTable;

        [FieldOffset(104)]
        public IMAGE_DATA_DIRECTORY ImportTable;

        [FieldOffset(112)]
        public IMAGE_DATA_DIRECTORY ResourceTable;

        [FieldOffset(120)]
        public IMAGE_DATA_DIRECTORY ExceptionTable;

        [FieldOffset(128)]
        public IMAGE_DATA_DIRECTORY CertificateTable;

        [FieldOffset(136)]
        public IMAGE_DATA_DIRECTORY BaseRelocationTable;

        [FieldOffset(144)]
        public IMAGE_DATA_DIRECTORY Debug;

        [FieldOffset(152)]
        public IMAGE_DATA_DIRECTORY Architecture;

        [FieldOffset(160)]
        public IMAGE_DATA_DIRECTORY GlobalPtr;

        [FieldOffset(168)]
        public IMAGE_DATA_DIRECTORY TLSTable;

        [FieldOffset(176)]
        public IMAGE_DATA_DIRECTORY LoadConfigTable;

        [FieldOffset(184)]
        public IMAGE_DATA_DIRECTORY BoundImport;

        [FieldOffset(192)]
        public IMAGE_DATA_DIRECTORY IAT;

        [FieldOffset(200)]
        public IMAGE_DATA_DIRECTORY DelayImportDescriptor;

        [FieldOffset(208)]
        public IMAGE_DATA_DIRECTORY CLRRuntimeHeader;

        [FieldOffset(216)]
        public IMAGE_DATA_DIRECTORY Reserved;
    }

    [StructLayout(LayoutKind.Explicit)]
    public struct IMAGE_OPTIONAL_HEADER64
    {
        [FieldOffset(0)]
        public MagicType Magic;

        [FieldOffset(2)]
        public byte MajorLinkerVersion;

        [FieldOffset(3)]
        public byte MinorLinkerVersion;

        [FieldOffset(4)]
        public uint SizeOfCode;

        [FieldOffset(8)]
        public uint SizeOfInitializedData;

        [FieldOffset(12)]
        public uint SizeOfUninitializedData;

        [FieldOffset(16)]
        public uint AddressOfEntryPoint;

        [FieldOffset(20)]
        public uint BaseOfCode;

        [FieldOffset(24)]
        public ulong ImageBase;

        [FieldOffset(32)]
        public uint SectionAlignment;

        [FieldOffset(36)]
        public uint FileAlignment;

        [FieldOffset(40)]
        public ushort MajorOperatingSystemVersion;

        [FieldOffset(42)]
        public ushort MinorOperatingSystemVersion;

        [FieldOffset(44)]
        public ushort MajorImageVersion;

        [FieldOffset(46)]
        public ushort MinorImageVersion;

        [FieldOffset(48)]
        public ushort MajorSubsystemVersion;

        [FieldOffset(50)]
        public ushort MinorSubsystemVersion;

        [FieldOffset(52)]
        public uint Win32VersionValue;

        [FieldOffset(56)]
        public uint SizeOfImage;

        [FieldOffset(60)]
        public uint SizeOfHeaders;

        [FieldOffset(64)]
        public uint CheckSum;

        [FieldOffset(68)]
        public SubSystemType Subsystem;

        [FieldOffset(70)]
        public DllCharacteristicsType DllCharacteristics;

        [FieldOffset(72)]
        public ulong SizeOfStackReserve;

        [FieldOffset(80)]
        public ulong SizeOfStackCommit;

        [FieldOffset(88)]
        public ulong SizeOfHeapReserve;

        [FieldOffset(96)]
        public ulong SizeOfHeapCommit;

        [FieldOffset(104)]
        public uint LoaderFlags;

        [FieldOffset(108)]
        public uint NumberOfRvaAndSizes;

        [FieldOffset(112)]
        public IMAGE_DATA_DIRECTORY ExportTable;

        [FieldOffset(120)]
        public IMAGE_DATA_DIRECTORY ImportTable;

        [FieldOffset(128)]
        public IMAGE_DATA_DIRECTORY ResourceTable;

        [FieldOffset(136)]
        public IMAGE_DATA_DIRECTORY ExceptionTable;

        [FieldOffset(144)]
        public IMAGE_DATA_DIRECTORY CertificateTable;

        [FieldOffset(152)]
        public IMAGE_DATA_DIRECTORY BaseRelocationTable;

        [FieldOffset(160)]
        public IMAGE_DATA_DIRECTORY Debug;

        [FieldOffset(168)]
        public IMAGE_DATA_DIRECTORY Architecture;

        [FieldOffset(176)]
        public IMAGE_DATA_DIRECTORY GlobalPtr;

        [FieldOffset(184)]
        public IMAGE_DATA_DIRECTORY TLSTable;

        [FieldOffset(192)]
        public IMAGE_DATA_DIRECTORY LoadConfigTable;

        [FieldOffset(200)]
        public IMAGE_DATA_DIRECTORY BoundImport;

        [FieldOffset(208)]
        public IMAGE_DATA_DIRECTORY IAT;

        [FieldOffset(216)]
        public IMAGE_DATA_DIRECTORY DelayImportDescriptor;

        [FieldOffset(224)]
        public IMAGE_DATA_DIRECTORY CLRRuntimeHeader;

        [FieldOffset(232)]
        public IMAGE_DATA_DIRECTORY Reserved;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct IMAGE_DEBUG_DIRECTORY
    {
        public uint Characteristics;
        public uint TimeDateStamp;
        public short MajorVersion;
        public short MinorVersion;
        public uint Type;
        public uint SizeOfData;
        public uint AddressOfRawData;
        public uint PointerToRawData;

#if _INCLUDE_MANAGED_STRUCTS
        internal unsafe CodeViewRSDS GetCodeViewHeader(IntPtr codeViewPtr)
        {
            CodeView_Header dir = (CodeView_Header)Marshal.PtrToStructure(codeViewPtr, typeof(CodeView_Header));

            switch (dir.Signature)
            {
                case (uint)CodeViewSignature.RSDS:
                    _CodeViewRSDS item = (_CodeViewRSDS)Marshal.PtrToStructure(codeViewPtr, typeof(_CodeViewRSDS));
                    CodeViewRSDS rsds = new CodeViewRSDS();
                    rsds.CvSignature = item.CvSignature;
                    rsds.Signature = item.Signature;
                    rsds.Age = item.Age;

                    int rsdsSize = Marshal.SizeOf(item);
                    IntPtr fileNamePtr = codeViewPtr + rsdsSize;
                    rsds.PdbFileName = Marshal.PtrToStringAnsi(fileNamePtr);
                    return rsds;

                default:
                    throw new ApplicationException("Push the author to impl this: " + dir.Signature);
            }
        }
#endif
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct CodeView_Header
    {
        public uint Signature;
        public int Offset;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct _CodeViewRSDS
    {
        public uint CvSignature;
        public Guid Signature;
        public uint Age;
        // fixed byte PdbFileName[1];
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct IMAGE_EXPORT_DIRECTORY
    {
        public uint Characteristics;
        public uint TimeDateStamp;
        public short MajorVersion;
        public short MinorVersion;
        public uint Name;
        public uint Base;
        public uint NumberOfFunctions;
        public uint NumberOfNames;
        public uint AddressOfFunctions;
        public uint AddressOfNames;
        public uint AddressOfNameOrdinals;
    }

    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct IMAGE_SECTION_HEADER
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 8)]
        public byte[] Name;
        public int PhysicalAddressOrVirtualSize;
        public int VirtualAddress;
        public int SizeOfRawData;
        public int PointerToRawData;
        public int PointerToRelocations;
        public int PointerToLineNumbers;
        public short NumberOfRelocations;
        public short NumberOfLineNumbers;
        public uint Characteristics;

        public string GetName()
        {
            string name = Encoding.ASCII.GetString(this.Name).Trim('\0');
            return name;
        }

        public override string ToString()
        {
            int startAddress = VirtualAddress;
            int endAddress = startAddress + PhysicalAddressOrVirtualSize;

            return $"{GetName()}: 0x{startAddress.ToString("x")} ~ 0x{endAddress.ToString("x")}";
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _IMAGEHLP_LINE64
    {
        public uint SizeOfStruct;           // set to sizeof(IMAGEHLP_LINE64)
        IntPtr Key;                    // internal
        public uint LineNumber;             // line number in file
        IntPtr FileName;               // full filename
        public ulong Address;                // first instruction of line
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SYSTEM_INFO
    {
        public ushort wProcessorArchitecture;
        public ushort wReserved;
        public uint dwPageSize;
        public IntPtr lpMinimumApplicationAddress;
        public IntPtr lpMaximumApplicationAddress;
        public IntPtr dwActiveProcessorMask;
        public uint dwNumberOfProcessors;
        public uint dwProcessorType;
        public uint dwAllocationGranularity;
        public ushort wProcessorLevel;
        public ushort wProcessorRevision;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _MODLOAD_DATA
    {
        public uint ssize;
        public uint ssig;
        public IntPtr data;
        public uint size;
        public uint flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _IMAGEHLP_MODULE64
    {
        public uint SizeOfStruct;           // set to sizeof(IMAGEHLP_MODULE64)
        public ulong BaseOfImage;            // base load address of module
        public uint ImageSize;              // virtual size of the loaded module
        public uint TimeDateStamp;          // date/time stamp from pe header
        public uint CheckSum;               // checksum from the pe header
        public uint NumSyms;                // number of symbols in the symbol table
        public SYM_TYPE SymType;                // type of symbols loaded

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
        public byte[] ModuleName;         // module name
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public byte[] ImageName;         // image name
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public byte[] LoadedImageName;   // symbol file name
                                         // new elements: 07-Jun-2002
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 256)]
        public byte[] LoadedPdbName;     // pdb file name

        public uint CVSig;                  // Signature of the CV record in the debug directories

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = NativeMethods.MAX_PATH * 3)]
        public byte[] CVData;   // Contents of the CV record
        public uint PdbSig;                 // Signature of PDB
        public Guid PdbSig70;               // Signature of PDB (VC 7 and up)
        public uint PdbAge;                 // DBI age of pdb
        public bool PdbUnmatched;           // loaded an unmatched pdb
        public bool DbgUnmatched;           // loaded an unmatched dbg
        public bool LineNumbers;            // we have line number information
        public bool GlobalSymbols;          // we have internal symbol information
        public bool TypeInfo;               // we have type information
                                            // new elements: 17-Dec-2003
        public bool SourceIndexed;          // pdb supports source server
        public bool Publics;                // contains public symbols
                                            // new element: 15-Jul-2009
        public uint MachineType;            // IMAGE_FILE_MACHINE_XXX from ntimage.h and winnt.h
        public uint Reserved;               // Padding - don't remove.

        public unsafe string GetLoadedPdbName()
        {
            fixed (byte* ptr = LoadedPdbName)
            {
                return Marshal.PtrToStringAnsi(new IntPtr(ptr));
            }
        }

        internal static _IMAGEHLP_MODULE64 Create()
        {
            _IMAGEHLP_MODULE64 mod = new _IMAGEHLP_MODULE64();
            mod.SizeOfStruct = (uint)Marshal.SizeOf(mod);
            return mod;
        }

        /*
            _IMAGEHLP_MODULE64 module = _IMAGEHLP_MODULE64.Create();
            if (NativeMethods.SymGetModuleInfo64(processHandle, (ulong)base_addr.ToInt64(), ref module) == false)
            {
                Console.WriteLine("Unexpected failure from SymGetModuleInfo64().");
                return;
            }

            Console.WriteLine("GetLoadedPdbName: " + module.GetLoadedPdbName());
        */
    }

    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    internal unsafe struct _SYMBOL_INFO
    {
        public uint SizeOfStruct;
        public uint TypeIndex;        // Type Index of symbol

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public ulong[] Reserved;
        public uint Index;
        public uint Size;
        public ulong ModBase;          // Base Address of module comtaining this symbol
        public uint Flags;
        public ulong Value;            // Value of symbol, ValuePresent should be 1
        public ulong Address;          // Address of symbol including base address of module
        public uint Register;         // register holding value or pointer to value
        public uint Scope;            // scope of the symbol
        public uint Tag;              // pdb classification
        public uint NameLen;          // Actual length of name
        public uint MaxNameLen;
        public byte Name;          // Name of symbol
    }

    [StructLayout(LayoutKind.Sequential, Pack = 8)]
    public unsafe struct SYMBOL_INFO
    {
        public uint SizeOfStruct;
        public uint TypeIndex;        // Type Index of symbol

        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public ulong Reserved0;
        public ulong Reserved1;
        public uint Index;
        public uint Size;
        public ulong ModBase;          // Base Address of module comtaining this symbol
        public uint Flags;
        public ulong Value;            // Value of symbol, ValuePresent should be 1
        public ulong Address;          // Address of symbol including base address of module
        public uint Register;         // register holding value or pointer to value
        public uint Scope;            // scope of the symbol
        public uint Tag;              // pdb classification
        public uint NameLen;          // Actual length of name
        public uint MaxNameLen;
        public string Name;          // Name of symbol

        public static SYMBOL_INFO Create(IntPtr baseAddress)
        {
            _SYMBOL_INFO info = (_SYMBOL_INFO)Marshal.PtrToStructure(baseAddress, typeof(_SYMBOL_INFO));

            SYMBOL_INFO si = new SYMBOL_INFO();
            si.SizeOfStruct = info.SizeOfStruct;
            si.TypeIndex = info.TypeIndex;
            si.Reserved0 = info.Reserved[0];
            si.Reserved1 = info.Reserved[1];
            si.Index = info.Index;
            si.Size = info.Size;
            si.ModBase = info.ModBase;
            si.Flags = info.Flags;
            si.Value = info.Value;
            si.Address = info.Address;
            si.Register = info.Register;
            si.Scope = info.Scope;
            si.Tag = info.Tag;
            si.NameLen = info.NameLen;
            si.MaxNameLen = info.MaxNameLen;

            int offset = Marshal.OffsetOf(typeof(_SYMBOL_INFO), nameof(info.Name)).ToInt32();
            si.Name = Marshal.PtrToStringAuto(baseAddress + offset, (int)info.NameLen).Trim('\0');

            return si;
        }
    }
}
