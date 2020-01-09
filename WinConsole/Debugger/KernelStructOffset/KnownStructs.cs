using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KernelStructOffset
{
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
            _LIST_ENTRY* thisItem = (_LIST_ENTRY*)thisLink.ToPointer();
            thisItem->Blink = IntPtr.Zero;
            thisItem->Flink = IntPtr.Zero;

            pNext->Blink = new IntPtr(pPrev);
            pPrev->Flink = new IntPtr(pNext);

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

            _LIST_ENTRY* targetItem = (_LIST_ENTRY*)hiddenModuleLink.ToPointer();

            targetItem->Flink = new IntPtr(nextItem);
            targetItem->Blink = new IntPtr(thisItem);

            thisItem->Flink = hiddenModuleLink;
            nextItem->Blink = hiddenModuleLink;
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_INFORMATION
    {
        public int HandleCount;
        public _SYSTEM_HANDLE_TABLE_ENTRY_INFO Handles; /* Handles[0] */
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct _SYSTEM_HANDLE_INFORMATION_EX
    {
        public int HandleCount;
        public _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX Handles; /* Handles[0] */
    }

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
}
