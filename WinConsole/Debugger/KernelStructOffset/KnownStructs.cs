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
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
        public IntPtr[] Reserved2;
        public _LIST_ENTRY InMemoryOrderModuleList;

        public static _PEB_LDR_DATA Create(IntPtr ldrAddress)
        {
            _PEB_LDR_DATA ldrData = (_PEB_LDR_DATA)Marshal.PtrToStructure(ldrAddress, typeof(_PEB_LDR_DATA));
            return ldrData;
        }

        public IEnumerable<_LDR_DATA_TABLE_ENTRY> EnumerateModules()
        {
            IntPtr startLink = InMemoryOrderModuleList.Flink;
            _LDR_DATA_TABLE_ENTRY item = _LDR_DATA_TABLE_ENTRY.Create(startLink);

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

                item = _LDR_DATA_TABLE_ENTRY.Create(item.InMemoryOrderLinks.Flink);
            }
        }

        public _LDR_DATA_TABLE_ENTRY Find(string dllName)
        {
            foreach (_LDR_DATA_TABLE_ENTRY entry in EnumerateModules())
            {
                if (entry.FullDllName.GetText().EndsWith(dllName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    return entry;
                }
            }

            return default(_LDR_DATA_TABLE_ENTRY);
        }
    }

    // https://docs.microsoft.com/en-us/windows/win32/api/winternl/ns-winternl-peb_ldr_data
    [StructLayout(LayoutKind.Sequential)]
    public unsafe struct _LDR_DATA_TABLE_ENTRY
    {
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
        public IntPtr[] Reserved1;
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

        public static _LDR_DATA_TABLE_ENTRY Create(IntPtr memoryOrderLink)
        {
            IntPtr head = memoryOrderLink - Marshal.SizeOf(typeof(_LIST_ENTRY));

            _LDR_DATA_TABLE_ENTRY entry = (_LDR_DATA_TABLE_ENTRY)Marshal.PtrToStructure(
                head, typeof(_LDR_DATA_TABLE_ENTRY));

            return entry;
        }
    }
}
