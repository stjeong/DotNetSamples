using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc.Clr
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct MethodTableInternal
    {
        readonly uint _dwFlags;
        public uint Flags => _dwFlags;

        // Base size of instance of this class when allocated on the heap
        readonly uint _baseSize;
        public uint BaseSize => _baseSize;

        readonly ushort _wFlags2;
        public ushort Flags2 => _wFlags2;

        // Class token if it fits into 16-bits. If this is (WORD)-1, the class token is stored in the TokenOverflow optional member.
        readonly ushort _wToken;
        public ushort Token => _wToken;

        // <NICE> In the normal cases we shouldn't need a full word for each of these </NICE>
        readonly ushort _wNumVirtuals;
        public ushort NumVirtuals => _wNumVirtuals;

        readonly ushort _wNumInterfaces;
        public ushort NumInterfaces => _wNumInterfaces;

        /*
    #if defined(FEATURE_NGEN_RELOCS_OPTIMIZATIONS)
    RelativePointer<PTR_MethodTableWriteableData> m_pWriteableData;
    #else
    PlainPointer<PTR_MethodTableWriteableData> m_pWriteableData;
    #endif
        */
        readonly IntPtr _pWriteableData;

        /*
    union {
    #if defined(FEATURE_NGEN_RELOCS_OPTIMIZATIONS)
        RelativePointer<DPTR(EEClass)> m_pEEClass;
        RelativePointer<TADDR> m_pCanonMT;
    #else
        PlainPointer<DPTR(EEClass)> m_pEEClass;
        PlainPointer<TADDR> m_pCanonMT;
    #endif
    };
        */
        readonly IntPtr _pEEClass_Or_pCanonMT;

        /*
        #if defined(FEATURE_NGEN_RELOCS_OPTIMIZATIONS)
            typedef RelativePointer<PTR_Dictionary> PerInstInfoElem_t;
            typedef RelativePointer<DPTR(PerInstInfoElem_t)> PerInstInfo_t;
        #else
            typedef PlainPointer<PTR_Dictionary> PerInstInfoElem_t;
            typedef PlainPointer<DPTR(PerInstInfoElem_t)> PerInstInfo_t;
        #endif

        union
        {
            PerInstInfo_t m_pPerInstInfo;
            TADDR         m_ElementTypeHnd;
            TADDR         m_pMultipurposeSlot1;
        };
        */
        readonly IntPtr _pPerInstInfo_Or_ElementTypeHnd_Or_pMultipurposeSlot1;

        /*
            union
            {
        #if defined(FEATURE_NGEN_RELOCS_OPTIMIZATIONS)
                RelativePointer<PTR_InterfaceInfo>   m_pInterfaceMap;
        #else
                PlainPointer<PTR_InterfaceInfo>   m_pInterfaceMap;
        #endif
                TADDR               m_pMultipurposeSlot2;
            };
        */
        readonly IntPtr _pMultipurposeSlot2_Or_pInterfaceMap;

        // VTable and Non-Virtual slots go here

        // Overflow multipurpose slots go here

        // Optional Members go here
        //    See above for the list of optional members

        // Generic dictionary pointers go here

        // Interface map goes here

        // Generic instantiation+dictionary goes here

        public MethodTableInternal(uint dwFlags, uint baseSize, ushort wFlags2, ushort wToken, ushort wNumVirtuals,
            ushort wNumInterfaces, IntPtr pWriteableData, IntPtr pEEClass_Or_pCanonMT,
            IntPtr pPerInstInfo_Or_ElementTypeHnd_Or_pMultipurposeSlot1, IntPtr pMultipurposeSlot2_Or_pInterfaceMap)
        {
            _dwFlags = dwFlags;
            _baseSize = baseSize;
            _wFlags2 = wFlags2;
            _wToken = wToken;
            _wNumVirtuals = wNumVirtuals;
            _wNumInterfaces = wNumInterfaces;
            _pWriteableData = pWriteableData;
            _pEEClass_Or_pCanonMT = pEEClass_Or_pCanonMT;
            _pPerInstInfo_Or_ElementTypeHnd_Or_pMultipurposeSlot1 = pPerInstInfo_Or_ElementTypeHnd_Or_pMultipurposeSlot1;
            _pMultipurposeSlot2_Or_pInterfaceMap = pMultipurposeSlot2_Or_pInterfaceMap;
        }
    }

    public class MethodTable
    {
        readonly MethodTableInternal _internal;

        readonly IntPtr _address;
        public IntPtr Address => _address;

        public static uint SizeOf
        {
            get
            {
                return (uint)Marshal.SizeOf(typeof(MethodTableInternal));
            }
        }

        public bool HasNonVirtualSlots()
        {
            return GetWEnumFlag(WFLAGS2_ENUM.enum_flag_HasNonVirtualSlots) != 0;
        }

        public uint GetHighEnumFlag(WFLAGS_HIGH_ENUM flag)
        {
            return _internal.Flags & (uint)flag;
        }

        public uint GetLowEnumFlag(WFLAGS_LOW_ENUM flag)
        {
            return (IsStringOrArray() ? ((uint)WFLAGS_LOW_ENUM.enum_flag_StringArrayValues & (uint)flag) : (_internal.Flags & (uint)flag));
        }

        public uint GetWEnumFlag(WFLAGS2_ENUM flag)
        {
            return _internal.Flags2 & (uint)flag;
        }

        public bool IsStringOrArray()
        {
            return HasComponentSize();
        }

        public bool HasComponentSize()
        {
            return GetHighEnumFlag(WFLAGS_HIGH_ENUM.enum_flag_HasComponentSize) != 0;
        }

        public bool HasSingleNonVirtualSlot()
        {
            return GetWEnumFlag(WFLAGS2_ENUM.enum_flag_HasSingleNonVirtualSlot) != 0;
        }

        public uint GetNumVirtuals()
        {
            return _internal.NumVirtuals;
        }

        private MethodTable(IntPtr address)
        {
            int offset = 0;

            _internal = new MethodTableInternal(
                dwFlags: address.ReadUInt32(ref offset),
                baseSize: address.ReadUInt32(ref offset),
                wFlags2: address.ReadUInt16(ref offset),
                wToken: address.ReadUInt16(ref offset),
                wNumVirtuals: address.ReadUInt16(ref offset),
                wNumInterfaces: address.ReadUInt16(ref offset),

                pWriteableData: address.ReadPtr(ref offset),
                pEEClass_Or_pCanonMT: address.ReadPtr(ref offset),
                pPerInstInfo_Or_ElementTypeHnd_Or_pMultipurposeSlot1: address.ReadPtr(ref offset),
                pMultipurposeSlot2_Or_pInterfaceMap: address.ReadPtr(ref offset)
            );

            _address = address;
        }

        public static MethodTable ReadFromAddress(IntPtr ptr)
        {
            return new MethodTable(ptr);
        }
    }

}
