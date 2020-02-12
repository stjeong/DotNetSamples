using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc.ClrType
{
    // https://github.com/dotnet/runtime/blob/master/src/coreclr/src/vm/method.hpp

    [StructLayout(LayoutKind.Sequential)]
    internal struct MethodDescInternal
    {
        readonly ushort _wTokenRemainder;
        public ushort TokenRemainder => _wTokenRemainder;
        public uint Token
        {
            get { return ((uint)_wTokenRemainder & (uint)MethodDescFlags3.TokenRemainderMask) | (uint)CorTokenType.mdtMethodDef; }
        }

        readonly byte _chunkIndex;
        public byte ChunkIndex => _chunkIndex;

        readonly byte _bFlags2;
        public MethodDescFlags2 Flags2
        {
            get
            {
                MethodDescFlags2 flag = (MethodDescFlags2)_bFlags2;
                return flag;
            }
        }

        readonly ushort _wSlotNumber;
        public ushort SlotNumber => _wSlotNumber;

        readonly ushort _wFlags;
        public ushort Flags => _wFlags;

        public MethodDescInternal(ushort tokenRemainder, byte chunkIndex, byte flags2, ushort slotNumber, ushort flags)
        {
            _wTokenRemainder = tokenRemainder;
            _chunkIndex = chunkIndex;
            _bFlags2 = flags2;
            _wSlotNumber = slotNumber;
            _wFlags = flags;
        }
    }

    public class MethodDesc
    {
        static readonly int ALIGNMENT_SHIFT = (IntPtr.Size == 8) ? 3 : 2;
        public static readonly int ALIGNMENT = (1 << ALIGNMENT_SHIFT);
        // static readonly int ALIGNMENT_MASK = (ALIGNMENT - 1);

        readonly MethodDescInternal _internal;
        readonly IntPtr _address;

        public bool HasStableEntryPoint()
        {
            return (_internal.Flags2 & MethodDescFlags2.HasStableEntryPoint) != 0;
        }

        public uint GetBaseSize()
        {
            return MethodDesc.GetBaseSize(GetClassification());
        }

        public static uint GetBaseSize(uint classification)
        {
            switch ((MethodClassification)classification)
            {
                case MethodClassification.mcIL:
                    return MethodDesc.SizeOf;
            }

            throw new NotSupportedException();
        }

        public int GetSlot()
        {
            if (RequiresFullSlotNumber() == false)
            {
                return (_internal.SlotNumber & (ushort)PackedSlotLayout.SlotMask);
            }

            return _internal.SlotNumber;
        }

        int GetMethodDescIndex()
        {
            return _internal.ChunkIndex;
        }

        bool RequiresFullSlotNumber()
        {
            return (_internal.Flags & (ushort)MethodDescClassification.mdcRequiresFullSlotNumber) != 0;
        }

        public MethodTable GetMethodTable()
        {
            MethodDescChunk chunk = GetMethodDescChunk();
            return MethodTable.ReadFromAddress(chunk.GetMethodTablePtr());
        }

        public MethodDescChunk GetMethodDescChunk()
        {
            int offset = (int)MethodDescChunk.SizeOf + (GetMethodDescIndex() * ALIGNMENT);
            IntPtr chunkPtr = _address - offset;

            return MethodDescChunk.ReadFromAddress(chunkPtr);
        }

        uint GetClassification()
        {
            return (_internal.Flags & (uint)MethodDescClassification.mdcClassification);
        }

        string GetMethodDescType()
        {
            switch ((MethodClassification)GetClassification())
            {
                case MethodClassification.mcIL:
                    return "IL";

                case MethodClassification.mcFCall:
                    return "FCall";

                case MethodClassification.mcNDirect:
                    return "N/Dircet";

                case MethodClassification.mcEEImpl:
                    return "EEImpl";

                case MethodClassification.mcArray:
                    return "ArrayECall";

                case MethodClassification.mcInstantiated:
                    return "InstantiatedGenericMethod";

                case MethodClassification.mcComInterop:
                    return "ComInterop";

                case MethodClassification.mcDynamic:
                    return "Dynamic";
            }

            return null;
        }

        public bool HasNonVtableSlot()
        {
            return (_internal.Flags & (ushort)MethodDescClassification.mdcHasNonVtableSlot) != 0;
        }

        public static uint SizeOf
        {
            get
            {
                return (uint)Marshal.SizeOf(typeof(MethodDescInternal));
            }
        }

        protected MethodDesc(IntPtr address)
        {
            int offset = 0;

            _internal = new MethodDescInternal
            (
                tokenRemainder: address.ReadUInt16(ref offset),
                chunkIndex: address.ReadByte(ref offset),
                flags2: address.ReadByte(ref offset),
                slotNumber: address.ReadUInt16(ref offset),
                flags: address.ReadUInt16(ref offset)
            );

            _address = address;
        }

        public static MethodDesc ReadFromAddress(IntPtr address)
        {
            return new MethodDesc(address);
        }

        public virtual void Dump(TextWriter writer)
        {
            StringBuilder sb = new StringBuilder();

            sb.AppendLine($"[MethodDesc 0x{_address.ToInt64():x} - {GetMethodDescType()}]");
            sb.AppendLine($"\twTokenRemainder = {_internal.TokenRemainder:x} (Token = {_internal.Token:x})");
            sb.AppendLine($"\tchunkIndex = {_internal.ChunkIndex:x}");
            sb.AppendLine($"\tbFlags2 = {_internal.Flags2:x} (Flags2 == {_internal.Flags2})");
            sb.AppendLine($"\twSlotNumber = {GetSlot():x}");
            sb.AppendLine($"\twFlags = {_internal.Flags:x} (IsFullSlotNumber == {RequiresFullSlotNumber()})");
            sb.AppendLine($"\tMethodTablePtr = {GetMethodDescChunk().GetMethodTablePtr().ToInt64():x}");

            writer.WriteLine(sb.ToString());
        }
    }

}
