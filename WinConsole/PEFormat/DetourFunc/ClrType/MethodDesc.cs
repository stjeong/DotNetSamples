using SharpDisasm;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc.Clr
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
        // readonly MethodInfo _methodInfo;

        public bool HasStableEntryPoint()
        {
            return (_internal.Flags2 & MethodDescFlags2.HasStableEntryPoint) != 0;
        }

        public MethodDescFlags2 Flags => _internal.Flags2;

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

        public IntPtr GetNativeFunctionPointer()
        {
            if (HasStableEntryPoint() == false)
            {
                return IntPtr.Zero;
            }

            IntPtr ptrEntry = GetFunctionPointer();
            if (ptrEntry == IntPtr.Zero)
            {
                return IntPtr.Zero;
            }

            SharpDisasm.ArchitectureMode mode = (IntPtr.Size == 8) ? SharpDisasm.ArchitectureMode.x86_64 : SharpDisasm.ArchitectureMode.x86_32;
            SharpDisasm.Disassembler.Translator.IncludeAddress = false;
            SharpDisasm.Disassembler.Translator.IncludeBinary = false;

            {
                byte[] buf = ptrEntry.ReadBytes(NativeMethods.MaxLengthOpCode);
                var disasm = new SharpDisasm.Disassembler(buf, mode, (ulong)ptrEntry.ToInt64());

                Instruction inst = disasm.Disassemble().First();
                if (inst.Mnemonic == SharpDisasm.Udis86.ud_mnemonic_code.UD_Ijmp)
                {
                    // Visual Studio + F5 Debug = Always point to "Fixup Precode"
                    long address = (long)inst.PC + inst.Operands[0].Value;
                    return new IntPtr(address);
                }
                else
                {
                    return ptrEntry;
                }
            }
        }

        public IntPtr GetFunctionPointer()
        {
            switch ((MethodClassification)GetClassification())
            {
                case MethodClassification.mcIL:
                    IntPtr funcPtr = _address + (int)MethodDesc.SizeOf;
                    return funcPtr.ReadPtr();

                case MethodClassification.mcFCall:
                case MethodClassification.mcNDirect:
                case MethodClassification.mcEEImpl:
                case MethodClassification.mcArray:
                case MethodClassification.mcInstantiated:
                case MethodClassification.mcComInterop:
                case MethodClassification.mcDynamic:
                    throw new NotSupportedException("IL method only");
            }

            return IntPtr.Zero;
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

        protected MethodDesc(MethodInfo mi) : this(mi.MethodHandle.Value)
        {
            // this._methodInfo = mi;
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

        /*
        public static void SetFlags(IntPtr methodDescAddress, MethodDescFlags2 flags)
        {
            methodDescAddress.WriteByte(sizeof(UInt16) + sizeof(byte), (byte)flags);
        }
        */

        public static MethodDesc ReadFromAddress(IntPtr address)
        {
            return new MethodDesc(address);
        }

        public static MethodDesc ReadFromMethodInfo(MethodInfo mi, bool jit = true)
        {
            if (jit == true)
            {
                RuntimeHelpers.PrepareMethod(mi.MethodHandle);
            }

            return new MethodDesc(mi);
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

            if (GetClassification() == (uint)MethodClassification.mcIL)
            {
                sb.AppendLine($"\tFunctionPtr = {GetFunctionPointer().ToInt64():x}");
            }

            writer.WriteLine(sb.ToString());
        }
    }
}
