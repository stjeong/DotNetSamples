using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using WindowsPE;

namespace DetourFunc.ClrType
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct MethodDescChunkInternal
    {
        readonly IntPtr _methodTable;
        public IntPtr MethodTable => _methodTable;

        readonly IntPtr _next;
        public IntPtr Next => _next;

        readonly byte _size;
        public int Size => _size;

        readonly byte _count;
        public int Count => _count;

        readonly ushort _flagsAndTokenRange;
        public ushort FlagsAndTokenRange => _flagsAndTokenRange;

        public MethodDescChunkInternal(IntPtr methodTable, IntPtr next, byte size, byte count, ushort flagsAndTokenRange)
        {
            _methodTable = methodTable;
            _next = next;
            _size = size;
            _count = count;
            _flagsAndTokenRange = flagsAndTokenRange;
        }
    }

    public class MethodDescChunk
    {
        readonly MethodDescChunkInternal _internal;

        readonly IntPtr _address;
        public IntPtr Address => _address;

        public static uint SizeOf
        {
            get
            {
                return (uint)Marshal.SizeOf(typeof(MethodDescChunkInternal));
            }
        }

        public IntPtr GetMethodTablePtr()
        {
            return _internal.MethodTable;
        }

        private MethodDescChunk(IntPtr address)
        {
            int offset = 0;

            _internal = new MethodDescChunkInternal
            (
                methodTable: address.Add(address.ReadPtr(ref offset)),
                next: address.ReadPtr(ref offset),
                size: address.ReadByte(ref offset),
                count: address.ReadByte(ref offset),
                flagsAndTokenRange: address.ReadUInt16(ref offset)
            );

            _address = address;
        }

        public static MethodDescChunk ReadFromAddress(IntPtr ptr)
        {
            return new MethodDescChunk(ptr);
        }
    }

}
