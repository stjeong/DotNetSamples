using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsPE
{
    public static class ExtensionHelper
    {
        public static ushort PeekUInt16(this BinaryReader reader)
        {
            long oldPosition = reader.BaseStream.Position;
            ushort result = reader.ReadUInt16();
            reader.BaseStream.Position = oldPosition;

            return result;
        }

        public unsafe static T Read<T>(this BinaryReader reader) where T: new()
        {
            T obj = new T();
            int typeSize = Marshal.SizeOf(obj);

            byte[] buffer = new byte[typeSize];
            reader.Read(buffer, 0, typeSize);

            fixed (byte *p = buffer)
            {
                IntPtr ptr = new IntPtr(p);
                T objSectionHeader = (T)Marshal.PtrToStructure(ptr, typeof(T));
                return objSectionHeader;
            }
        }

        public static ushort ReadUInt16(this UnmanagedMemoryStream reader)
        {
            byte[] buf = new byte[4];
            reader.Read(buf, 0, 4);

            return BitConverter.ToUInt16(buf, 0);
        }

        public static int ReadInt32(this UnmanagedMemoryStream reader)
        {
            byte[] buf = new byte[4];
            reader.Read(buf, 0, 4);

            return BitConverter.ToInt32(buf, 0);
        }

        public static uint ReadUInt32(this UnmanagedMemoryStream reader)
        {
            byte[] buf = new byte[4];
            reader.Read(buf, 0, 4);

            return BitConverter.ToUInt32(buf, 0);
        }

        public static byte ReadByte(this IntPtr addresss, ref int offset)
        {
            byte result = Marshal.ReadByte(addresss, offset);
            offset += 1;
            return result;
        }

        public static unsafe byte ReadByte(this IntPtr ptr, int position)
        {
            return Marshal.ReadByte(ptr, position);
        }

        public static unsafe byte [] ReadBytes(this IntPtr ptr, int length)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), length);

            byte[] buf = new byte[length];
            ums.Read(buf, 0, length);

            return buf;
        }

        public static unsafe ushort ReadUInt16ByIndex(this IntPtr ptr, int index)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), (index + 1) * sizeof(ushort))
            {
                Position = index * sizeof(ushort),
            };

            return ums.ReadUInt16();
        }

        public static unsafe uint ReadUInt32ByIndex(this IntPtr ptr, int index)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), (index + 1) * sizeof(uint))
            {
                Position = index * sizeof(uint),
            };

            return ums.ReadUInt32();
        }

        public static unsafe IntPtr ReadPtr(this IntPtr addresss)
        {
            return ReadPtr(addresss, 0);
        }

        public static IntPtr ReadPtr(this IntPtr ptr, int offset)
        {
            return Marshal.ReadIntPtr(ptr, offset);
        }

        public static unsafe IntPtr ReadPtr(this IntPtr addresss, ref int offset)
        {
            IntPtr target = addresss + offset;
            offset += IntPtr.Size;

            return Marshal.ReadIntPtr(target, 0);
        }

        public static uint ReadUInt32(this IntPtr ptr, int offset)
        {
            return (uint)Marshal.ReadInt32(ptr, offset);
        }

        public static ulong ReadUInt64(this IntPtr ptr)
        {
            return (ulong)Marshal.ReadInt64(ptr, 0);
        }

        public static ulong ReadUInt64(this IntPtr ptr, int offset)
        {
            return (ulong)Marshal.ReadInt64(ptr, offset);
        }

        public static long ReadInt64(this IntPtr ptr)
        {
            return Marshal.ReadInt64(ptr, 0);
        }

        public static int ReadInt32(this IntPtr ptr)
        {
            return Marshal.ReadInt32(ptr, 0);
        }

        public static uint ReadUInt32(this IntPtr addresss, ref int offset)
        {
            uint result = (uint)Marshal.ReadInt32(addresss, offset);
            offset += 4;
            return result;
        }

        public static void WriteInt64(this IntPtr ptr, long value)
        {
            Marshal.WriteInt64(ptr, value);
        }

        public static void WriteInt32(this IntPtr ptr, int value)
        {
            Marshal.WriteInt32(ptr, value);
        }

        public static void WriteByte(this IntPtr ptr, int offset, byte value)
        {
            Marshal.WriteByte(ptr, offset, value);
        }

        public static unsafe void WriteBytes(this IntPtr ptr, byte [] buf)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), buf.Length, buf.Length, FileAccess.Write);
            ums.Write(buf, 0, buf.Length);
        }

        public static ulong ToUInt64(this IntPtr ptr)
        {
            return (ulong)ptr.ToInt64();
        }

        public static uint ToUInt32(this IntPtr ptr)
        {
            return (uint)ptr.ToInt32();
        }

        public static short ReadInt16(this IntPtr addresss, ref int offset)
        {
            short result = Marshal.ReadInt16(addresss, offset);
            offset += 2;
            return result;
        }

        public static ushort ReadUInt16(this IntPtr addresss, ref int offset)
        {
            ushort result = (ushort)Marshal.ReadInt16(addresss, offset);
            offset += 2;
            return result;
        }

        public static IntPtr Add(this IntPtr address, IntPtr offset)
        {
            long newPtr = address.ToInt64() + offset.ToInt64();
            return new IntPtr(newPtr);
        }

        public static IntPtr Add(this IntPtr address, uint offset)
        {
            long newPtr = address.ToInt64() + offset;
            return new IntPtr(newPtr);
        }

    }
}
