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

        public static unsafe ushort ReadUInt16ByIndex(this IntPtr ptr, int index)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), (index + 1) * sizeof(ushort));
            ums.Position = index * sizeof(ushort);
            return ums.ReadUInt16();
        }

        public static unsafe uint ReadUInt32ByIndex(this IntPtr ptr, int index)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)ptr.ToPointer(), (index + 1) * sizeof(uint));
            ums.Position = index * sizeof(uint);
            return ums.ReadUInt32();
        }

        public static IntPtr ReadPtr(this IntPtr ptr, int offset)
        {
            return Marshal.ReadIntPtr(ptr, offset);
        }

        public static uint ReadUInt32(this IntPtr ptr, int offset)
        {
            return (uint)Marshal.ReadInt32(ptr, offset);
        }
    }
}
