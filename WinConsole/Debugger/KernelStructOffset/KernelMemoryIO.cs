using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KernelStructOffset
{
    // Prerequisite:
    //  Register and start "KernelMemoryIO" kernel driver
    //  https://github.com/stjeong/KernelMemoryIO/tree/master/KernelMemoryIO
    //
    // sc create "KernelMemoryIO" binPath= "D:\Debug\KernelMemoryIO.sys" type= kernel start= demand
    // sc delete "KernelMemoryIO"
    // net start KernelMemoryIO
    // net stop KernelMemoryIO
    public class KernelMemoryIO : IDisposable
    {
        const uint IOCTL_READ_MEMORY = ((((uint)40000) << 16) | ((0) << 14) | ((0x904) << 2) | (0)); // 0x9c402410;
        const uint IOCTL_READ_PHYSICAL_MEMORY = ((((uint)40000) << 16) | ((0) << 14) | ((0x914) << 2) | (0)); // 0x9c402450;

        const uint IOCTL_WRITE_MEMORY = 0x9c402414;
        const uint IOCTL_GETPOS_MEMORY = 0x9c402418;
        const uint IOCTL_SETPOS_MEMORY = 0x9c40241c;

        const uint IOCTL_READ_PORT_UCHAR = ((((uint)40000) << 16) | ((0) << 14) | ((0x908) << 2) | (0));  // 9c402420
        const uint IOCTL_READ_PORT_USHORT = ((((uint)40000) << 16) | ((0) << 14) | ((0x918) << 2) | (0)); // 9c402460
        const uint IOCTL_READ_PORT_ULONG = ((((uint)40000) << 16) | ((0) << 14) | ((0x928) << 2) | (0));  // 9c4024a0

        const uint IOCTL_WRITE_PORT_UCHAR = ((((uint)40000) << 16) | ((0) << 14) | ((0x909) << 2) | (0));  // 0x9c402424
        const uint IOCTL_WRITE_PORT_USHORT = ((((uint)40000) << 16) | ((0) << 14) | ((0x919) << 2) | (0)); // 0x9c402464
        const uint IOCTL_WRITE_PORT_ULONG = ((((uint)40000) << 16) | ((0) << 14) | ((0x929) << 2) | (0));  // 0x9c4024a4

        const uint IOCTL_KMIO_TEST = ((((uint)40000) << 16) | ((0) << 14) | ((0x100) << 2) | (0));

        SafeFileHandle fileHandle;

        public KernelMemoryIO()
        {
            InitializeDevice();
        }

        public bool InitializeDevice()
        {
            Dispose();

            fileHandle = NativeMethods.CreateFile(@"\\.\KernelMemoryIO", NativeFileAccess.FILE_GENERIC_READ,
                NativeFileShare.NONE, IntPtr.Zero, NativeFileMode.OPEN_EXISTING, NativeFileFlag.FILE_ATTRIBUTE_NORMAL, IntPtr.Zero);

            if (fileHandle.IsInvalid == true)
            {
                return false;
            }

            return true;
        }

        public void Dispose()
        {
            if (fileHandle != null)
            {
                fileHandle.Close();
                fileHandle = null;
            }
        }

        public bool IsInitialized
        {
            get
            {
                return fileHandle != null && fileHandle.IsInvalid == false;
            }
        }

        public IntPtr Position
        {
            get
            {
                if (this.IsInitialized == false)
                {
                    return IntPtr.Zero;
                }

                byte[] addressBytes = new byte[IntPtr.Size];

                if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_GETPOS_MEMORY,
                    null, 0, addressBytes, addressBytes.Length, out int _ /* pBytesReturned */, IntPtr.Zero) == true)
                {
                    if (IntPtr.Size == 8)
                    {
                        return new IntPtr(BitConverter.ToInt64(addressBytes, 0));
                    }

                    return new IntPtr(BitConverter.ToInt32(addressBytes, 0));
                }

                return IntPtr.Zero;
            }
            set
            {
                if (this.IsInitialized == false)
                {
                    return;
                }

                byte[] addressBytes;

                if (IntPtr.Size == 8)
                {
                    addressBytes = BitConverter.GetBytes(value.ToInt64());
                }
                else
                {
                    addressBytes = BitConverter.GetBytes(value.ToInt32());
                }

                NativeMethods.DeviceIoControl(fileHandle, IOCTL_SETPOS_MEMORY, addressBytes, addressBytes.Length, null, 9,
                    out int _ /* pBytesReturned */, IntPtr.Zero);
            }
        }

        public int WriteMemory(IntPtr ptr, byte[] buffer)
        {
            if (this.IsInitialized == false)
            {
                return 0;
            }

            this.Position = ptr;

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_WRITE_MEMORY, buffer, buffer.Length,
                null, 0, out int pBytesReturned, IntPtr.Zero) == true)
            {
                return pBytesReturned;
            }

            return 0;
        }

        public unsafe int WriteMemory<T>(void* ptr, T instance) where T : struct
        {
            return WriteMemory<T>(new IntPtr(ptr), instance);
        }

        public unsafe int WriteMemory<T>(IntPtr ptr, T instance) where T : struct
        {
            if (this.IsInitialized == false)
            {
                return 0;
            }

            int size = Marshal.SizeOf(instance);
            byte[] buffer = new byte[size];

            fixed (byte* ptrBuffer = buffer)
            {
                IntPtr targetBuffer = new IntPtr(ptrBuffer);
                Marshal.StructureToPtr(instance, targetBuffer, true);

                return WriteMemory(ptr, buffer);
            }
        }

        public int ReadPhysicalMemory(IntPtr position, byte[] buffer)
        {
            if (this.IsInitialized == false)
            {
                return 0;
            }

            byte[] addressBytes;

            if (IntPtr.Size == 8)
            {
                addressBytes = BitConverter.GetBytes(position.ToInt64());
            }
            else
            {
                addressBytes = BitConverter.GetBytes(position.ToInt32());
            }

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_READ_PHYSICAL_MEMORY, addressBytes, addressBytes.Length,
                buffer, buffer.Length,
                out int pBytesReturned, IntPtr.Zero) == true)
            {
                return pBytesReturned;
            }

            return 0;
        }

        public int ReadMemory(IntPtr position, byte[] buffer)
        {
            if (this.IsInitialized == false)
            {
                return 0;
            }

            byte[] addressBytes;

            if (IntPtr.Size == 8)
            {
                addressBytes = BitConverter.GetBytes(position.ToInt64());
            }
            else
            {
                addressBytes = BitConverter.GetBytes(position.ToInt32());
            }

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_READ_MEMORY, addressBytes, addressBytes.Length,
                buffer, buffer.Length,
                out int pBytesReturned, IntPtr.Zero) == true)
            {
                return pBytesReturned;
            }

            return 0;
        }

        public unsafe T ReadMemory<T>(void* ptr) where T : struct
        {
            return ReadMemory<T>(new IntPtr(ptr));
        }

        public unsafe T ReadMemory<T>(IntPtr ptr) where T : struct
        {
            T dummy = new T();
            int size = Marshal.SizeOf(dummy);

            byte[] buffer = new byte[size];

            if (ReadMemory(ptr, buffer) != buffer.Length)
            {
                return default;
            }

            fixed (byte* ptrBuffer = buffer)
            {
                IntPtr ptrTarget = new IntPtr(ptrBuffer);
                T instance = (T)Marshal.PtrToStructure(ptrTarget, typeof(T));

                return instance;
            }
        }

        public byte Inportb(ushort portNumber, out int errorNo)
        {
            errorNo = 0;
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                errorNo = -1;
                return 0;
            }

            byte[] portBytes = BitConverter.GetBytes(portNumber);

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_READ_PORT_UCHAR, portBytes, portBytes.Length, portBytes, portBytes.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero) == true)
            {
                return portBytes[0];
            }

            errorNo = Marshal.GetLastWin32Error();
            return 0;
        }

        public ushort Inports(ushort portNumber, out int errorNo)
        {
            errorNo = 0;
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                errorNo = -1;
                return 0;
            }

            byte[] portBytes = BitConverter.GetBytes(portNumber);

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_READ_PORT_USHORT, portBytes, portBytes.Length, portBytes, portBytes.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero) == true)
            {
                return BitConverter.ToUInt16(portBytes, 0);
            }

            errorNo = Marshal.GetLastWin32Error();
            return 0;
        }

        public uint Inportl(ushort portNumber, out int errorNo)
        {
            errorNo = 0;
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                errorNo = -1;
                return 0;
            }

            byte[] portBytes = BitConverter.GetBytes((uint)portNumber);

            if (NativeMethods.DeviceIoControl(fileHandle, IOCTL_READ_PORT_ULONG, portBytes, portBytes.Length, portBytes, portBytes.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero) == true)
            {
                return BitConverter.ToUInt32(portBytes, 0);
            }

            errorNo = Marshal.GetLastWin32Error();
            return 0;
        }

        public bool Outportb(ushort portNumber, byte data)
        {
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                return false;
            }

            byte[] portBytes = BitConverter.GetBytes(portNumber);
            byte[] inBuffer = new byte[3] { portBytes[0], portBytes[1], data };

            return NativeMethods.DeviceIoControl(fileHandle, IOCTL_WRITE_PORT_UCHAR, inBuffer, inBuffer.Length, inBuffer, inBuffer.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero);
        }

        public bool Outports(ushort portNumber, ushort data)
        {
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                return false;
            }

            byte[] portBytes = BitConverter.GetBytes(portNumber);
            byte[] dataBytes = BitConverter.GetBytes(data);

            byte[] inBuffer = new byte[4];
            Array.Copy(portBytes, inBuffer, portBytes.Length);
            Array.Copy(dataBytes, 0, inBuffer, 2, dataBytes.Length);

            return NativeMethods.DeviceIoControl(fileHandle, IOCTL_WRITE_PORT_USHORT, inBuffer, inBuffer.Length, inBuffer, inBuffer.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero);
        }

        public bool Outportl(uint portNumber, uint data)
        {
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                return false;
            }

            byte[] portBytes = BitConverter.GetBytes(portNumber);
            byte[] dataBytes = BitConverter.GetBytes(data);

            byte[] inBuffer = new byte[8];
            Array.Copy(portBytes, inBuffer, portBytes.Length);
            Array.Copy(dataBytes, 0, inBuffer, 4, dataBytes.Length);

            return NativeMethods.DeviceIoControl(fileHandle, IOCTL_WRITE_PORT_ULONG, inBuffer, inBuffer.Length, inBuffer, inBuffer.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero);
        }

        public void Test()
        {
            if (fileHandle == null || fileHandle.IsInvalid == true)
            {
                return;
            }

            byte[] inBuffer = new byte[0];

            NativeMethods.DeviceIoControl(fileHandle, IOCTL_KMIO_TEST, inBuffer, inBuffer.Length, inBuffer, inBuffer.Length,
                out int _ /* pBytesReturned */, IntPtr.Zero);
        }
    }
}
