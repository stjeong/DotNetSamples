using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KernelStructOffset
{
    /// <summary>
    /// Supported since Windows 8/2012
    /// </summary>
    public sealed class ProcessHandleInfo : IDisposable
    {
        IntPtr _ptr = IntPtr.Zero;
        int _handleCount = 0;
        int _handleOffset = 0;

        public int HandleCount
        {
            get { return _handleCount; }
        }

        public ProcessHandleInfo(int pid)
        {
            Initialize(pid);
        }

        public void Dispose()
        {
            if (_ptr == IntPtr.Zero)
            {
                return;
            }

            Marshal.FreeHGlobal(_ptr);
            _ptr = IntPtr.Zero;
        }

        public _PROCESS_HANDLE_TABLE_ENTRY_INFO this[int index]
        {
            get
            {
                if (_ptr == IntPtr.Zero)
                {
                    return default;
                }

                unsafe
                {
                    /*
                    Span<_PROCESS_HANDLE_TABLE_ENTRY_INFO> handles = new Span<_PROCESS_HANDLE_TABLE_ENTRY_INFO>((_ptr + _handleOffset).ToPointer(), _handleCount);
                    return handles[index];
                    */

                    IntPtr entryPtr;

                    if (IntPtr.Size == 8)
                    {
                        IntPtr handleTable = new IntPtr(_ptr.ToInt64() + _handleOffset);
                        entryPtr = new IntPtr(handleTable.ToInt64() + sizeof(_PROCESS_HANDLE_TABLE_ENTRY_INFO) * index);
                    }
                    else
                    {
                        IntPtr handleTable = new IntPtr(_ptr.ToInt32() + _handleOffset);
                        entryPtr = new IntPtr(handleTable.ToInt32() + sizeof(_PROCESS_HANDLE_TABLE_ENTRY_INFO) * index);
                    }

                    _PROCESS_HANDLE_TABLE_ENTRY_INFO entry =
                        (_PROCESS_HANDLE_TABLE_ENTRY_INFO)Marshal.PtrToStructure(entryPtr, typeof(_PROCESS_HANDLE_TABLE_ENTRY_INFO));
                    return entry;
                }
            }
        }

        private void Initialize(int pid)
        {
            int guessSize = 4096;
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);
            IntPtr processHandle = NativeMethods.OpenProcess(
                ProcessAccessRights.PROCESS_QUERY_INFORMATION | ProcessAccessRights.PROCESS_DUP_HANDLE, false, pid);
            if (processHandle == IntPtr.Zero)
            {
                return;
            }

            while (true)
            {
                ret = NativeMethods.NtQueryInformationProcess(processHandle, PROCESS_INFORMATION_CLASS.ProcessHandleInformation, ptr, guessSize, out int requiredSize);

                if (ret == NT_STATUS.STATUS_INFO_LENGTH_MISMATCH)
                {
                    Marshal.FreeHGlobal(ptr);
                    guessSize = requiredSize;
                    ptr = Marshal.AllocHGlobal(guessSize);
                    continue;
                }

                if (ret == NT_STATUS.STATUS_SUCCESS)
                {
                    /*
                        typedef struct _PROCESS_HANDLE_SNAPSHOT_INFORMATION {
                            ULONG_PTR NumberOfHandles;
                            ULONG_PTR Reserved;
                            PROCESS_HANDLE_TABLE_ENTRY_INFO Handles[1];
                        } PROCESS_HANDLE_SNAPSHOT_INFORMATION, * PPROCESS_HANDLE_SNAPSHOT_INFORMATION;
                    */

                    _handleCount = Marshal.ReadIntPtr(ptr).ToInt32();

#pragma warning disable IDE0059 // Unnecessary assignment of a value
                    _PROCESS_HANDLE_SNAPSHOT_INFORMATION dummy = new _PROCESS_HANDLE_SNAPSHOT_INFORMATION();
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                    _handleOffset = Marshal.OffsetOf(typeof(_PROCESS_HANDLE_SNAPSHOT_INFORMATION), nameof(dummy.Handles)).ToInt32();
                    _ptr = ptr;
                    break;
                }

                Marshal.FreeHGlobal(ptr);
                break;
            }
        }
    }

}
