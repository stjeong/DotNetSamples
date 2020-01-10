using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace KernelStructOffset
{
    public class WindowsHandleInfo : IDisposable
    {
        IntPtr _ptr = IntPtr.Zero;
        int _handleCount = 0;
        int _handleOffset = 0;

        public int HandleCount
        {
            get { return _handleCount; }
        }

        public WindowsHandleInfo()
        {
            Initialize();
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

        public _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX this[int index]
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

                    Span<_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX> handles = new Span<_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX>((_ptr + _handleOffset).ToPointer(), _handleCount);
                    return handles[index];
                    */

                    IntPtr entryPtr = IntPtr.Zero;

                    if (IntPtr.Size == 8)
                    {
                        IntPtr handleTable = new IntPtr(_ptr.ToInt64() + _handleOffset);
                        entryPtr = new IntPtr(handleTable.ToInt64() + sizeof(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX) * index);
                    }
                    else
                    {
                        IntPtr handleTable = new IntPtr(_ptr.ToInt32() + _handleOffset);
                        entryPtr = new IntPtr(handleTable.ToInt32() + sizeof(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX) * index);
                    }

                    _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX entry = 
                        (_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX)Marshal.PtrToStructure(entryPtr, typeof(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX));
                    return entry;
                }
            }
        }

        private void Initialize()
        {
            int guessSize = 4096;
            int requiredSize = 0;
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            while (true)
            {
                ret = NativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemExtendedHandleInformation, ptr, guessSize, out requiredSize);

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
                    typedef struct _SYSTEM_HANDLE_INFORMATION
                    {
                        ULONG HandleCount;
                        _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX Handles[1];
                    } SYSTEM_HANDLE_INFORMATION, *PSYSTEM_HANDLE_INFORMATION;
                    */

                    _handleCount = Marshal.ReadInt32(ptr);

                    _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX dummy = new _SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX();
                    _handleOffset = Marshal.OffsetOf(typeof(_SYSTEM_HANDLE_TABLE_ENTRY_INFO_EX), nameof(dummy.HandleValue)).ToInt32();
                    _ptr = ptr;
                    break;
                }

                Marshal.FreeHGlobal(ptr);
                break;
            }
        }
    }

    class FileNameFromHandleState : IDisposable
    {
        private ManualResetEvent _mr;
        private IntPtr _handle;
        private string _fileName;
        private bool _retValue;

        public IntPtr Handle
        {
            get
            {
                return _handle;
            }
        }

        public string FileName
        {
            get
            {
                return _fileName;
            }
            set
            {
                _fileName = value;
            }

        }

        public bool RetValue
        {
            get
            {
                return _retValue;
            }
            set
            {
                _retValue = value;
            }
        }

        public FileNameFromHandleState(IntPtr handle)
        {
            _mr = new ManualResetEvent(false);
            this._handle = handle;
        }

        public bool WaitOne(int wait)
        {
            return _mr.WaitOne(wait, false);
        }

        public void Set()
        {
            if (_mr == null)
            {
                return;
            }

            _mr.Set();
        }

        public void Dispose()
        {
            if (_mr != null)
            {
                _mr.Close();
                _mr = null;
            }
        }
    }
}
