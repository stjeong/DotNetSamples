using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace EnumHandles
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

        public SYSTEM_HANDLE_ENTRY this[int index]
        {
            get
            {
                if (_ptr == IntPtr.Zero)
                {
                    return default;
                }

                unsafe
                {
                    Span<SYSTEM_HANDLE_ENTRY> handles = new Span<SYSTEM_HANDLE_ENTRY>((_ptr + _handleOffset).ToPointer(), _handleCount);
                    return handles[index];
                }
            }
        }

        private void Initialize()
        {
            int guessSize = 1024;
            int requiredSize = 0;
            NT_STATUS ret;

            IntPtr ptr = Marshal.AllocHGlobal(guessSize);

            while (true)
            {
                ret = NativeMethods.NtQuerySystemInformation(SYSTEM_INFORMATION_CLASS.SystemHandleInformation, ptr, guessSize, out requiredSize);

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
                        SYSTEM_HANDLE Handles[1];
                    } SYSTEM_HANDLE_INFORMATION, *PSYSTEM_HANDLE_INFORMATION;
                    */

                    _handleCount = Marshal.ReadInt32(ptr);
                    _handleOffset = Marshal.OffsetOf(typeof(SYSTEM_HANDLE_INFORMATION), "Handles").ToInt32();
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
