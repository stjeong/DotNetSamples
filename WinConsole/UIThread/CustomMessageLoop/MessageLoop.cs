
using System;
using System.Threading;

namespace CustomMessageLoop
{
    public class MessageLoop : IDisposable
    {
        const uint QS_ALLPOSTMESSAGE = 0x100;
        const uint MWMO_INPUTAVAILABLE = 0x0004;

        uint _tid = 0;
        EventWaitHandle _ewh_Sync = new EventWaitHandle(true, EventResetMode.ManualReset);
        EventWaitHandle _ewh_Exit = new EventWaitHandle(true, EventResetMode.ManualReset);
        EventWaitHandle _ewh_Quit = new EventWaitHandle(false, EventResetMode.ManualReset);

        IntPtr[] _pHandles;

        public ApartmentState COMApartment => _uiThread?.GetApartmentState() ?? ApartmentState.STA;

        public void PostMessage(Win32Message msg)
        {
            PostMessage((uint)msg);
        }

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed == false)
            {
                if (disposing == true)
                {
                    _ewh_Quit.Set();
                    _ewh_Exit.WaitOne();
                }
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void PostMessage(uint msg)
        {
            if (_tid == 0)
            {
                return;
            }

            NativeMethods.PostThreadMessage((uint)_tid, msg, UIntPtr.Zero, IntPtr.Zero);
        }

        public void WaitForExit(int millisecondsTimeout = Timeout.Infinite)
        {
            _ewh_Exit.WaitOne(millisecondsTimeout);
        }

        public void SendMessage(Win32Message msg)
        {
            SendMessage((uint)msg);
        }

        public void SendMessage(uint msg)
        {
            _ewh_Sync.Reset();
            NativeMethods.PostThreadMessage((uint)_tid, msg, UIntPtr.Zero, IntPtr.Zero);
            _ewh_Sync.WaitOne();
        }

        Thread? _uiThread;
        bool _useBackgroundThread;
        ApartmentState _apartment;

        public MessageLoop() : this(true, ApartmentState.STA) { }

        public MessageLoop(bool useBackgroundThread) : this(useBackgroundThread, ApartmentState.STA) { }

        public MessageLoop(bool useBackgroundThread, ApartmentState apartment)
        {
            _pHandles = new IntPtr[1] { _ewh_Quit.GetSafeWaitHandle().DangerousGetHandle() };
            _useBackgroundThread = useBackgroundThread;
            _apartment = apartment;
        }

        public void Run()
        {
            _ewh_Exit.Reset();

            _uiThread = new Thread(Start);
            _uiThread.SetApartmentState(_apartment);
            _uiThread.IsBackground = _useBackgroundThread;
            _uiThread.Start();
        }

        public event EventHandler? Loaded;
        public event EventHandler? Closed;
        public event EventHandler<MessageEventArgs>? MessageArrived;

        protected virtual void OnLoad() { Loaded?.Invoke(this, EventArgs.Empty); }

        protected virtual void OnClose() { Closed?.Invoke(this, EventArgs.Empty); }

        protected virtual void WindowProc(MSG msg) { }

        void Start()
        {
            try
            {
                _tid = NativeMethods.GetCurrentThreadId();

                OnLoad();

                while (true)
                {
                    try
                    {
                        uint cObjects = (uint)_pHandles.Length;
                        uint result = NativeMethods.MsgWaitForMultipleObjectsEx(cObjects, _pHandles, 0xFFFFFFFF, QS_ALLPOSTMESSAGE,
                           MWMO_INPUTAVAILABLE);

                        if (result == cObjects)
                        {
                            // NativeMethods.PeekMessage(out MSG msg, IntPtr.Zero, 0, 0, PM_REMOVE);
                            NativeMethods.GetMessage(out MSG msg, IntPtr.Zero, 0, 0);
                            WindowProc(msg);
                            MessageArrived?.Invoke(this, new MessageEventArgs(msg));

                            switch (msg.message)
                            {
                                case (uint)Win32Message.WM_QUIT:
                                case (uint)Win32Message.WM_CLOSE:
                                    OnClose();
                                    return;
                            }

                            NativeMethods.DispatchMessage(ref msg);
                        }
                        else
                        {
                            OnClose();
                            return;
                        }
                    }
                    finally
                    {
                        _ewh_Sync.Set();
                    }
                }
            }
            finally
            {
                _disposed = true;
                _tid = 0;
                _ewh_Exit.Set();
            }
        }
    }
}
