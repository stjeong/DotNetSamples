using System;
using System.Runtime.InteropServices;

namespace CustomMessageLoop
{
    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PeekMessage(out MSG message, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

        
        [DllImport(@"user32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        internal static extern int GetMessage(out MSG message, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

        [DllImport(@"user32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        internal static extern bool TranslateMessage(ref MSG message);
        
        [DllImport(@"user32.dll", CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        internal static extern IntPtr DispatchMessage(ref MSG message);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool PostThreadMessage(uint threadId, uint msg, UIntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        internal static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        internal static extern void PostQuitMessage(int nExitCode);

        [DllImport("user32.dll")]
        internal static extern uint MsgWaitForMultipleObjects(uint nCount, IntPtr[]? pHandles, bool fWaitAll, uint dwMilliseconds, uint dwWakeMask);

        [DllImport("user32.dll")]
        internal static extern uint MsgWaitForMultipleObjectsEx(uint nCount, IntPtr[]? pHandles, uint dwMilliseconds, uint dwWakeMask, uint dwFlags);

    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public long x;
        public long y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
    }

    public class MessageEventArgs : EventArgs
    {
        MSG _msg;
        public MSG Message => _msg;

        public MessageEventArgs(MSG msg)
        {
            _msg = msg;
        }
    }
}
