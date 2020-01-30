using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace CustomMessageLoop
{
    public static class NativeMethods
    {
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

    public enum Win32Message : uint
    {
        WM_CLOSE = 0x0010,
        WM_USER = 0x0400,
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
