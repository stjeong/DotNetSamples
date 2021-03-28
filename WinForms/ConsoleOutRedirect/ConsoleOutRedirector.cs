using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace WindowsFormsApp1
{
    public class ConsoleOutRedirector : IDisposable
    {
        public ConsoleOutRedirector()
        {
            NativeMethods.AllocConsole();

            InitConsole();
        }

        private void InitConsole()
        {
            // Use AllocConsole instead of Visual Studio Output
            // https://stackoverflow.com/questions/41624103/console-out-output-is-showing-in-output-window-needed-in-allocconsole
            IntPtr stdHandle = NativeMethods.CreateFile(
                    "CONOUT$",
                    NativeMethods.GENERIC_WRITE,
                    NativeMethods.FILE_SHARE_WRITE,
                    0, NativeMethods.OPEN_EXISTING, 0, 0
                );

            SafeFileHandle safeFileHandle = new SafeFileHandle(stdHandle, true);
            FileStream fileStream = new FileStream(safeFileHandle, FileAccess.Write);
            Encoding encoding = System.Text.Encoding.GetEncoding(NativeMethods.MY_CODE_PAGE);
            StreamWriter standardOutput = new StreamWriter(fileStream, encoding);
            standardOutput.AutoFlush = true;
            Console.SetOut(standardOutput);
        }

        public void Dispose()
        {
            NativeMethods.FreeConsole();
        }
    }

    static class NativeMethods
    {
        internal const int MY_CODE_PAGE = 437;
        internal const uint GENERIC_WRITE = 0x40000000;
        internal const uint FILE_SHARE_WRITE = 0x2;
        internal const uint OPEN_EXISTING = 0x3;

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern IntPtr CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            uint lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            uint hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AllocConsole();

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool FreeConsole();
    }
}
