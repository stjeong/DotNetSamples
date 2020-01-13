using Microsoft.Diagnostics.Runtime.Interop;
using SimpleDebugger;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using WindowsPE;

namespace DisplayStruct
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                return;
            }

            if (args.Length == 2)
            {
                UnpackDummyApp();
            }

            string typeName = args[0];
            string moduleFullName = args[1];
            string moduleFileName = Path.GetFileNameWithoutExtension(moduleFullName);
            int pid = 0;
            Process child = null;
            string proxyExePath = "DummyApp.exe";

            if (args.Length >= 3)
            {
                if (Int32.TryParse(args[2], out pid) == false)
                {
                    proxyExePath = args[2];
                }
            }

            string rootPathToSave = Path.Combine(Environment.CurrentDirectory, "sym");

            using (UserDebugger debugger = new UserDebugger())
            {
                debugger.SetOutputText(false);
                debugger.FlushCallbacks();

                if (pid == 0)
                {
                    ProcessStartInfo psi = new ProcessStartInfo();
                    psi.FileName = proxyExePath;
                    psi.UseShellExecute = false;
                    psi.CreateNoWindow = true;
                    psi.LoadUserProfile = false;

                    child = Process.Start(psi);
                    pid = child.Id;
                }

                bool attached = false;

                try
                {
                    debugger.ModuleLoaded += (ModuleInfo modInfo) =>
                    {
                        if (modInfo.ModuleName.IndexOf(moduleFileName, StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            return;
                        }

                        byte[] buffer;
                        int readBytes = debugger.ReadMemory(modInfo.BaseOffset, modInfo.ModuleSize, out buffer);
                        if (readBytes != modInfo.ModuleSize)
                        {
                            return;
                        }

                        PEImage.DownloadPdb(modInfo.ModuleName, buffer, new IntPtr((long)modInfo.BaseOffset), (int)modInfo.ModuleSize, rootPathToSave);

                        RunDTSetCommand(debugger, rootPathToSave, moduleFullName, typeName);

                        debugger.SetOutputText(false);
                        debugger.Execute(DEBUG_OUTCTL.IGNORE, (child == null) ? "q" : "qd", DEBUG_EXECUTE.NOT_LOGGED);
                    };

                    if (debugger.AttachTo(pid) == false)
                    {
                        Console.WriteLine("Failed to attach");
                        return;
                    }

                    attached = true;
                    debugger.WaitForEvent(DEBUG_WAIT.DEFAULT, 1000 * 5);
                }
                finally
                {
                    if (attached == true)
                    {
                        debugger.Detach();
                    }

                    try
                    {
                        if (child != null)
                        {
                            child.Kill();
                        }
                    }
                    catch { }
                }
            }
        }

        private static void UnpackDummyApp()
        {
            UnpackDummyAppFromRes("DummyApp.exe");
        }

        private static void UnpackDummyAppFromRes(string fileName)
        {
            if (File.Exists(fileName) == true)
            {
                return;
            }

            Type pgType = typeof(Program);

            using (Stream manifestResourceStream =
                pgType.Assembly.GetManifestResourceStream($@"{pgType.Namespace}.files.{fileName}"))
            {
                using (BinaryReader br = new BinaryReader(manifestResourceStream))
                {
                    byte[] buf = new byte[br.BaseStream.Length];
                    br.Read(buf, 0, buf.Length);

                    File.WriteAllBytes(fileName, buf);
                }
            }
        }

        private static void RunDTSetCommand(UserDebugger debugger, string rootPathToSave, string moduleFullName, string typeName)
        {
            {
                string cmd = ".sympath \"" + rootPathToSave + "\"";
                int result = debugger.Execute(DEBUG_OUTCTL.IGNORE, cmd, DEBUG_EXECUTE.NOT_LOGGED);
                if (result != (int)HResult.S_OK)
                {
                    Console.WriteLine("failed to run command: " + cmd);
                    return;
                }
            }

            {
                string cmd = $".reload /s /f {moduleFullName}";
                int result = debugger.Execute(DEBUG_OUTCTL.IGNORE, cmd, DEBUG_EXECUTE.NOT_LOGGED);
                if (result != (int)HResult.S_OK)
                {
                    Console.WriteLine("failed to run command: " + cmd);
                    return;
                }
            }

            debugger.SetOutputText(true);
            {
                string cmd = $"dt {typeName}";
                int result = debugger.Execute(cmd);
                if (result != (int)HResult.S_OK)
                {
                    Console.WriteLine("failed to run command: " + cmd);
                    return;
                }
            }
        }


    }
}
