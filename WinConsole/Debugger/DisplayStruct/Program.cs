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

                        DownloadPdb(modInfo.ModuleName, buffer, new IntPtr((long)modInfo.BaseOffset), (int)modInfo.ModuleSize, rootPathToSave);

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

        private static string DownloadPdb(string modulePath, byte [] buffer, IntPtr baseOffset, int imageSize, string rootPathToSave)
        {
            PEImage pe = PEImage.ReadFromMemory(buffer, baseOffset, imageSize);
            if (pe == null)
            {
                Console.WriteLine("Failed to read images");
                return null;
            }

            Uri baseUri = new Uri("https://msdl.microsoft.com/download/symbols/");
            string pdbDownloadedPath = string.Empty;

            foreach (CodeViewRSDS codeView in pe.EnumerateCodeViewDebugInfo())
            {
                if (string.IsNullOrEmpty(codeView.PdbFileName) == true)
                {
                    continue;
                }

                string pdbFileName = codeView.PdbFileName;
                if (Path.IsPathRooted(codeView.PdbFileName) == true)
                {
                    pdbFileName = Path.GetFileName(codeView.PdbFileName);
                }

                string localPath = Path.Combine(rootPathToSave, pdbFileName);
                string localFolder = Path.GetDirectoryName(localPath);

                if (Directory.Exists(localFolder) == false)
                {
                    try
                    {
                        Directory.CreateDirectory(localFolder);
                    }
                    catch (DirectoryNotFoundException)
                    {
                        Console.WriteLine("NOT Found on local: " + codeView.PdbLocalPath);
                        continue;
                    }
                }

                if (File.Exists(localPath) == true)
                {
                    if (Path.GetExtension(localPath).Equals(".pdb", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        pdbDownloadedPath = localPath;
                    }

                    continue;
                }

                if (CopyPdbFromLocal(modulePath, codeView.PdbFileName, localPath) == true)
                {
                    continue;
                }

                Uri target = new Uri(baseUri, codeView.PdbUriPath);
                Uri pdbLocation = GetPdbLocation(target);

                if (pdbLocation == null)
                {
                    string underscorePath = ProbeWithUnderscore(target.AbsoluteUri);
                    pdbLocation = GetPdbLocation(new Uri(underscorePath));
                }

                if (pdbLocation != null)
                {
                    DownloadPdbFile(pdbLocation, localPath);

                    if (Path.GetExtension(localPath).Equals(".pdb", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        pdbDownloadedPath = localPath;
                    }
                }
                else
                {
                    Console.WriteLine("Not Found on symbol server: " + codeView.PdbFileName);
                }
            }

            return pdbDownloadedPath;
        }

        private static bool CopyPdbFromLocal(string modulePath, string pdbFileName, string localTargetPath)
        {
            if (File.Exists(pdbFileName) == true)
            {
                File.Copy(pdbFileName, localTargetPath);
                return File.Exists(localTargetPath);
            }

            string fileName = Path.GetFileName(pdbFileName);
            string pdbPath = Path.Combine(Environment.CurrentDirectory, fileName);

            if (File.Exists(pdbPath) == true)
            {
                File.Copy(pdbPath, localTargetPath);
                return File.Exists(localTargetPath);
            }

            pdbPath = Path.ChangeExtension(modulePath, ".pdb");
            if (File.Exists(pdbPath) == true)
            {
                File.Copy(pdbPath, localTargetPath);
                return File.Exists(localTargetPath);
            }

            return false;
        }

        private static string ProbeWithUnderscore(string path)
        {
            path = path.Remove(path.Length - 1);
            path = path.Insert(path.Length, "_");
            return path;
        }

        private static void DownloadPdbFile(Uri target, string pathToSave)
        {
            System.Net.HttpWebRequest req = System.Net.WebRequest.Create(target) as System.Net.HttpWebRequest;

            using (System.Net.HttpWebResponse resp = req.GetResponse() as System.Net.HttpWebResponse)
            using (FileStream fs = new FileStream(pathToSave, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            using (BinaryWriter bw = new BinaryWriter(fs))
            {
                BinaryReader reader = new BinaryReader(resp.GetResponseStream());
                long contentLength = resp.ContentLength;

                while (contentLength > 0)
                {
                    byte[] buffer = new byte[4096];
                    int readBytes = reader.Read(buffer, 0, buffer.Length);
                    bw.Write(buffer, 0, readBytes);

                    contentLength -= readBytes;
                }
            }
        }

        private static Uri GetPdbLocation(Uri target)
        {
            System.Net.HttpWebRequest req = System.Net.WebRequest.Create(target) as System.Net.HttpWebRequest;
            req.Method = "HEAD";

            try
            {
                using (System.Net.HttpWebResponse resp = req.GetResponse() as System.Net.HttpWebResponse)
                {
                    return resp.ResponseUri;
                }
            }
            catch (System.Net.WebException)
            {
                return null;
            }
        }
    }
}
