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
            UnpackDummyApp();

            if (args.Length < 2)
            {
                return;
            }

            string typeName = args[0];
            string moduleFullName = args[1];
            string moduleFileName = Path.GetFileNameWithoutExtension(moduleFullName);

            string rootPathToSave = Path.Combine(Environment.CurrentDirectory, "sym");

            using (UserDebugger debugger = new UserDebugger())
            {
                debugger.SetOutputText(false);
                debugger.FlushCallbacks();

                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "DummyApp.exe";
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.LoadUserProfile = false;

                Process child = Process.Start(psi);
                bool attached = false;

                try
                {
                    DownloadPdb(child, rootPathToSave, moduleFullName);

                    debugger.ModuleLoaded += (ModuleInfo modInfo) =>
                    {
                        if (modInfo.ModuleName.IndexOf(moduleFileName, StringComparison.OrdinalIgnoreCase) == -1)
                        {
                            return;
                        }

                        RunDTSetCommand(debugger, rootPathToSave, moduleFullName, typeName);

                        debugger.SetOutputText(false);
                        debugger.Execute(DEBUG_OUTCTL.IGNORE, "q", DEBUG_EXECUTE.NOT_LOGGED);
                    };

                    if (debugger.AttachTo(child.Id) == false)
                    {
                        Console.WriteLine("Failed to attach");
                        return;
                    }

                    attached = true;
                    debugger.WaitForEvent();
                }
                finally
                {
                    if (attached == true)
                    {
                        debugger.Detach();
                    }

                    try
                    {
                        child.Kill();
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
                string cmd = $"dt {Path.GetFileNameWithoutExtension(moduleFullName)}!{typeName}";
                int result = debugger.Execute(cmd);
                if (result != (int)HResult.S_OK)
                {
                    Console.WriteLine("failed to run command: " + cmd);
                    return;
                }
            }
        }

        private static string DownloadPdb(Process child, string rootPathToSave, string moduleFullName)
        {
            foreach (ProcessModule module in child.Modules)
            {
                if (module.FileName.EndsWith($"{moduleFullName}", StringComparison.OrdinalIgnoreCase) == true)
                {
                    string modulePath = module.FileName;
                    PEImage pe = PEImage.ReadFromFile(modulePath);
                    return DownloadPdb(modulePath, rootPathToSave);
                }
            }

            return "";
        }

        private static string DownloadPdb(string modulePath, string rootPathToSave)
        {
            if (File.Exists(modulePath) == false)
            {
                Console.WriteLine("NOT Found: " + modulePath);
                return null;
            }

            PEImage pe = PEImage.ReadFromFile(modulePath);
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
