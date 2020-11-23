using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using WindowsPE;

#pragma warning disable IDE0051, CA1031

namespace PEFormatSample
{
    class Program
    {
        static int Main(string[] args)
        {
            //// Is DLL managed/unmanaged?
            //CheckDlls();

            //// Show exported functions from EXPORT Data Directory
            //ShowExportFunctions(true);
            //ShowExportFunctions(false);

            //// Show debug info from DEBUG Data Directory
            //ShowPdbInfo(true);
            //ShowPdbInfo(false);

            //// Show how to download PDB from Microsoft Symbol Server
            //DownloadPdbs(true);
            //DownloadPdbs(false);

            if (args.Length >= 2)
            {
                // ex)
                //  PEFormatSample pdb http://msdl.microsoft.com/download/symbols/ntdll.pdb/6DFD0B387E7941A587A3B64F824B1CAC1/ntdll.pdb
                if (args[0] == "pdb")
                {
                    string dirPath = Path.GetDirectoryName(typeof(Program).Assembly.Location);
                    Uri uri = new Uri(args[1]);
                    string localFilePath = Path.Combine(dirPath, uri.GetFileName());
                    DownloadPdbFile(uri, localFilePath);

                    bool downloaded = File.Exists(localFilePath);
                    Console.WriteLine($"{localFilePath}: {downloaded}");

                    return downloaded == true ? 0 : 1;
                }
            }

            return 0;
        }

        private static void DownloadPdbs(bool fromFile)
        {
            string rootPathToSave = Path.Combine(Environment.CurrentDirectory, "sym");
            if (Directory.Exists(rootPathToSave) == false)
            {
                Directory.CreateDirectory(rootPathToSave);
            }

            Process currentProcess = Process.GetCurrentProcess();
            foreach (ProcessModule pm in currentProcess.Modules)
            {
                Console.WriteLine($"[{pm.FileName}, 0x{pm.BaseAddress.ToString("x")}]");

                PEImage pe;

                if (fromFile == true)
                {
                    pe = PEImage.ReadFromFile(pm.FileName);
                }
                else
                {
                    pe = PEImage.ReadFromMemory(pm.BaseAddress, pm.ModuleMemorySize);
                }

                if (pe == null)
                {
                    Console.WriteLine("Failed to read images");
                    return;
                }

                string modulePath = pm.FileName;
                DownloadPdb(modulePath, rootPathToSave);

                Console.WriteLine();
            }
        }

        private static void DownloadPdb(string modulePath, string rootPathToSave)
        {
            if (File.Exists(modulePath) == false)
            {
                Console.WriteLine("NOT Found: " + modulePath);
                return;
            }

            PEImage pe = PEImage.ReadFromFile(modulePath);
            if (pe == null)
            {
                Console.WriteLine("Failed to read images");
                return;
            }

            Uri baseUri = new Uri("https://msdl.microsoft.com/download/symbols/");

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

                Console.WriteLine("\tFileName: " + pdbFileName);
                Console.WriteLine("\tPdbFileName: " + codeView.PdbFileName);
                Console.WriteLine("\tLocal: " + codeView.PdbLocalPath);
                Console.WriteLine("\tUri: " + codeView.PdbUriPath);

                string localPath = Path.Combine(rootPathToSave, codeView.PdbLocalPath);
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
                    Console.WriteLine("Found on local: " + pdbFileName);
                    continue;
                }

                if (CopyPdbFromLocal(modulePath, codeView.PdbFileName, localPath) == true)
                {
                    Console.WriteLine("Found on local: " + pdbFileName);
                    continue;
                }

                Uri target = new Uri(baseUri, codeView.PdbUriPath);
                Console.WriteLine(target);

                Uri pdbLocation = GetPdbLocation(target);

                if (pdbLocation == null)
                {
                    string underscorePath = ProbeWithUnderscore(target.AbsoluteUri);
                    pdbLocation = GetPdbLocation(new Uri(underscorePath));
                }

                if (pdbLocation != null)
                {
                    DownloadPdbFile(pdbLocation, localPath);
                }
                else
                {
                    Console.WriteLine("Not Found on symbol server: " + codeView.PdbFileName);
                }
            }

            Console.WriteLine();
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

        private static void ShowPdbInfo(bool fromFile)
        {
            Process currentProcess = Process.GetCurrentProcess();
            foreach (ProcessModule pm in currentProcess.Modules)
            {
                Console.WriteLine($"[{pm.FileName}, 0x{pm.BaseAddress.ToString("x")}]");

                PEImage pe;

                if (fromFile == true)
                {
                    pe = PEImage.ReadFromFile(pm.FileName);
                }
                else
                {
                    pe = PEImage.ReadFromMemory(pm.BaseAddress, pm.ModuleMemorySize);
                }

                if (pe == null)
                {
                    Console.WriteLine("Failed to read images");
                    return;
                }

                foreach (IMAGE_DEBUG_DIRECTORY debugDir in pe.EnumerateDebugDir())
                {
                    Console.WriteLine("\tDebugType: " + Enum.GetName(typeof(DebugDirectoryType), debugDir.Type));
                }

                foreach (CodeViewRSDS codeView in pe.EnumerateCodeViewDebugInfo())
                {
                    Console.WriteLine("\t\t" + codeView.PdbLocalPath);
                }

                Console.WriteLine();
            }
        }

        private static void CheckDlls()
        {
            Console.WriteLine("native.dll: IsManaged == " + PEImage.ReadFromFile("native.dll").IsManaged);
            Console.WriteLine("net20.dll: IsManaged == " + PEImage.ReadFromFile("net20.dll").IsManaged);
            Console.WriteLine("net40.dll: IsManaged == " + PEImage.ReadFromFile("net40.dll").IsManaged);
            Console.WriteLine("WindowsPE.dll: IsManaged == " + PEImage.ReadFromFile("WindowsPE.dll").IsManaged);
        }

        private static void ShowExportFunctions(bool fromFile)
        {
            Process currentProcess = Process.GetCurrentProcess();
            foreach (ProcessModule pm in currentProcess.Modules)
            {
                Console.WriteLine($"[{pm.FileName}, 0x{pm.BaseAddress.ToString("x")}]");

                PEImage pe;

                if (fromFile == true)
                {
                    pe = PEImage.ReadFromMemory(pm.BaseAddress, pm.ModuleMemorySize);
                }
                else
                {
                    pe = PEImage.ReadFromFile(pm.FileName);
                }

                if (pe == null)
                {
                    Console.WriteLine("Failed to read images");
                    return;
                }

                foreach (IMAGE_SECTION_HEADER section in pe.EnumerateSections())
                {
                    Console.WriteLine(section);
                }

                Console.WriteLine();

                // foreach (ExportFunctionInfo efi in pe.EnumerateExportFunctions())
                foreach (ExportFunctionInfo efi in pe.EnumerateExportFunctions().Take(5))
                {
                    Console.WriteLine("\t" + efi);
                }

                Console.WriteLine("\t...");

                Console.WriteLine();
            }
        }
    }

    static class Extension
    {
        public static string GetFileName(this Uri uri)
        {
            string path = uri.LocalPath;

            int pos1 = path.LastIndexOf('\\');
            int pos2 = path.LastIndexOf('/');

            int epos = Math.Max(pos1, pos2);
            return path.Substring(epos + 1);
        }
    }
}
