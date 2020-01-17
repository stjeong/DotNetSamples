using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace WindowsPE
{
    public class PEImage
    {
        IMAGE_DOS_HEADER _dosHeader;
        public IMAGE_DOS_HEADER DosHeader
        {
            get { return _dosHeader; }
        }

        IMAGE_FILE_HEADER _fileHeader;
        IMAGE_OPTIONAL_HEADER32 _optionalHeader32;
        IMAGE_OPTIONAL_HEADER64 _optionalHeader64;

        IMAGE_SECTION_HEADER[] _sections;

        bool _is64BitHeader;

        byte[] _bufferCached;

        IntPtr _baseAddress;
        public IntPtr BaseAddress
        {
            get { return _baseAddress; }
        }
        
        string _filePath;
        public string ModulePath
        {
            get { return _filePath; }
        }

        int _memorySize;
        public int MemorySize
        {
            get { return _memorySize; }
        }

        static bool IsValidNTHeaders(int signature)
        {
            // PE 헤더임을 확인 (IMAGE_NT_SIGNATURE == 0x00004550)
            // if (signature[0] == 0x50 && signature[1] == 0x45 && signature[2] == 0 && signature[3] == 0)
            if (signature == 0x00004550)
            {
                return true;
            }

            return false;
        }

        public bool IsManaged
        {
            get
            {
                return CLRRuntimeHeader.VirtualAddress != 0;
            }
        }

        public IMAGE_DATA_DIRECTORY CLRRuntimeHeader
        {
            get
            {
                if (_is64BitHeader == true)
                {
                    return _optionalHeader64.CLRRuntimeHeader;
                }
                else
                {
                    return _optionalHeader32.CLRRuntimeHeader;
                }
            }
        }

        public IMAGE_DATA_DIRECTORY Export
        {
            get
            {
                if (_is64BitHeader == true)
                {
                    return _optionalHeader64.ExportTable;
                }
                else
                {
                    return _optionalHeader32.ExportTable;
                }
            }
        }

        public IMAGE_DATA_DIRECTORY Debug
        {
            get
            {
                if (_is64BitHeader == true)
                {
                    return _optionalHeader64.Debug;
                }
                else
                {
                    return _optionalHeader32.Debug;
                }
            }
        }

        public IEnumerable<IMAGE_SECTION_HEADER> EnumerateSections()
        {
            for (int i = 0; i < _sections?.Length; i++)
            {
                yield return _sections[i];
            }
        }

        public ExportFunctionInfo GetExportFunction(string functionName)
        {
            ExportFunctionInfo[] functions = GetExportFunctions();

            for (int i = 0; i < functions?.Length; i++)
            {
                if (functions[i].Name == functionName)
                {
                    return functions[i];
                }
            }

            return default;
        }

        public IEnumerable<ExportFunctionInfo> EnumerateExportFunctions()
        {
            ExportFunctionInfo[] functions = GetExportFunctions();

            for (int i = 0; i < functions?.Length; i++)
            {
                yield return functions[i];
            }
        }

        public unsafe ExportFunctionInfo[] GetExportFunctions()
        {
            if (Export.VirtualAddress == 0)
            {
                return null;
            }

            IMAGE_SECTION_HEADER section = GetSection((int)Export.VirtualAddress);

            GetSafeBuffer(0, (uint)section.VirtualAddress + (uint)section.SizeOfRawData, out BufferPtr buffer);
            List<ExportFunctionInfo> list = new List<ExportFunctionInfo>();

            try
            {
                IntPtr exportDir = GetSafeBuffer(buffer, Export.VirtualAddress);
                IMAGE_EXPORT_DIRECTORY dir = (IMAGE_EXPORT_DIRECTORY)Marshal.PtrToStructure(exportDir, typeof(IMAGE_EXPORT_DIRECTORY));

                IntPtr nameListPtr = GetSafeBuffer(buffer, dir.AddressOfNames);
                UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)nameListPtr.ToPointer(), dir.NumberOfNames * sizeof(int));
                BinaryReader br = new BinaryReader(ums);

                for (int i = 0; i < dir.NumberOfNames; i++)
                {
                    uint namePos = br.ReadUInt32();
                    IntPtr namePtr = GetSafeBuffer(buffer, namePos);

                    ExportFunctionInfo efi;
                    efi.Name = Marshal.PtrToStringAnsi(namePtr);

                    efi.NameOrdinal = GetSafeBuffer(buffer, dir.AddressOfNameOrdinals).ReadUInt16ByIndex(i);
                    efi.RvaAddress = GetSafeBuffer(buffer, dir.AddressOfFunctions).ReadUInt32ByIndex(efi.NameOrdinal);

                    efi.Ordinal = efi.NameOrdinal + dir.Base;

                    list.Add(efi);
                }
            }
            finally
            {
                buffer.Clear();
            }

            return list.ToArray();
        }

#if _INCLUDE_MANAGED_STRUCTS
        public IEnumerable<CodeViewRSDS> EnumerateCodeViewDebugInfo()
        {
            foreach (IMAGE_DEBUG_DIRECTORY debugDir in EnumerateDebugDir())
            {
                if (debugDir.Type != (uint)DebugDirectoryType.IMAGE_DEBUG_TYPE_CODEVIEW)
                {
                    continue;
                }

                IntPtr debugDirPtr = GetSafeBuffer(debugDir.AddressOfRawData, debugDir.SizeOfData, out BufferPtr buffer);

                try
                {
                    yield return debugDir.GetCodeViewHeader(debugDirPtr);
                }
                finally
                {
                    buffer.Clear();
                }
            }
        }
#endif

        IntPtr GetSafeBuffer(uint rva, uint size, out BufferPtr buffer)
        {
            buffer = null;

            if (_baseAddress == IntPtr.Zero)
            {
                int startAddress = Rva2Raw((int)rva);
                int endAddress = startAddress + (int)size;

                buffer = new BufferPtr(endAddress);

                using (FileStream fs = new FileStream(_filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    BinaryReader br = new BinaryReader(fs);
                    br.Read(buffer.Buffer, 0, buffer.Length);
                }
            }
            else if (_bufferCached != null)
            {
                buffer = new BufferPtr(_bufferCached.Length);
                Array.Copy(_bufferCached, 0, buffer.Buffer, 0, _bufferCached.Length);
            }

            return GetSafeBuffer(buffer, rva);
        }

        IntPtr GetSafeBuffer(BufferPtr buffer, uint rva)
        {
            IntPtr ptr;

            if (_bufferCached != null)
            {
                ptr = buffer.GetPtr((int)rva);
            }
            else if (buffer != null)
            {
                int startAddress = Rva2Raw((int)rva);
                ptr = buffer.GetPtr(startAddress);
            }
            else
            {
                ptr = _baseAddress + (int)rva;
            }

            return ptr;
        }

        public IEnumerable<IMAGE_DEBUG_DIRECTORY> EnumerateDebugDir()
        {
            if (Debug.VirtualAddress == 0)
            {
                yield break;
            }

            IntPtr debugDirPtr = GetSafeBuffer(Debug.VirtualAddress, Debug.Size, out BufferPtr buffer);

            try
            {
                IMAGE_DEBUG_DIRECTORY safeObj = new IMAGE_DEBUG_DIRECTORY();
                int sizeOfDir = Marshal.SizeOf(safeObj);

                int count = (int)Debug.Size / sizeOfDir;

                for (int i = 0; i < count; i++)
                {
                    IMAGE_DEBUG_DIRECTORY dir = (IMAGE_DEBUG_DIRECTORY)Marshal.PtrToStructure(debugDirPtr, typeof(IMAGE_DEBUG_DIRECTORY));
                    yield return dir;

                    debugDirPtr += sizeOfDir;
                }
            }
            finally
            {
                buffer.Clear();
            }
        }

        private int Rva2Raw(int virtualAddress)
        {
            IMAGE_SECTION_HEADER section = GetSection(virtualAddress);
            return virtualAddress - section.VirtualAddress + section.PointerToRawData;
        }

        private IMAGE_SECTION_HEADER GetSection(int virtualAddress)
        {
            for (int i = 0; i < _sections.Length; i++)
            {
                IMAGE_SECTION_HEADER section = _sections[i];

                int startAddr = section.VirtualAddress;
                int endAddr = section.VirtualAddress + section.PhysicalAddressOrVirtualSize;

                if (startAddr <= virtualAddress && virtualAddress <= endAddr)
                {
                    return section;
                }
            }

            return default;
        }

        unsafe static PEImage ReadPEHeader(BinaryReader br)
        {
            PEImage image = new PEImage();

            // IMAGE_DOS_HEADER 를 읽어들이고,
            IMAGE_DOS_HEADER dosHeader = br.Read<IMAGE_DOS_HEADER>();
            {
                if (dosHeader.IsValid == false)
                {
                    return null;
                }

                image._dosHeader = dosHeader;
            }

            // IMAGE_NT_HEADERS - signature 를 읽어들이고,
            {
                br.BaseStream.Position = dosHeader.e_lfanew;
                if (IsValidNTHeaders(br.ReadInt32()) == false)
                {
                    return null;
                }
            }

            // IMAGE_NT_HEADERS - IMAGE_FILE_HEADER를 읽어들임
            IMAGE_FILE_HEADER ntFileHeader = br.Read<IMAGE_FILE_HEADER>();
            image._fileHeader = ntFileHeader;

            /*
            AnyCPU로 빌드된 .NET Image인 경우,
            OptionalHeader의 첫 번째 필드 Magic 값이 원래는 0x010B로 PE32였지만, 실행 후 메모리에 매핑되면서 0x020B로 변경

            따라서 64비트 이미지로 매핑되었음을 판단하려면 Magic 필드로 판단해야 함.
            */

            ushort magic = br.PeekUInt16();
            if (magic == (ushort)MagicType.IMAGE_NT_OPTIONAL_HDR64_MAGIC)
            {
                image._is64BitHeader = true;
            }

            // ushort optionalHeaderSize = ntFileHeader.SizeOfOptionalHeader;
            // optionalHeaderSize
            // 32bit PE == 0xe0(224)bytes
            // 64bit PE == 0xF0(240)bytes

            if (image._is64BitHeader == false)
            {
                image._optionalHeader32 = br.Read<IMAGE_OPTIONAL_HEADER32>();
            }
            else
            {
                image._optionalHeader64 = br.Read<IMAGE_OPTIONAL_HEADER64>();
            }

            {
                List<IMAGE_SECTION_HEADER> sections = new List<IMAGE_SECTION_HEADER>();

                for (int i = 0; i < image._fileHeader.NumberOfSections; i++)
                {
                    IMAGE_SECTION_HEADER sectionHeader = br.Read<IMAGE_SECTION_HEADER>();
                    sections.Add(sectionHeader);
                }

                image._sections = sections.ToArray();
            }

            return image;
        }

        public unsafe static PEImage FromLoadedModule(string moduleName)
        {
            foreach (ProcessModule pm in Process.GetCurrentProcess().Modules)
            {
                if (pm.ModuleName.Equals(moduleName, StringComparison.OrdinalIgnoreCase) == true)
                {
                    PEImage image = ReadFromMemory(pm.BaseAddress, pm.ModuleMemorySize);
                    image._filePath = pm.FileName;
                    image._baseAddress = pm.BaseAddress;
                    image._memorySize = pm.ModuleMemorySize;
                    return image;
                }
            }

            return null;
        }

        public unsafe static PEImage ReadFromMemory(IntPtr baseAddress, int memorySize)
        {
            UnmanagedMemoryStream ums = new UnmanagedMemoryStream((byte*)baseAddress.ToPointer(), memorySize);
            return ReadFromMemory(ums, baseAddress, memorySize);
        }

        public unsafe static PEImage ReadFromMemory(byte[] buffer, IntPtr baseAddress, int memorySize)
        {
            MemoryStream ms = new MemoryStream(buffer);
            PEImage image = ReadFromMemory(ms, baseAddress, memorySize);
            image._bufferCached = buffer;

            return image;
        }

        public unsafe static PEImage ReadFromMemory(Stream stream, IntPtr baseAddress, int memorySize)
        {
            BinaryReader br = new BinaryReader(stream);

            PEImage image = ReadPEHeader(br);
            if (image == null)
            {
                return null;
            }

            image._baseAddress = baseAddress;
            image._memorySize = memorySize;

            return image;
        }

        public unsafe static PEImage ReadFromFile(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                BinaryReader br = new BinaryReader(fs);
                PEImage image = ReadPEHeader(br);
                if (image == null)
                {
                    return null;
                }

                image._filePath = filePath;

                return image;
            }
        }

        public static string DownloadPdb(string modulePath, byte[] buffer, IntPtr baseOffset, int imageSize, string rootPathToSave)
        {
            PEImage pe = PEImage.ReadFromMemory(buffer, baseOffset, imageSize);

            if (pe == null)
            {
                Console.WriteLine("Failed to read images");
                return null;
            }

            return pe.DownloadPdb(modulePath, rootPathToSave);
        }

        public string DownloadPdb(string modulePath, string rootPathToSave)
        {
            Uri baseUri = new Uri("https://msdl.microsoft.com/download/symbols/");
            string pdbDownloadedPath = string.Empty;

            foreach (CodeViewRSDS codeView in EnumerateCodeViewDebugInfo())
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

        private static string ProbeWithUnderscore(string path)
        {
            path = path.Remove(path.Length - 1);
            path = path.Insert(path.Length, "_");
            return path;
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
    }
}
