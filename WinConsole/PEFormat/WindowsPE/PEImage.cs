﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
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
        IMAGE_COR20_HEADER? _corHeader;

        IMAGE_SECTION_HEADER[] _sections;

        bool _is64BitHeader;
        public bool Is64Bitness
        {
            get { return _is64BitHeader; }
        }

        byte[] _bufferCached;

        IntPtr _baseAddress;
        public IntPtr BaseAddress
        {
            get { return _baseAddress; }
        }

        bool _readFromFile = false;

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
                return CLRRuntimeHeaderDirectory.VirtualAddress != 0;
            }
        }

        public IMAGE_DATA_DIRECTORY CLRRuntimeHeaderDirectory
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

        public IMAGE_DATA_DIRECTORY ExportDirectory
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

        public IMAGE_DATA_DIRECTORY ResourceTable
        {
            get
            {
                if (_is64BitHeader == true)
                {
                    return _optionalHeader64.ResourceTable;
                }
                else
                {
                    return _optionalHeader32.ResourceTable;
                }
            }
        }

        public IEnumerable<IMAGE_SECTION_HEADER> EnumerateSections()
        {
            return _sections;
        }

        public IEnumerable<(string, IMAGE_DATA_DIRECTORY)> EnumerateDirectories()
        {
            object targetHeader = _optionalHeader32;
            if (_is64BitHeader == true)
            {
                targetHeader = _optionalHeader64;
            }
            
            IEnumerable<(string, IMAGE_DATA_DIRECTORY)> fields = targetHeader.GetType().GetFields(System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public)
                .Where(x => x.FieldType == typeof(IMAGE_DATA_DIRECTORY))
                .Select(x => (x.Name, (IMAGE_DATA_DIRECTORY)x.GetValue(targetHeader)));

            foreach (var field in fields)
            {
                yield return  field;
            }
        }

        public void ShowHeader()
        {
            Console.WriteLine($"File Type: {_fileHeader.GetFileType()}");
            Console.WriteLine("FILE HEADER VALUES");
            Console.WriteLine($"{_fileHeader.Machine,8:X} machine {_fileHeader.GetMachineType()}");
            Console.WriteLine($"{_fileHeader.NumberOfSections,8:X} number of sections");
            Console.WriteLine($"{_fileHeader.TimeDateStamp,8:X} time date stamp {_fileHeader.GetTimeDateStamp()}");
            Console.WriteLine();

            Console.WriteLine($"{_fileHeader.PointerToSymbolTable,8:X} file pointer to symbol table");
            Console.WriteLine($"{_fileHeader.NumberOfSymbols,8:X} number of symbols");
            Console.WriteLine($"{_fileHeader.Characteristics,8:X} characteristics");
            foreach (string text in _fileHeader.GetCharacteristics())
            {
                Console.WriteLine($"\t   {text}");
            }
            Console.WriteLine();

            Console.WriteLine("OPTIONAL HEADER VALUES");
            if (_is64BitHeader == true)
            {
                ShowOptionalHeader64();
            }
            else
            {
                ShowOptionalHeader32();
            }

            Console.WriteLine();
            Console.WriteLine();

            int number = 1;
            foreach (var item in this.EnumerateSections())
            {
                Console.WriteLine($"SECTION HEADER #{number}");
                Console.WriteLine($"{item.GetName(),8} name");
                Console.WriteLine($"{item.PhysicalAddressOrVirtualSize,8:X} virtual size");
                Console.WriteLine($"{item.VirtualAddress,8:X} virtual address");
                Console.WriteLine($"{item.SizeOfRawData,8:X} size of raw data");
                Console.WriteLine($"{item.PointerToRawData,8:X} pointer to raw data");
                Console.WriteLine($"{item.PointerToRelocations,8:X} pointer to relocations");
                Console.WriteLine($"{item.PointerToLineNumbers,8:X} pointer to line numbers");
                Console.WriteLine($"{item.NumberOfRelocations,8:X} number of relocations");
                Console.WriteLine($"{item.NumberOfLineNumbers,8:X} number of line numbers");
                Console.WriteLine($"{item.Characteristics,8:X} flags");
                foreach (string text in item.GetCharacteristics())
                {
                    Console.WriteLine($"\t   {text}");
                }
                Console.WriteLine();
                number++;
            }

            Console.WriteLine();

            ShowDebugInfo();
        }

        private void ShowDebugInfo()
        {
            List<IMAGE_DEBUG_DIRECTORY> list = new List<IMAGE_DEBUG_DIRECTORY>();
            list.AddRange(EnumerateDebugDir());
            Console.WriteLine($"Debug Directories({list.Count})");
            Console.WriteLine($"    {"Type",-11}{"Size",-9}{"Address",-9}{"Pointer",-7}");

            foreach (IMAGE_DEBUG_DIRECTORY debugDir in list)
            {
                Console.Write($"    {debugDir.GetDebugType(),-7}");
                Console.Write($"{debugDir.SizeOfData,8:X}");
                Console.Write($"{debugDir.AddressOfRawData,12:X}");
                Console.Write($"{debugDir.PointerToRawData,9:X}");

                if (debugDir.Type == (uint)DebugDirectoryType.IMAGE_DEBUG_TYPE_CODEVIEW)
                {
                    CodeViewRSDS codeView = debugDir.GetCodeViewHeader(GetSafeBuffer(debugDir.AddressOfRawData, debugDir.SizeOfData, out BufferPtr buffer));
                    Console.Write($"    Format: {codeView.GetSignature()}, {codeView.Age}, {codeView.PdbFileName}");
                }
                Console.WriteLine();
            }
        }

        private void ShowOptionalHeader32()
        {
            Console.WriteLine("Not yet implemented");
        }

        private void ShowOptionalHeader64()
        {
            Console.WriteLine($"{_optionalHeader64.Magic,8:X} magic #");
            string version = $"{_optionalHeader64.MajorLinkerVersion,2}.{_optionalHeader64.MinorLinkerVersion,02:D2}";
            Console.WriteLine($"{version,8:X} linker version");
            Console.WriteLine($"{_optionalHeader64.SizeOfCode,8:X} size of code");
            Console.WriteLine($"{_optionalHeader64.SizeOfInitializedData,8:X} size of initialized data");
            Console.WriteLine($"{_optionalHeader64.SizeOfUninitializedData,8:X} size of uninitialized data");
            Console.WriteLine($"{_optionalHeader64.AddressOfEntryPoint,8:X} address of entry point");
            Console.WriteLine($"{_optionalHeader64.BaseOfCode,8:X} base of code");
            Console.WriteLine($"\t ----- new -----");

            Console.WriteLine($"{_optionalHeader64.ImageBase,16:X16} image base");
            Console.WriteLine($"{_optionalHeader64.SectionAlignment,8:X} section alignment");
            Console.WriteLine($"{_optionalHeader64.FileAlignment,8:X} file alignment");
            Console.WriteLine($"{(ushort)_optionalHeader64.Subsystem,8:X} subsystem ({IMAGE_OPTIONAL_HEADER64.GetSubsystem((ushort)_optionalHeader64.Subsystem)})");
            version = $"{_optionalHeader64.MajorOperatingSystemVersion,2}.{_optionalHeader64.MinorOperatingSystemVersion,02:D2}";
            Console.WriteLine($"{version,8:X} operating system version");
            version = $"{_optionalHeader64.MajorImageVersion,2}.{_optionalHeader64.MinorImageVersion,02:D2}";
            Console.WriteLine($"{version,8:X} image version");
            version = $"{_optionalHeader64.MajorSubsystemVersion,2}.{_optionalHeader64.MinorSubsystemVersion,02:D2}";
            Console.WriteLine($"{version,8:X} subsystem version");
            Console.WriteLine($"{_optionalHeader64.SizeOfImage,8:X} size of image");
            Console.WriteLine($"{_optionalHeader64.SizeOfHeaders,8:X} size of headers");
            Console.WriteLine($"{_optionalHeader64.CheckSum,8:X} checksum");
            Console.WriteLine($"{_optionalHeader64.SizeOfStackReserve,16:X16} size of stack reserve");
            Console.WriteLine($"{_optionalHeader64.SizeOfStackCommit,16:X16} size of stack commit");
            Console.WriteLine($"{_optionalHeader64.SizeOfHeapReserve,16:X16} size of heap reserve");
            Console.WriteLine($"{_optionalHeader64.SizeOfHeapCommit,16:X16} size of heap commit");
            Console.WriteLine($"{_optionalHeader64.DllCharacteristics,8:X} DLL characteristics");
            foreach (string text in _optionalHeader64.GetDllCharacteristics())
            {
                Console.WriteLine($"\t   {text}");
            }

            foreach (var (name, dir) in EnumerateDirectories())
            {
                Console.WriteLine($"{dir.VirtualAddress,8:X} [{dir.Size,8:X}] address [size] of {name} Directory");
            }
        }

        public IEnumerable<VTableFixups> EnumerateVTableFixups()
        {
            IMAGE_COR20_HEADER corHeader = GetClrDirectoryHeader();
            VTableFixups[] vtfs = Reads<VTableFixups>(corHeader.VTableFixups.VirtualAddress, corHeader.VTableFixups.Size);
            return vtfs;
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
            return GetExportFunctions();
        }

        public unsafe byte[] ReadBytes(uint rvaAddress, int nBytes)
        {
            byte[] byteBuffer = new byte[nBytes];

            IMAGE_SECTION_HEADER section = GetSection((int)rvaAddress);
            GetSafeBuffer(0, (uint)section.EndAddress, out BufferPtr buffer);

            try
            {
                IntPtr bytePos = GetSafeBuffer(buffer, rvaAddress);
                int maxRead = Math.Min((int)(rvaAddress + nBytes), (int)section.EndAddress);

                for (int i = (int)rvaAddress; i < maxRead; i++)
                {
                    int pos = (int)(i - rvaAddress);
                    byteBuffer[pos] = bytePos.ReadByte(pos);
                }
            }
            finally
            {
                buffer.Clear();
            }

            return byteBuffer;
        }

        public unsafe T Read<T>(uint rvaAddress) where T : struct
        {
            IMAGE_SECTION_HEADER section = GetSection((int)rvaAddress);
            GetSafeBuffer(0, (uint)section.VirtualAddress + (uint)section.SizeOfRawData, out BufferPtr buffer);

            try
            {
                IntPtr bytePos = GetSafeBuffer(buffer, rvaAddress);
                return (T)Marshal.PtrToStructure(bytePos, typeof(T));
            }
            finally
            {
                buffer.Clear();
            }
        }

        public unsafe T[] Reads<T>(uint rvaAddress, uint totalSize) where T : struct
        {
            IMAGE_SECTION_HEADER section = GetSection((int)rvaAddress);
            GetSafeBuffer(0, (uint)section.VirtualAddress + (uint)section.SizeOfRawData, out BufferPtr buffer);

            List<T> list = new List<T>();

            uint entrySize = (uint)Marshal.SizeOf(default(T));
            uint count = totalSize / entrySize;

            try
            {
                for (uint i = 0; i < count; i++)
                {
                    IntPtr bytePos = GetSafeBuffer(buffer, rvaAddress + (i * entrySize));
                    list.Add((T)Marshal.PtrToStructure(bytePos, typeof(T)));
                }
            }
            finally
            {
                buffer.Clear();
            }

            return list.ToArray();
        }

        public IMAGE_COR20_HEADER GetClrDirectoryHeader()
        {
            if (CLRRuntimeHeaderDirectory.VirtualAddress == 0)
            {
                return default;
            }

            if (_corHeader == null)
            {
                _corHeader = Read<IMAGE_COR20_HEADER>(CLRRuntimeHeaderDirectory.VirtualAddress);
            }

            return _corHeader.Value;
        }

        static ExportFunctionInfo[] EmptyExportFunctionInfo = new ExportFunctionInfo[0];

        public unsafe ExportFunctionInfo[] GetExportFunctions()
        {
            if (ExportDirectory.VirtualAddress == 0)
            {
                return EmptyExportFunctionInfo;
            }

            IMAGE_SECTION_HEADER section = GetSection((int)ExportDirectory.VirtualAddress);

            GetSafeBuffer(0, (uint)section.VirtualAddress + (uint)section.SizeOfRawData, out BufferPtr buffer);
            List<ExportFunctionInfo> list = new List<ExportFunctionInfo>();

            try
            {
                IntPtr exportDirPos = GetSafeBuffer(buffer, ExportDirectory.VirtualAddress);
                IMAGE_EXPORT_DIRECTORY dir = (IMAGE_EXPORT_DIRECTORY)Marshal.PtrToStructure(exportDirPos, typeof(IMAGE_EXPORT_DIRECTORY));

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

        IntPtr GetSafeBuffer(uint rva, uint size, out BufferPtr buffer)
        {
            buffer = null;

            if (_readFromFile == true)
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

        _IMAGE_RESOURCE_DIRECTORY _resourceRoot;
        _IMAGE_RESOURCE_DIRECTORY ResourceRoot
        {
            get
            {
                if (_resourceRoot == null)
                {
                    LoadResourceTable();
                }

                return _resourceRoot;
            }
        }

        void LoadResourceTable()
        {
            if (ResourceTable.VirtualAddress == 0)
            {
                return;
            }

            IntPtr resourceTablePtr = GetSafeBuffer(ResourceTable.VirtualAddress, ResourceTable.Size, out BufferPtr buffer);

            try
            {
                _resourceRoot = BuildResourceEntries(resourceTablePtr, resourceTablePtr);
            }
            finally
            {
                buffer.Clear();
            }
        }

        private _IMAGE_RESOURCE_DIRECTORY BuildResourceEntries(IntPtr rootDirectoryPtr, IntPtr itemPtr)
        {
            try
            {
                // .NET 4.5.1 or later, .NET Standard 1.2 or later
                // IMAGE_RESOURCE_DIRECTORY header = Marshal.PtrToStructure<IMAGE_RESOURCE_DIRECTORY>(resourceTablePtr);

                IMAGE_RESOURCE_DIRECTORY resDir = (IMAGE_RESOURCE_DIRECTORY)Marshal.PtrToStructure(itemPtr, typeof(IMAGE_RESOURCE_DIRECTORY));
                itemPtr += IMAGE_RESOURCE_DIRECTORY.StructSize;

                int numberOfEntries = resDir.NumberOfNamedEntries + resDir.NumberOfIdEntries;
                _IMAGE_RESOURCE_DIRECTORY node = new _IMAGE_RESOURCE_DIRECTORY(numberOfEntries);

                for (int i = 0; i < numberOfEntries; i++)
                {
                    IMAGE_RESOURCE_DIRECTORY_ENTRY entry = (IMAGE_RESOURCE_DIRECTORY_ENTRY)Marshal.PtrToStructure(itemPtr, typeof(IMAGE_RESOURCE_DIRECTORY_ENTRY));

                    _IMAGE_RESOURCE_DIRECTORY_ENTRY item = new _IMAGE_RESOURCE_DIRECTORY_ENTRY(entry, rootDirectoryPtr);
                    node.Entries[i] = item;

                    if (entry.IsDirectory)
                    {
                        IntPtr childPtr = entry.GetDataPtr(rootDirectoryPtr);
                        item.Next = BuildResourceEntries(rootDirectoryPtr, childPtr);
                    }
                    else
                    {
                        IntPtr childPtr = entry.GetDataPtr(rootDirectoryPtr);
                        IMAGE_RESOURCE_DATA_ENTRY data = (IMAGE_RESOURCE_DATA_ENTRY)Marshal.PtrToStructure(childPtr, typeof(IMAGE_RESOURCE_DATA_ENTRY));
                        item.Data = new _IMAGE_RESOURCE_DATA_ENTRY(data);
                    }

                    itemPtr += IMAGE_RESOURCE_DIRECTORY_ENTRY.StructSize;
                }

                return node;

            }
            catch { }

            return null;
        }

        public VS_VERSION_INFO FindVersionInfo(int lcid = 0)
        {
            foreach (var item in this.ResourceRoot.Entries)
            {
                if (item.Id == ResourceTypeId.RT_VERSION)
                {
                    var data = item.Next.Entries[0].Next.FindLcidEntry(lcid);

                    IntPtr dataPtr = GetSafeBuffer(data.OffsetToData, data.Size, out BufferPtr buffer);
                    try
                    {
                        VS_VERSION_INFO versionInfo = VS_VERSION_INFO.Parse(dataPtr, data.Size);
                        return versionInfo;
                    }
                    finally
                    {
                        buffer.Clear();
                    }
                }
            }

            return null;
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
                    image._readFromFile = false;
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

                image._readFromFile = true;
                image._filePath = filePath;
                image._baseAddress = new IntPtr((image._is64BitHeader == true) ? (long)image._optionalHeader64.ImageBase
                                        : image._optionalHeader32.ImageBase);

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
