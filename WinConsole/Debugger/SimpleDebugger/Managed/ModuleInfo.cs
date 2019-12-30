using Microsoft.Diagnostics.Runtime.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleDebugger
{
    public struct ModuleInfo
    {
        public ulong ImageFileHandle;
        public ulong BaseOffset;
        public uint ModuleSize;
        public string ModuleName;
        public string ImageName;
        public uint CheckSum;
        public uint TimeDateStamp;

        public ModuleInfo(ulong imageFileHandle, ulong baseOffset, uint moduleSize, string moduleName, string imageName, uint checkSum, uint timeDateStamp)
        {
            ImageFileHandle = imageFileHandle;
            BaseOffset = baseOffset;
            ModuleSize = moduleSize;
            ModuleName = moduleName;
            ImageName = imageName;
            CheckSum = checkSum;
            TimeDateStamp = timeDateStamp;
        }
    }

    public struct ExceptionInfo
    {
        public EXCEPTION_RECORD64 Record;
        public bool FirstChance;

        public ExceptionInfo(EXCEPTION_RECORD64 record, uint firstChance)
        {
            Record = record;
            FirstChance = firstChance != 0;
        }
    }
}
