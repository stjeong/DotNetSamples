using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WindowsPE
{
    public struct ExportFunctionInfo
    {
        public string Name;
        public ushort NameOrdinal;
        public uint RvaAddress;

        /// <summary>
        /// Biased of NameOrdinal
        /// </summary>
        public uint Ordinal;

        public override string ToString()
        {
            return $"{Name} at 0x{RvaAddress.ToString("x")}";
        }
    }

}
