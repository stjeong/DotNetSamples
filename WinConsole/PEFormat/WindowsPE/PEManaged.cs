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

    public struct CodeViewRSDS
    {
        public uint CvSignature;
        public Guid Signature;
        public uint Age;
        public string PdbFileName;

        public string PdbLocalPath
        {
            get
            {
                string fileName = Path.GetFileName(PdbFileName);

                string uid = Signature.ToString("N") + Age;
                return Path.Combine(fileName, uid, fileName);
            }
        }

        public string PdbUriPath
        {
            get
            {
                string fileName = Path.GetFileName(PdbFileName);

                string uid = Signature.ToString("N") + Age;
                return $"{fileName}/{uid}/{fileName}";
            }
        }
    }
}
