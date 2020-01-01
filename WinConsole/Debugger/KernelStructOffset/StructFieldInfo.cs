using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace KernelStructOffset
{
    public struct StructFieldInfo
    {
        public readonly string Name;
        public readonly int Offset;
        public readonly string Type;

        public StructFieldInfo(int offset, string name, string type)
        {
            Name = name;
            Offset = offset;
            Type = type;
        }
    }
}
