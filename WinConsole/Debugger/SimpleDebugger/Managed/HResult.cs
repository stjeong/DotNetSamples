using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleDebugger
{
    public enum HResult : uint
    {
        S_OK = 0,
        S_FALSE = 1U,
        E_PENDING = 0x8000000A,
        E_UNEXPECTED = 0x8000FFFF,
        E_FAIL = 0x80004005,
    }
}
