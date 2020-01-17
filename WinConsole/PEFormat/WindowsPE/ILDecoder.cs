using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WindowsPE
{
    public static class ILDecoder
    {
        public static bool IsMethodDef(long tokenValue)
        {
            long lowerValue = (tokenValue & 0x00000000FFFFFFFF);
            if (lowerValue != tokenValue)
            {
                return false;
            }

            if (((int)lowerValue & (int)CorTokenType.mdtMethodDef) == (int)CorTokenType.mdtMethodDef)
            {
                return true;
            }

            return false;
        }
    }
}
