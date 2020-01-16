using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDbgExt
{
    public static class Util
    {
        public static DateTime ToDateTime(long u)
        {
            bool tryGetDateTime(long ticks, DateTimeKind kind, out DateTime result) // Local Function C# 7.0
            {
                result = DateTime.MinValue;

                long checkMask = (long)kind << 0x3e;
                long masked = ticks & checkMask;

                if (masked == checkMask)
                {
                    result = new DateTime(ticks ^ checkMask, kind);
                    return true;
                }

                return false;
            };

            switch (u)
            {
                // Pattern matching C# 7.0
                case long value when tryGetDateTime(value, DateTimeKind.Utc, out DateTime time) == true: // out C# 7.0
                    return time;

                case long value when tryGetDateTime(value, DateTimeKind.Local, out DateTime time) == true: // out C# 7.0
                    return time;
            }

            return new DateTime(u, DateTimeKind.Unspecified);
        }

    }
}
