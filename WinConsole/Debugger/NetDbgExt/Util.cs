using Microsoft.Diagnostics.Runtime.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDbgExt
{
    public static class Util
    {
        public static bool CheckActiveTarget(IDebugControl dbgControl)
        {
            dbgControl.GetDebuggeeType(out DEBUG_CLASS debugClass, out DEBUG_CLASS_QUALIFIER qualifier);

            if (debugClass != DEBUG_CLASS.USER_WINDOWS ||
                (debugClass == DEBUG_CLASS.USER_WINDOWS && qualifier == DEBUG_CLASS_QUALIFIER.USER_WINDOWS_SMALL_DUMP))
            {
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"The current target is not a user-mode target.\n");
                dbgControl.Output(DEBUG_OUTPUT.NORMAL, $"DebuggeeType: {debugClass}, {qualifier}.\n");
                return false;
            }

            return true;
        }

        public static DateTime ToDateTime(long u) =>
            u switch
            {
                long value when TryGetDateTime(value, DateTimeKind.Utc, out DateTime time) == true => time,
                long value when TryGetDateTime(value, DateTimeKind.Local, out DateTime time) == true => time,
                _ => new DateTime(u, DateTimeKind.Unspecified)
            };

        static bool TryGetDateTime(long ticks, DateTimeKind kind, out DateTime result)
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
        }
    }
}
