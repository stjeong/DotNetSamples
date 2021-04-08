using Microsoft.Win32;
using System;
using System.Management;
using System.Runtime.InteropServices;

namespace OSVersionInfo
{
    // C# - Environment.OSVersion의 문제점 및 윈도우 운영체제의 버전을 구하는 다양한 방법
    // ; https://www.sysnet.pe.kr/2/0/12589

    class Program
    {
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        extern static bool VerifyVersionInfoW(ref OSVERSIONINFOEXW lpVersionInformation, int dwTypeMask, ulong dwlConditionMask);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        extern static ulong VerSetConditionMask(ulong conditionMask, int typeMask, byte condition);

        [DllImport("ntdll.dll", CharSet = CharSet.Auto)]
        extern static int RtlGetVersion(ref OSVERSIONINFOEXW lpVersionInformation);

        static void Main(string[] args)
        {
            Console.WriteLine($"{nameof(Environment.OSVersion)}: {Environment.OSVersion.Version} (depends on app.manifest)");
            // Windows 10 - 6.2.9200.0

            string releaseId = Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId", "").ToString();
            Console.WriteLine($"Registry: {releaseId} (depends on osversion >= 10.0.0, and just only ReleaseId)");
            Console.WriteLine($"WMI: {GetOsVer()} (no dependencies, but slow)");

            Console.WriteLine($"RtlGetVersion: {GetRtlVersion()} (no dependencies)");

            Console.WriteLine($">= Windows 8: {IsWindows8OrGreater()} (no dependencies)");
            Console.WriteLine($">= Windows 10: {IsWindows10OrGreater()} (depends on app.manifest)");
            Console.WriteLine($">= 10.0.17063: {IsWindowsVersionOrGreater(10, 0, 17063)} (depends on app.manifest)");
            Console.WriteLine($">= 10.0.19042: {IsWindowsVersionOrGreater(10, 0, 19042)} (depends on app.manifest)");
            Console.WriteLine($">= 10.0.20000: {IsWindowsVersionOrGreater(10, 0, 20000)} (depends on app.manifest)");
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        struct OSVERSIONINFOEXW
        {
            public int dwOSVersionInfoSize;
            public int dwMajorVersion;
            public int dwMinorVersion;
            public int dwBuildNumber;
            public int dwPlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szCSDVersion;
            public UInt16 wServicePackMajor;
            public UInt16 wServicePackMinor;
            public UInt16 wSuiteMask;
            public byte wProductType;
            public byte wReserved;
        }

        static int VER_MINORVERSION = 0x0000001;
        static int VER_MAJORVERSION = 0x0000002;
        static int VER_BUILDNUMBER = 0x0000004;
        // static int VER_PLATFORMID = 0x0000008;
        // static int VER_SERVICEPACKMINOR = 0x0000010;
        // static int VER_SERVICEPACKMAJOR = 0x0000020;
        // static int VER_SUITENAME = 0x0000040;
        // static int VER_PRODUCT_TYPE = 0x0000080;

        // static int VER_EQUAL = 1;
        // static int VER_GREATER = 2;
        static int VER_GREATER_EQUAL = 3;
        // static int VER_LESS = 4;
        // static int VER_LESS_EQUAL = 5;
        // static int VER_AND = 6;
        // static int VER_OR = 7;

        static bool IsWindows8OrGreater()
        {
            return IsWindowsVersionOrGreater(6, 2, 0);
        }

        static bool IsWindows10OrGreater()
        {
            return IsWindowsVersionOrGreater(10, 0, 0);
        }

        static Version GetRtlVersion()
        {
            OSVERSIONINFOEXW info = new OSVERSIONINFOEXW();
            info.dwOSVersionInfoSize = Marshal.SizeOf(info);

            RtlGetVersion(ref info);

            return new Version(info.dwMajorVersion, info.dwMinorVersion, info.dwBuildNumber);
        }

        static bool IsWindowsVersionOrGreater(ushort wMajorVersion, ushort wMinorVersion, ushort wBuildVersion /* wServicePackMajor */)
        {
            OSVERSIONINFOEXW osvi = new OSVERSIONINFOEXW();
            osvi.dwOSVersionInfoSize = Marshal.SizeOf(osvi);

            int buildNumberMask = VER_BUILDNUMBER; /* VER_SERVICEPACKMAJOR */

            ulong dwlConditionMask = VerSetConditionMask(
                VerSetConditionMask(
                VerSetConditionMask(
                    0, VER_MAJORVERSION, (byte)VER_GREATER_EQUAL),
                       VER_MINORVERSION, (byte)VER_GREATER_EQUAL),
                       buildNumberMask, (byte)VER_GREATER_EQUAL);

            osvi.dwMajorVersion = wMajorVersion;
            osvi.dwMinorVersion = wMinorVersion;
            osvi.dwBuildNumber = wBuildVersion;
            // osvi.wServicePackMajor = wServicePackMajor;

            return VerifyVersionInfoW(ref osvi, VER_MAJORVERSION | VER_MINORVERSION | buildNumberMask, dwlConditionMask) != false;
        }

        private static ManagementObject GetMngObj(string className)
        {
            var wmi = new ManagementClass(className);

            foreach (var o in wmi.GetInstances())
            {
                var mo = (ManagementObject)o;
                if (mo != null) return mo;
            }

            return null;
        }

        public static string GetOsVer()
        {
            try
            {
                ManagementObject mo = GetMngObj("Win32_OperatingSystem");

                if (null == mo)
                    return string.Empty;

                return mo["Version"] as string;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
