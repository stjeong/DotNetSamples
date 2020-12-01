using Microsoft.Win32.SafeHandles;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace KeyPal
{
    // Enumerating container names of the strong name CSP
    // ; https://stackoverflow.com/questions/16658541/enumerating-container-names-of-the-strong-name-csp
    class Program
    {
        static void Main(string[] args)
        {
            CspProviderFlags flag = CspProviderFlags.UseMachineKeyStore;

            if (args.Length >= 1)
            {
                if (args[0].EndsWith(".pfx", StringComparison.OrdinalIgnoreCase) == true)
                {
                    Console.WriteLine(GetKeyContainerName(args[0]));
                    return;
                }

                flag = (CspProviderFlags)Enum.Parse(typeof(CspProviderFlags), args[0]);
            }

            foreach (var kc in KeyUtilities.EnumerateKeyContainers("Microsoft Strong Cryptographic Provider"))
            {
                CspParameters cspParams = new CspParameters();
                cspParams.KeyContainerName = kc;
                cspParams.Flags = flag;
                using (RSACryptoServiceProvider prov = new RSACryptoServiceProvider(cspParams))
                {
                    if (prov.CspKeyContainerInfo.Exportable)
                    {
                        var blob = prov.ExportCspBlob(true);
                        StrongNameKeyPair kp = new StrongNameKeyPair(prov.ExportCspBlob(false));
                        Console.WriteLine(kc + " pk length:" + kp.PublicKey.Length);
                    }
                }
                Console.WriteLine();
            }
        }

        // https://github.com/honzajscz/SnInstallPfx/blob/a7caed70d17ce8b0c244c27439193579438ca1c2/src/MSBuildCode/ResolveKeySourceTask.cs#L40
        private static string GetKeyContainerName(string pfxFilePath)
        {
            string currentUserName = Environment.UserDomainName + "\\" + Environment.UserName;
            byte[] userNameBytes = System.Text.Encoding.Unicode.GetBytes(currentUserName.ToLower(CultureInfo.InvariantCulture));

            byte [] keyBytes = File.ReadAllBytes(pfxFilePath);
            UInt64 hash = HashFromBlob(keyBytes);
            hash ^= HashFromBlob(userNameBytes); // modify it with the username hash, so each user would get different hash for the same key

            return "VS_KEY_" + hash.ToString("X016", CultureInfo.InvariantCulture);
        }

        // https://github.com/honzajscz/SnInstallPfx/blob/a7caed70d17ce8b0c244c27439193579438ca1c2/src/MSBuildCode/ResolveKeySourceTask.cs#L91
        private static UInt64 HashFromBlob(byte[] data)
        {
            UInt32 dw1 = 17339221;
            UInt32 dw2 = 19619429;
            UInt32 pos = 10803503;

            foreach (byte b in data)
            {
                UInt32 value = b ^ pos;
                pos *= 10803503;
                dw1 += ((value ^ dw2) * 15816943) + 17368321;
                dw2 ^= ((value + dw1) * 14984549) ^ 11746499;
            }
            UInt64 result = dw1;
            result <<= 32;
            result |= dw2;
            return result;
        }
    }
}

public static class KeyUtilities
{
    public static IList<string> EnumerateKeyContainers(string provider)
    {
        ProvHandle prov;
        if (!CryptAcquireContext(out prov, null, provider, PROV_RSA_FULL, CRYPT_MACHINE_KEYSET | CRYPT_VERIFYCONTEXT))
            throw new Win32Exception(Marshal.GetLastWin32Error());

        List<string> list = new List<string>();
        IntPtr data = IntPtr.Zero;
        try
        {
            int flag = CRYPT_FIRST;
            int len = 0;
            if (!CryptGetProvParam(prov, PP_ENUMCONTAINERS, IntPtr.Zero, ref len, flag))
            {
                if (Marshal.GetLastWin32Error() != ERROR_MORE_DATA)
                    throw new Win32Exception(Marshal.GetLastWin32Error());
            }

            data = Marshal.AllocHGlobal(len);
            do
            {
                if (!CryptGetProvParam(prov, PP_ENUMCONTAINERS, data, ref len, flag))
                {
                    if (Marshal.GetLastWin32Error() == ERROR_NO_MORE_ITEMS)
                        break;

                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                list.Add(Marshal.PtrToStringAnsi(data));
                flag = CRYPT_NEXT;
            }
            while (true);
        }
        finally
        {
            if (data != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(data);
            }

            prov.Dispose();
        }
        return list;
    }

    private sealed class ProvHandle : SafeHandleZeroOrMinusOneIsInvalid
    {
        public ProvHandle()
            : base(true)
        {
        }

        protected override bool ReleaseHandle()
        {
            return CryptReleaseContext(handle, 0);
        }

        [DllImport("advapi32.dll")]
        private static extern bool CryptReleaseContext(IntPtr hProv, int dwFlags);

    }

    const int PP_ENUMCONTAINERS = 2;
    const int PROV_RSA_FULL = 1;
    const int ERROR_MORE_DATA = 234;
    const int ERROR_NO_MORE_ITEMS = 259;
    const int CRYPT_FIRST = 1;
    const int CRYPT_NEXT = 2;
    const int CRYPT_MACHINE_KEYSET = 0x20;
    const int CRYPT_VERIFYCONTEXT = unchecked((int)0xF0000000);

    [DllImport("advapi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CryptAcquireContext(out ProvHandle phProv, string pszContainer, string pszProvider, int dwProvType, int dwFlags);

    [DllImport("advapi32.dll", SetLastError = true)]
    private static extern bool CryptGetProvParam(ProvHandle hProv, int dwParam, IntPtr pbData, ref int pdwDataLen, int dwFlags);
}