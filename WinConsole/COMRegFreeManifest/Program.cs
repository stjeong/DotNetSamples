using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace COMRegFreeManifest
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 1)
            {
                Console.WriteLine("COMRegFreeManifest.exe [com_dll_path]");
                return;
            }

            string path = args[0];

            string currentPath = Environment.CurrentDirectory;
            string dllManifestPath = Path.Combine(currentPath, Path.GetFileName(path));
            string exeManifestPath = Path.Combine(currentPath, "Sample.exe");

            string impLibPath = Path.Combine(Path.GetTempPath(), "test.dll");

            try
            {
                if (RunTlbImp(path, impLibPath) == false)
                {
                    Console.WriteLine("TLBIMP not work");
                    return;
                }

                byte[] buf = File.ReadAllBytes(impLibPath);
                Assembly asm = Assembly.ReflectionOnlyLoad(buf);

                ExportDllManifest(asm, impLibPath, dllManifestPath);
                ExportExeManifest(asm, impLibPath, dllManifestPath, exeManifestPath);
            }
            finally
            {
                if (File.Exists(impLibPath) == true)
                {
                    File.Delete(impLibPath);
                }
            }
        }
        private static void ExportExeManifest(Assembly asm, string impLibPath, string dllManifestPath, string exeManifestPath)
        {
            string manifestTemplate = @"<?xml version=""1.0"" encoding=""utf-8""?>
<asmv1:assembly manifestVersion=""1.0"" xmlns=""urn:schemas-microsoft-com:asm.v1"" xmlns:asmv1=""urn:schemas-microsoft-com:asm.v1"" xmlns:asmv2=""urn:schemas-microsoft-com:asm.v2"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">
  <assemblyIdentity version=""1.0.0.0"" name=""MyApplication.app""/>

  <trustInfo xmlns=""urn:schemas-microsoft-com:asm.v2"">
    <security>
      <requestedPrivileges xmlns=""urn:schemas-microsoft-com:asm.v3"">
        <requestedExecutionLevel level=""asInvoker"" uiAccess=""false"" />
      </requestedPrivileges>
    </security>
  </trustInfo>
  
  <compatibility xmlns=""urn:schemas-microsoft-com:compatibility.v1"">
    <application>
    </application>
  </compatibility>
  
  <dependency>
    <dependentAssembly asmv2:dependencyType=""install"" asmv2:codebase=""{0}.manifest"">
      <assemblyIdentity name=""{0}"" version=""{1}"" type=""win32"" />
    </dependentAssembly>
  </dependency>

</asmv1:assembly>
";


            string dllName = Path.GetFileName(dllManifestPath);
            string version = asm.GetName().Version.ToString();

            string manifestText = string.Format(manifestTemplate, dllName, version);
            File.WriteAllText(exeManifestPath + ".manifest", manifestText);
        }

        private static void ExportDllManifest(Assembly asm, string impLibPath, string dllManifestPath)
        {
            string txt = GetManifestContents(asm, dllManifestPath);
            File.WriteAllText(dllManifestPath + ".manifest", txt);
        }

        static string GetManifestContents(Assembly asm, string dllManifestPath)
        { 
            string manifestTemplate = @"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>
<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">

  <assemblyIdentity version=""{0}"" name=""{1}"" type=""win32"">
  </assemblyIdentity>
  
  <file name=""{1}"">
    <comClass clsid=""{2}"" threadingModel=""Apartment"" tlbid=""{3}"" />
    <typelib tlbid=""{3}"" version=""{4}"" helpdir="""" resourceid=""0"" />
  </file>

  {5}
</assembly>
";
            string interfaceTemplate = @"
  <comInterfaceExternalProxyStub name=""{0}"" 
                                 iid=""{1}"" 
                                 proxyStubClsid32=""{{00020424-0000-0000-C000-000000000046}}"" 
                                 baseInterface=""{2}"" 
                                 tlbid=""{3}"">
  </comInterfaceExternalProxyStub>
";

            string version = asm.GetName().Version.ToString();
            string dllName = Path.GetFileName(dllManifestPath);
            string clsid = $"{{{GetClsid(asm)}}}";
            string tlbid = $"{{{GetAssemblyAttr(asm, typeof(GuidAttribute))}}}";
            string tlbVersion = GetTypeLibVersion(asm);

            StringBuilder sb = new StringBuilder();

            foreach (Type type in asm.GetTypes())
            {
                if (type.IsInterface == false || HasCoClassAttribute(type))
                {
                    continue;
                }

                string interfaceName = type.Name;
                string interfaceIid = $"{{{GetGuid(type)}}}";
                string baseInterfaceIid = GetBaseInterfaceIID(type);

                string interfaceText = string.Format(interfaceTemplate, interfaceName, interfaceIid, baseInterfaceIid, tlbid);
                sb.AppendLine(interfaceText);
            }

            string manifestText = string.Format(manifestTemplate, version, dllName, clsid, tlbid, tlbVersion, sb.ToString());

            return manifestText;
        }

        private static string GetBaseInterfaceIID(Type type)
        {
            foreach (CustomAttributeData attr in type.GetCustomAttributesData())
            {
                if (attr.Constructor.DeclaringType == typeof(TypeLibTypeAttribute))
                {
                    TypeLibTypeFlags flags = (TypeLibTypeFlags)attr.ConstructorArguments[0].Value;
                    if (flags.HasFlag(TypeLibTypeFlags.FDual))
                    {
                        return "{00020400-0000-0000-C000-000000000046}";
                    }
                    else
                    { 
                        return "{00000000-0000-0000-C000-000000000046}";
                    }
                }
            }
            return "";
        }

        private static string GetTypeLibVersion(Assembly asm)
        {
            foreach (CustomAttributeData attr in asm.GetCustomAttributesData())
            {
                if (attr.Constructor.DeclaringType == typeof(TypeLibVersionAttribute))
                {
                    return $"{attr.ConstructorArguments[0].Value}.{attr.ConstructorArguments[1].Value}";
                }
            }

            return "";
        }

        private static bool HasCoClassAttribute(Type type)
        {
            foreach (CustomAttributeData attr in type.GetCustomAttributesData())
            {
                if (attr.Constructor.DeclaringType == typeof(CoClassAttribute))
                {
                    return true;
                }
            }

            return false;
        }

        private static string GetClsid(Assembly asm)
        {
            foreach (Type type in asm.GetTypes())
            {
                foreach (CustomAttributeData attr in type.GetCustomAttributesData())
                {
                    if (attr.Constructor.DeclaringType == typeof(CoClassAttribute))
                    {
                        return GetGuid(attr.ConstructorArguments[0].Value as Type);
                    }
                }
            }

            return "";
        }

        private static string GetGuid(Type type)
        {
            foreach (CustomAttributeData attr in type.GetCustomAttributesData())
            {
                if (attr.Constructor.DeclaringType == typeof(GuidAttribute))
                {
                    return attr.ConstructorArguments[0].Value as string;
                }
            }

            return "";
        }

        private static string GetAssemblyAttr(Assembly asm, Type targetType)
        {
            foreach (CustomAttributeData attr in asm.GetCustomAttributesData())
            {
                if (attr.Constructor.DeclaringType == targetType)
                {
                    return attr.ConstructorArguments[0].Value as string;
                }
            }

            return "";
        }

        private static bool RunTlbImp(string path, string impLibPath)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = GetTlbImpPath("tlbimp.exe");
            psi.Arguments = $"\"{path}\" /out:\"{impLibPath}\"";
            psi.CreateNoWindow = true;
            psi.UseShellExecute = false;

            Process newProc = Process.Start(psi);
            newProc.WaitForExit();

            return File.Exists(impLibPath);
        }

        private static string GetTlbImpPath(string toolName)
        {
            using (RegistryKey regKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\WOW6432Node\Microsoft\Microsoft SDKs\NETFXSDK", false))
            {
                foreach (string version in regKey.GetSubKeyNames())
                {
                    using (RegistryKey versionKey = regKey.OpenSubKey(version, false))
                    {
                        foreach (string tool in versionKey.GetSubKeyNames())
                        {
                            using (RegistryKey toolKey = versionKey.OpenSubKey(tool))
                            {
                                string path = toolKey.GetValue("InstallationFolder") as string;

                                string tlbimpPath = Path.Combine(path, toolName);
                                if (File.Exists(tlbimpPath) == true)
                                {
                                    return tlbimpPath;
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }
    }
}
