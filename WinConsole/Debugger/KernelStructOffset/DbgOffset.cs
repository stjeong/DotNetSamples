using KernelStructOffset;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace KernelStructOffset
{
    public class DbgOffset
    {
        static Dictionary<string, DbgOffset> _cache = new Dictionary<string, DbgOffset>();

        Dictionary<string, StructFieldInfo> _fieldDict = new Dictionary<string, StructFieldInfo>();

        public static DbgOffset Get(string typeName)
        {
            return Get(typeName, "ntdll.dll");
        }

        public static DbgOffset Get(string typeName, string moduleName)
        {
            return Get(typeName, moduleName, 0);
        }

        public static DbgOffset Get(string typeName, string moduleName, int pid)
        {
            return Get(typeName, moduleName, (pid == 0) ? null : pid.ToString());
        }

        public static DbgOffset Get(string typeName, string moduleName, string targetExePath)
        {
            if (_cache.ContainsKey(typeName) == true)
            {
                return _cache[typeName];
            }

            List<StructFieldInfo> list = GetList(typeName, moduleName, targetExePath);
            if (list == null)
            {
                return null;
            }

            DbgOffset instance = new DbgOffset(list);
            _cache.Add(typeName, instance);

            return _cache[typeName];
        }

        private DbgOffset(List<StructFieldInfo> list)
        {
            _fieldDict = ListToDict(list);
        }

        public IntPtr GetPointer(IntPtr baseAddress, string fieldName)
        {
            if (_fieldDict.ContainsKey(fieldName) == false)
            {
                return IntPtr.Zero;
            }

            return baseAddress + _fieldDict[fieldName].Offset;
        }

        public unsafe bool TryRead<T>(IntPtr baseAddress, string fieldName, out T value) where T: struct
        {
            value = default(T);

            if (_fieldDict.ContainsKey(fieldName) == false)
            {
                return false;
            }

            IntPtr address = baseAddress + _fieldDict[fieldName].Offset;
            value = (T)Marshal.PtrToStructure(address, typeof(T));

            return true;
        }

        public int this[string fieldName]
        {
            get
            {
                if (_fieldDict.ContainsKey(fieldName) == false)
                {
                    return -1;
                }

                return _fieldDict[fieldName].Offset;
            }
        }

        private static Dictionary<string, StructFieldInfo> ListToDict(List<StructFieldInfo> list)
        {
            Dictionary<string, StructFieldInfo> dict = new Dictionary<string, StructFieldInfo>();

            foreach (StructFieldInfo item in list)
            {
                dict.Add(item.Name, item);
            }

            return dict;
        }

        private static List<StructFieldInfo> GetList(string typeName, string moduleName, string pidOrPath)
        {
            UnpackDisplayStructApp();

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = "DisplayStruct.exe";
            psi.UseShellExecute = false;
            psi.Arguments = $"{typeName} {moduleName}" + ((string.IsNullOrEmpty(pidOrPath) == true) ? "" : $" \"{pidOrPath}\"");
            psi.CreateNoWindow = true;
            psi.RedirectStandardOutput = true;
            psi.LoadUserProfile = false;

            Process child = Process.Start(psi);
            string text = child.StandardOutput.ReadToEnd();

            child.WaitForExit();

            return ParseOffset(text);
        }

        private static List<StructFieldInfo> ParseOffset(string text)
        {
            List<StructFieldInfo> list = new List<StructFieldInfo>();
            StringReader sr = new StringReader(text);

            while (true)
            {
                string line = sr.ReadLine();
                if (line == null)
                {
                    break;
                }

                int offsetEndPos = 0;
                int offset = ReadOffset(line, out offsetEndPos);
                if (offsetEndPos == -1 || offset == -1)
                {
                    continue;
                }

                int nameEndPos = 0;
                string name = ReadFieldName(line, offsetEndPos, out nameEndPos);
                if (string.IsNullOrEmpty(name) == true || nameEndPos == -1)
                {
                    continue;
                }

                string type = line.Substring(nameEndPos).Trim();

                StructFieldInfo sfi = new StructFieldInfo(offset, name, type);
                list.Add(sfi);
            }

            return list;
        }

        private static string ReadFieldName(string line, int offsetEndPos, out int nameEndPos)
        {
            nameEndPos = line.IndexOf(":", offsetEndPos);
            if (nameEndPos == -1)
            {
                return null;
            }

            string result = line.Substring(offsetEndPos, nameEndPos - offsetEndPos).Trim();
            nameEndPos += 1;

            return result;
        }

        private static int ReadOffset(string line, out int pos)
        {
            pos = -1;

            string offsetMark = "+0x";
            int offSetStartPos = line.IndexOf(offsetMark);
            if (offSetStartPos == -1)
            {
                return -1;
            }

            int offsetEndPos = line.IndexOf(" ", offSetStartPos);
            if (offsetEndPos == -1)
            {
                return -1;
            }

            offSetStartPos += offsetMark.Length;
            string offset = line.Substring(offSetStartPos, offsetEndPos - offSetStartPos);
            pos = offsetEndPos + 1;

            try
            {
                return int.Parse(offset, System.Globalization.NumberStyles.HexNumber);
            }
            catch
            {
                return -1;
            }
        }

        private static void UnpackDisplayStructApp()
        {
            UnpackDisplayStructAppFromRes("SimpleDebugger.dll");
            UnpackDisplayStructAppFromRes("WindowsPE.dll");
            UnpackDisplayStructAppFromRes("DisplayStruct.exe");
        }

        private static void UnpackDisplayStructAppFromRes(string fileName)
        {
            if (File.Exists(fileName) == true)
            {
                return;
            }

            Type type = typeof(StructFieldInfo);

            using (Stream manifestResourceStream =
                type.Assembly.GetManifestResourceStream($@"{type.Namespace}.files.{fileName}"))
            {
                using (BinaryReader br = new BinaryReader(manifestResourceStream))
                {
                    byte[] buf = new byte[br.BaseStream.Length];
                    br.Read(buf, 0, buf.Length);

                    File.WriteAllBytes(fileName, buf);
                }
            }
        }
    }
}