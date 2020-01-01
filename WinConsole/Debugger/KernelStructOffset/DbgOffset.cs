using KernelStructOffset;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

public class DbgOffset
{
    public static Dictionary<string, int> Get(string typeName)
    {
        return Get(typeName, "ntdll.dll");
    }

    public static Dictionary<string, int> Get(string typeName, string moduleName)
    {
        List<StructFieldInfo> list = GetList(typeName, moduleName);
        Dictionary<string, int> dict = new Dictionary<string, int>();

        foreach (StructFieldInfo item in list)
        {
            dict.Add(item.Name, item.Offset);
        }

        return dict;
    }

    private static List<StructFieldInfo> GetList(string typeName, string moduleName)
    {
        UnpackDisplayStructApp();

        ProcessStartInfo psi = new ProcessStartInfo();
        psi.FileName = "DisplayStruct.exe";
        psi.UseShellExecute = false;
        psi.Arguments = $"{typeName} {moduleName}";
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
