using System;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Reflection.PortableExecutable;

namespace Microsoft.Metadata.Visualizer
{
    // .NET Reflection을 대체할 System.Reflection.Metadata 소개
    // http://www.sysnet.pe.kr/2/0/11930

    /*
    Metadata Tools
    ; https://github.com/dotnet/metadata-tools

    dotnet/metadata-tools
    ; https://github.com/dotnet/metadata-tools/blob/master/src/Microsoft.Metadata.Visualizer/MetadataVisualizer.cs
    */
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0 || new[] { "/?", "-?", "-h", "--help" }.Any(x => string.Equals(args[0], x, StringComparison.OrdinalIgnoreCase)))
            {
                PrintUsage();
                return;
            }

            foreach (var fileName in args)
            {
                Console.WriteLine(fileName);
                Console.WriteLine(new string('*', 80));

                try
                {
                    using (var stream = File.OpenRead(fileName))
                    using (var peFile = new PEReader(stream))
                    {
                        var metadataReader = peFile.GetMetadataReader();
                        var visualizer = new MetadataVisualizer(metadataReader, Console.Out);
                        visualizer.Visualize();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }

        private static void PrintUsage()
        {
            Console.WriteLine("This tool dumps the contents of all tables in a set of PE files.");
            Console.WriteLine("usage: mddumper <file>...");
        }
    }
}