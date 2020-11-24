using System;

namespace SampleHostApp
{
    class Program
    {
        static void Main(string[] args)
        {
#if NET5_0
            Console.WriteLine("[NET5] Hello World!");
#else
            Console.WriteLine("[NETCORE] Hello World!");
#endif
            Console.WriteLine("# of args: " + args?.Length);
            foreach (string arg in args ?? Array.Empty<string>())
            {
                Console.WriteLine(arg);
            }
        }
    }

    public delegate void MainDelegate(string[] args);
}
