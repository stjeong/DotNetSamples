using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using WindowsFormsApp1;

namespace ConsoleOutRedirect
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Console.WriteLine("Shown in Output only");
#if DEBUG
            using (var consoleOutRedirector = new ConsoleOutRedirector())
#endif
            {
                Console.WriteLine("Shown in AllocConsole window");

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new WindowsFormsApp1.Form1());
            }
        }
    }
}
