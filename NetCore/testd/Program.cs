using System;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;

/*
[Register]
sudo dotnet ./testd.dll --install

[Un-register]
sudo dotnet ./testd.dll --uninstall

[Start-service]
sudo systemctl start dotnet-testd

[Stop service: SIGINT]
sudo systemctl stop dotnet-testd

[Kill service: SIGTERM]
sudo systemctl kill dotnet-testd
*/

namespace testd
{
    // C# - .NET Core Console로 리눅스 daemon 프로그램 만드는 방법
    // http://www.sysnet.pe.kr/2/0/11958

    class Program
    {
        static void Main(string[] args)
        {
            EventWaitHandle ewh = new EventWaitHandle(false, EventResetMode.ManualReset);

            if (args.Length >= 1)
            {
                string netDllPath = typeof(Program).Assembly.Location;
                if (args[0] == "--install" || args[0] == "-i")
                {
                    InstallService(netDllPath, true);
                }
                else if (args[0] == "--uninstall" || args[0] == "-u")
                {
                    InstallService(netDllPath, false);
                }
                return;
            }

            // SIGTERM
            // systemctl kill
            AppDomain.CurrentDomain.ProcessExit += (s, e) =>
            {
                CleanupResources();
                WriteLog("Exited gracefully!");
            };

            // SIGINT
            // systemctl stop
            Console.CancelKeyPress += (s, e) =>
            {
                WriteLog("stopped");
                e.Cancel = true;
                ewh.Set();
            };

            ewh.WaitOne();
        }

        static void CleanupResources()
        {
        }

        static void WriteLog(string text)
        {
            Console.WriteLine(text);
            WriteSyslog(SyslogPriority.LOG_INFO, text);
        }

        static int InstallService(string netDllPath, bool doInstall)
        {
            string serviceFile = @"
[Unit]
Description={0} running on {1}

[Service]
WorkingDirectory={2}
ExecStart={3} {4}
KillSignal=SIGINT
SyslogIdentifier={5}

[Install]
WantedBy=multi-user.target
";

            string dllFileName = Path.GetFileName(netDllPath);
            string osName = Environment.OSVersion.ToString();

            FileInfo fi = null;

            try
            {
                fi = new FileInfo(netDllPath);
            }
            catch { }

            if (doInstall == true && fi != null && fi.Exists == false)
            {
                WriteLog("NOT FOUND: " + fi.FullName);
                return 1;
            }

            string serviceName = "dotnet-" + Path.GetFileNameWithoutExtension(dllFileName).ToLower();

            string exeName = Process.GetCurrentProcess().MainModule.FileName;

            string workingDir = Path.GetDirectoryName(fi.FullName);
            string fullText = string.Format(serviceFile, dllFileName, osName, workingDir,
                    exeName, fi.FullName, serviceName);

            string serviceFilePath = $"/etc/systemd/system/{serviceName}.service";

            if (doInstall == true)
            {
                File.WriteAllText(serviceFilePath, fullText);
                WriteLog(serviceFilePath + " Created");
                ControlService(serviceName, "enable");
                ControlService(serviceName, "start");
            }
            else
            {
                if (File.Exists(serviceFilePath) == true)
                {
                    ControlService(serviceName, "stop");
                    File.Delete(serviceFilePath);
                    WriteLog(serviceFilePath + " Deleted");
                }
            }

            return 0;
        }

        static int ControlService(string serviceName, string mode)
        {
            string servicePath = $"/etc/systemd/system/{serviceName}.service";

            if (File.Exists(servicePath) == false)
            {
                WriteLog($"No service: {serviceName} to {mode}");
                return 1;
            }

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = "systemctl";
            psi.Arguments = $"{mode} {serviceName}";
            Process child = Process.Start(psi);
            child.WaitForExit();
            return child.ExitCode;
        }

        [DllImport("libc", CharSet = CharSet.Ansi, ExactSpelling = true, CallingConvention = CallingConvention.Cdecl)]
        protected static extern void syslog(int priority, string fmt, byte[] msg);

        public enum SyslogPriority
        {
            LOG_EMERG = 0,
            LOG_ALERT = 1,
            LOG_CRIT = 2,
            LOG_ERR = 3,
            LOG_WARNING = 4,
            LOG_NOTICE = 5,
            LOG_INFO = 6,
            LOG_DEBUG = 7,
        }

        public static void WriteSyslog(SyslogPriority priority, string text)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux) == true)
            {
                syslog((int)priority, text, null);
            }
        }
    }
}
