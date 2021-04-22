using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

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

tail -F /var/log/syslog
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
            // systemctl stop
            AppDomain.CurrentDomain.ProcessExit += (s, e) =>
            {
                CleanupResources();
                WriteLog("Exited gracefully!");
            };

            // Interactive: Ctrl+C
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
        }

        static int InstallService(string netDllPath, bool doInstall)
        {
            // 2021-04-22
            // remove: KillSignal=SIGINT
            // added: KillMode=mixed
            string serviceFile = @"
[Unit]
Description={0} running on {1}

[Service]
WorkingDirectory={2}
ExecStart={3} {4}
SyslogIdentifier={5}
KillMode=mixed

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
    }
}
