using Microsoft.Diagnostics.Runtime.Interop;
using SimpleDebugger;
using System;
using System.Diagnostics;

namespace DbgShell
{
    // DbgShell/ClrMemDiag/Debugger/
    // https://github.com/microsoft/DbgShell/tree/master/ClrMemDiag/Debugger

    // Writing a Simple Debugger with DbgEng.Dll
    // http://blogs.microsoft.co.il/pavely/2015/07/27/writing-a-simple-debugger-with-dbgeng-dll/
    class Program
    {
        static void Main(string[] args)
        {
            using (UserDebugger debugger = new UserDebugger())
            {
                Console.CancelKeyPress += (s, e) =>
                {
                    e.Cancel = true;
                    debugger.SetInterrupt();
                };

                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = "DummyApp.exe";
                psi.UseShellExecute = true;

                Process child = Process.Start(psi);

                try
                {
                    if (debugger.AttachTo(child.Id) == false)
                    {
                        Console.WriteLine("Failed to attach");
                        return;
                    }

                    int hr = debugger.WaitForEvent();

                    while (true)
                    {
                        hr = debugger.GetExecutionStatus(out DEBUG_STATUS status);
                        if (hr != (int)HResult.S_OK)
                        {
                            break;
                        }

                        if (status == DEBUG_STATUS.NO_DEBUGGEE)
                        {
                            Console.WriteLine("No Target");
                            break;
                        }

                        if (status == DEBUG_STATUS.GO || status == DEBUG_STATUS.STEP_BRANCH ||
                              status == DEBUG_STATUS.STEP_INTO || status == DEBUG_STATUS.STEP_OVER)
                        {
                            hr = debugger.WaitForEvent();
                            continue;
                        }

                        if (debugger.StateChanged)
                        {
                            Console.WriteLine();
                            debugger.StateChanged = false;
                            if (debugger.BreakpointHit)
                            {
                                debugger.OutputCurrentState(DEBUG_OUTCTL.THIS_CLIENT, DEBUG_CURRENT.DEFAULT);
                                debugger.BreakpointHit = false;
                            }
                        }

                        debugger.OutputPromptWide(DEBUG_OUTCTL.THIS_CLIENT, null);
                        Console.Write(" ");
                        Console.ForegroundColor = ConsoleColor.Gray;
                        string command = Console.ReadLine();
                        debugger.ExecuteWide(command);
                    }
                }
                finally
                {
                    try
                    {
                        debugger.Detach();
                    } catch { }

                    try
                    {
                        child.Kill();
                    }
                    catch { }
                }
            }
        }
    }
}
