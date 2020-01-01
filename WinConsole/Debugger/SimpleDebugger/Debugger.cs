using Microsoft.Diagnostics.Runtime.Interop;
using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace SimpleDebugger
{
    // DbgShell/ClrMemDiag/Debugger/
    // https://github.com/microsoft/DbgShell/tree/master/ClrMemDiag/Debugger

    // Writing a Simple Debugger with DbgEng.Dll
    // http://blogs.microsoft.co.il/pavely/2015/07/27/writing-a-simple-debugger-with-dbgeng-dll/
    public class UserDebugger : IDebugOutputCallbacks, IDebugEventCallbacksWide, IDisposable
    {
        [DllImport("dbgeng.dll")]
        internal static extern int DebugCreate(ref Guid InterfaceId, [MarshalAs(UnmanagedType.IUnknown)] out object Interface);

        IDebugClient5 _client;
        IDebugControl4 _control;

        bool _outputText;
        public void SetOutputText(bool output)
        {
            _outputText = output;
        }

        public bool BreakpointHit { get; set; }
        public bool StateChanged { get; set; }

        public delegate void ModuleLoadedDelegate(ModuleInfo modInfo);
        public ModuleLoadedDelegate ModuleLoaded;

        public delegate void ExceptionOccurredDelegate(ExceptionInfo exInfo);
        public ExceptionOccurredDelegate ExceptionOccurred;

        public UserDebugger()
        {
            Guid guid = new Guid("27fe5639-8407-4f47-8364-ee118fb08ac8");
            object obj = null;

            int hr = DebugCreate(ref guid, out obj);

            if (hr < 0)
            {
                Console.WriteLine("SourceFix: Unable to acquire client interface");
                return;
            }

            _client = obj as IDebugClient5;
            _control = _client as IDebugControl4;
            _client.SetOutputCallbacks(this);
            _client.SetEventCallbacksWide(this);
        }

        public bool AttachTo(int pid)
        {
            int hr = _client.AttachProcess(0, (uint)pid, DEBUG_ATTACH.DEFAULT);
            return hr >= 0;
        }

        public int GetExecutionStatus(out DEBUG_STATUS status)
        {
            return _control.GetExecutionStatus(out status);
        }

        public void OutputCurrentState(DEBUG_OUTCTL outputControl, DEBUG_CURRENT flags)
        {
            _control.OutputCurrentState(outputControl, flags);
        }

        public void OutputPromptWide(DEBUG_OUTCTL outputControl, string format)
        {
            _control.OutputPromptWide(outputControl, format);
        }
            
        public void FlushCallbacks()
        {
            _client.FlushCallbacks();
        }

        public int Execute(string command)
        {
            return _control.Execute(DEBUG_OUTCTL.THIS_CLIENT, command, DEBUG_EXECUTE.DEFAULT);
        }

        public int Execute(DEBUG_OUTCTL outputControl, string command, DEBUG_EXECUTE flags)
        {
            return _control.Execute(outputControl, command, flags);
        }

        public int ExecuteWide(string command)
        {
            return _control.ExecuteWide(DEBUG_OUTCTL.THIS_CLIENT, command, DEBUG_EXECUTE.DEFAULT);
        }

        public int ExecuteWide(DEBUG_OUTCTL outputControl, string command, DEBUG_EXECUTE flags)
        {
            return _control.ExecuteWide(outputControl, command, flags);
        }

        public int WaitForEvent(DEBUG_WAIT flag = DEBUG_WAIT.DEFAULT, int timeout = Timeout.Infinite)
        {
            unchecked
            {
                return _control.WaitForEvent(flag, (uint)timeout);
            }
        }

        public void SetInterrupt(DEBUG_INTERRUPT flag = DEBUG_INTERRUPT.ACTIVE)
        {
            _control.SetInterrupt(flag);
        }

        public void Detach()
        {
            _client.DetachProcesses();
        }

        public void Dispose()
        {
            if (_control != null)
            {
                Marshal.ReleaseComObject(_control);
                _control = null;
            }

            if (_client != null)
            {
                Marshal.ReleaseComObject(_client);
                _client = null;
            }
        }

        public int Output([In] DEBUG_OUTPUT Mask, [In, MarshalAs(UnmanagedType.LPStr)] string Text)
        {
            if (_outputText == false)
            {
                return 0;
            }

            switch (Mask)
            {
                case DEBUG_OUTPUT.DEBUGGEE:
                    Console.ForegroundColor = ConsoleColor.Gray;
                    break;

                case DEBUG_OUTPUT.PROMPT:
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    break;

                case DEBUG_OUTPUT.ERROR:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;

                case DEBUG_OUTPUT.EXTENSION_WARNING:
                case DEBUG_OUTPUT.WARNING:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;

                case DEBUG_OUTPUT.SYMBOLS:
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    break;

                default:
                    Console.ForegroundColor = ConsoleColor.White;
                    break;
            }

            Console.Write(Text);
            return 0;
        }

        public int GetInterestMask([Out] out DEBUG_EVENT Mask)
        {
            Mask = DEBUG_EVENT.BREAKPOINT | DEBUG_EVENT.CHANGE_DEBUGGEE_STATE
                    | DEBUG_EVENT.CHANGE_ENGINE_STATE | DEBUG_EVENT.CHANGE_SYMBOL_STATE 
                    | DEBUG_EVENT.CREATE_PROCESS | DEBUG_EVENT.CREATE_THREAD | DEBUG_EVENT.EXCEPTION | DEBUG_EVENT.EXIT_PROCESS 
                    | DEBUG_EVENT.EXIT_THREAD | DEBUG_EVENT.LOAD_MODULE | DEBUG_EVENT.SESSION_STATUS | DEBUG_EVENT.SYSTEM_ERROR 
                    | DEBUG_EVENT.UNLOAD_MODULE;

            return 0;
        }
        public int Breakpoint([In, MarshalAs(UnmanagedType.Interface)] IDebugBreakpoint2 Bp)
        {
            BreakpointHit = true;
            StateChanged = true;
            return (int)DEBUG_STATUS.BREAK;
        }

        public int Breakpoint([In, MarshalAs(UnmanagedType.Interface)] IDebugBreakpoint Bp)
        {
            BreakpointHit = true;
            StateChanged = true;
            return (int)DEBUG_STATUS.BREAK;
        }

        public int Exception([In] ref EXCEPTION_RECORD64 Exception, [In] uint FirstChance)
        {
            if (ExceptionOccurred != null)
            {
                ExceptionInfo exInfo = new ExceptionInfo(Exception, FirstChance);
                ExceptionOccurred(exInfo);
            }

            return (int)DEBUG_STATUS.BREAK;
        }

        public int CreateThread([In] ulong Handle, [In] ulong DataOffset, [In] ulong StartOffset)
        {
            return 0;
        }

        public int ExitThread([In] uint ExitCode)
        {
            return 0;
        }

        public int CreateProcess([In] ulong ImageFileHandle, [In] ulong Handle, [In] ulong BaseOffset, [In] uint ModuleSize, 
            [In, MarshalAs(UnmanagedType.LPStr)] string ModuleName, [In, MarshalAs(UnmanagedType.LPStr)] string ImageName,
            [In] uint CheckSum, [In] uint TimeDateStamp, [In] ulong InitialThreadHandle, [In] ulong ThreadDataOffset, [In] ulong StartOffset)
        {
            //IDebugBreakpoint2 bp;

            //_control.AddBreakpoint2(DEBUG_BREAKPOINT_TYPE.CODE, uint.MaxValue, out bp);
            //bp.SetOffset(BaseOffset + StartOffset);
            //bp.SetFlags(DEBUG_BREAKPOINT_FLAG.ENABLED);
            //bp.SetCommandWide(".echo Stopping on process attach");

            return (int)DEBUG_STATUS.NO_CHANGE;
        }

        public int ExitProcess([In] uint ExitCode)
        {
            return 0;
        }

        public int LoadModule([In] ulong ImageFileHandle, [In] ulong BaseOffset, [In] uint ModuleSize, [In, MarshalAs(UnmanagedType.LPStr)] string ModuleName,
            [In, MarshalAs(UnmanagedType.LPStr)] string ImageName, [In] uint CheckSum, [In] uint TimeDateStamp)
        {
            if (ModuleLoaded != null)
            {
                ModuleInfo modInfo = new ModuleInfo(ImageFileHandle, BaseOffset, ModuleSize, ModuleName, ImageName, CheckSum, TimeDateStamp);
                ModuleLoaded(modInfo);
            }

            return 0;
        }

        public int UnloadModule([In, MarshalAs(UnmanagedType.LPStr)] string ImageBaseName, [In] ulong BaseOffset)
        {
            return 0;
        }

        public int SystemError([In] uint Error, [In] uint Level)
        {
            return 0;
        }

        public int SessionStatus([In] DEBUG_SESSION Status)
        {
            return 0;
        }

        public int ChangeDebuggeeState([In] DEBUG_CDS Flags, [In] ulong Argument)
        {
            return 0;
        }

        public int ChangeEngineState([In] DEBUG_CES Flags, [In] ulong Argument)
        {
            return 0;
        }

        public int ChangeSymbolState([In] DEBUG_CSS Flags, [In] ulong Argument)
        {
            return 0;
        }
    }
}
