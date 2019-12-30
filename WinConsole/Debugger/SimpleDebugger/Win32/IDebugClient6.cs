﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Text;
using System.Runtime.InteropServices;

#pragma warning disable 1591

namespace Microsoft.Diagnostics.Runtime.Interop
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("e3acb9d7-7ec2-4f0c-a0da-e81e0cbbe628")]
    public interface IDebugClient6 : IDebugClient5
    {
        /* IDebugClient */

        [PreserveSig]
        new int AttachKernel(
            [In] DEBUG_ATTACH Flags,
            [In, MarshalAs(UnmanagedType.LPStr)] string ConnectOptions);

        [PreserveSig]
        new int GetKernelConnectionOptions(
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 OptionsSize);

        [PreserveSig]
        new int SetKernelConnectionOptions(
            [In, MarshalAs(UnmanagedType.LPStr)] string Options);

        [PreserveSig]
        new int StartProcessServer(
            [In] DEBUG_CLASS Flags,
            [In, MarshalAs(UnmanagedType.LPStr)] string Options,
            [In] IntPtr Reserved);

        [PreserveSig]
        new int ConnectProcessServer(
            [In, MarshalAs(UnmanagedType.LPStr)] string RemoteOptions,
            [Out] out UInt64 Server);

        [PreserveSig]
        new int DisconnectProcessServer(
            [In] UInt64 Server);

        [PreserveSig]
        new int GetRunningProcessSystemIds(
            [In] UInt64 Server,
            [Out, MarshalAs(UnmanagedType.LPArray)] UInt32[] Ids,
            [In] UInt32 Count,
            [Out] out UInt32 ActualCount);

        [PreserveSig]
        new int GetRunningProcessSystemIdByExecutableName(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPStr)] string ExeName,
            [In] DEBUG_GET_PROC Flags,
            [Out] out UInt32 Id);

        [PreserveSig]
        new int GetRunningProcessDescription(
            [In] UInt64 Server,
            [In] UInt32 SystemId,
            [In] DEBUG_PROC_DESC Flags,
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder ExeName,
            [In] Int32 ExeNameSize,
            [Out] out UInt32 ActualExeNameSize,
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Description,
            [In] Int32 DescriptionSize,
            [Out] out UInt32 ActualDescriptionSize);

        [PreserveSig]
        new int AttachProcess(
            [In] UInt64 Server,
            [In] UInt32 ProcessID,
            [In] DEBUG_ATTACH AttachFlags);

        [PreserveSig]
        new int CreateProcess(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPStr)] string CommandLine,
            [In] DEBUG_CREATE_PROCESS Flags);

        [PreserveSig]
        new int CreateProcessAndAttach(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPStr)] string CommandLine,
            [In] DEBUG_CREATE_PROCESS Flags,
            [In] UInt32 ProcessId,
            [In] DEBUG_ATTACH AttachFlags);

        [PreserveSig]
        new int GetProcessOptions(
            [Out] out DEBUG_PROCESS Options);

        [PreserveSig]
        new int AddProcessOptions(
            [In] DEBUG_PROCESS Options);

        [PreserveSig]
        new int RemoveProcessOptions(
            [In] DEBUG_PROCESS Options);

        [PreserveSig]
        new int SetProcessOptions(
            [In] DEBUG_PROCESS Options);

        [PreserveSig]
        new int OpenDumpFile(
            [In, MarshalAs(UnmanagedType.LPStr)] string DumpFile);

        [PreserveSig]
        new int WriteDumpFile(
            [In, MarshalAs(UnmanagedType.LPStr)] string DumpFile,
            [In] DEBUG_DUMP Qualifier);

        [PreserveSig]
        new int ConnectSession(
            [In] DEBUG_CONNECT_SESSION Flags,
            [In] UInt32 HistoryLimit);

        [PreserveSig]
        new int StartServer(
            [In, MarshalAs(UnmanagedType.LPStr)] string Options);

        [PreserveSig]
        new int OutputServer(
            [In] DEBUG_OUTCTL OutputControl,
            [In, MarshalAs(UnmanagedType.LPStr)] string Machine,
            [In] DEBUG_SERVERS Flags);

        [PreserveSig]
        new int TerminateProcesses();

        [PreserveSig]
        new int DetachProcesses();

        [PreserveSig]
        new int EndSession(
            [In] DEBUG_END Flags);

        [PreserveSig]
        new int GetExitCode(
            [Out] out UInt32 Code);

        [PreserveSig]
        new int DispatchCallbacks(
            [In] UInt32 Timeout);

        [PreserveSig]
        new int ExitDispatch(
            [In, MarshalAs(UnmanagedType.Interface)] IDebugClient Client);

        [PreserveSig]
        new int CreateClient(
            [Out, MarshalAs(UnmanagedType.Interface)] out IDebugClient Client);

        [PreserveSig]
        new int GetInputCallbacks(
            [Out, MarshalAs(UnmanagedType.Interface)] out IDebugInputCallbacks Callbacks);

        [PreserveSig]
        new int SetInputCallbacks(
            [In, MarshalAs(UnmanagedType.Interface)] IDebugInputCallbacks Callbacks);

        /* GetOutputCallbacks could a conversion thunk from the debugger engine so we can't specify a specific interface */

        [PreserveSig]
        new int GetOutputCallbacks(
            [Out] out IDebugOutputCallbacks Callbacks);

        /* We may have to pass a debugger engine conversion thunk back in so we can't specify a specific interface */

        [PreserveSig]
        new int SetOutputCallbacks(
            [In] IDebugOutputCallbacks Callbacks);

        [PreserveSig]
        new int GetOutputMask(
            [Out] out DEBUG_OUTPUT Mask);

        [PreserveSig]
        new int SetOutputMask(
            [In] DEBUG_OUTPUT Mask);

        [PreserveSig]
        new int GetOtherOutputMask(
            [In, MarshalAs(UnmanagedType.Interface)] IDebugClient Client,
            [Out] out DEBUG_OUTPUT Mask);

        [PreserveSig]
        new int SetOtherOutputMask(
            [In, MarshalAs(UnmanagedType.Interface)] IDebugClient Client,
            [In] DEBUG_OUTPUT Mask);

        [PreserveSig]
        new int GetOutputWidth(
            [Out] out UInt32 Columns);

        [PreserveSig]
        new int SetOutputWidth(
            [In] UInt32 Columns);

        [PreserveSig]
        new int GetOutputLinePrefix(
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 PrefixSize);

        [PreserveSig]
        new int SetOutputLinePrefix(
            [In, MarshalAs(UnmanagedType.LPStr)] string Prefix);

        [PreserveSig]
        new int GetIdentity(
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 IdentitySize);

        [PreserveSig]
        new int OutputIdentity(
            [In] DEBUG_OUTCTL OutputControl,
            [In] UInt32 Flags,
            [In, MarshalAs(UnmanagedType.LPStr)] string Format);

        /* GetEventCallbacks could a conversion thunk from the debugger engine so we can't specify a specific interface */

        [PreserveSig]
        new int GetEventCallbacks(
            [Out] out IDebugEventCallbacks Callbacks);

        /* We may have to pass a debugger engine conversion thunk back in so we can't specify a specific interface */

        [PreserveSig]
        new int SetEventCallbacks(
            [In] IDebugEventCallbacks Callbacks);

        [PreserveSig]
        new int FlushCallbacks();

        /* IDebugClient2 */

        [PreserveSig]
        new int WriteDumpFile2(
            [In, MarshalAs(UnmanagedType.LPStr)] string DumpFile,
            [In] DEBUG_DUMP Qualifier,
            [In] DEBUG_FORMAT FormatFlags,
            [In, MarshalAs(UnmanagedType.LPStr)] string Comment);

        [PreserveSig]
        new int AddDumpInformationFile(
            [In, MarshalAs(UnmanagedType.LPStr)] string InfoFile,
            [In] DEBUG_DUMP_FILE Type);

        [PreserveSig]
        new int EndProcessServer(
            [In] UInt64 Server);

        [PreserveSig]
        new int WaitForProcessServerEnd(
            [In] UInt32 Timeout);

        [PreserveSig]
        new int IsKernelDebuggerEnabled();

        [PreserveSig]
        new int TerminateCurrentProcess();

        [PreserveSig]
        new int DetachCurrentProcess();

        [PreserveSig]
        new int AbandonCurrentProcess();

        /* IDebugClient3 */

        [PreserveSig]
        new int GetRunningProcessSystemIdByExecutableNameWide(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPWStr)] string ExeName,
            [In] DEBUG_GET_PROC Flags,
            [Out] out UInt32 Id);

        [PreserveSig]
        new int GetRunningProcessDescriptionWide(
            [In] UInt64 Server,
            [In] UInt32 SystemId,
            [In] DEBUG_PROC_DESC Flags,
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder ExeName,
            [In] Int32 ExeNameSize,
            [Out] out UInt32 ActualExeNameSize,
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Description,
            [In] Int32 DescriptionSize,
            [Out] out UInt32 ActualDescriptionSize);

        [PreserveSig]
        new int CreateProcessWide(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPWStr)] string CommandLine,
            [In] DEBUG_CREATE_PROCESS CreateFlags);

        [PreserveSig]
        new int CreateProcessAndAttachWide(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPWStr)] string CommandLine,
            [In] DEBUG_CREATE_PROCESS CreateFlags,
            [In] UInt32 ProcessId,
            [In] DEBUG_ATTACH AttachFlags);

        /* IDebugClient4 */

        [PreserveSig]
        new int OpenDumpFileWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string FileName,
            [In] UInt64 FileHandle);

        [PreserveSig]
        new int WriteDumpFileWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string DumpFile,
            [In] UInt64 FileHandle,
            [In] DEBUG_DUMP Qualifier,
            [In] DEBUG_FORMAT FormatFlags,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Comment);

        [PreserveSig]
        new int AddDumpInformationFileWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string FileName,
            [In] UInt64 FileHandle,
            [In] DEBUG_DUMP_FILE Type);

        [PreserveSig]
        new int GetNumberDumpFiles(
            [Out] out UInt32 Number);

        [PreserveSig]
        new int GetDumpFile(
            [In] UInt32 Index,
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 NameSize,
            [Out] out UInt64 Handle,
            [Out] out UInt32 Type);

        [PreserveSig]
        new int GetDumpFileWide(
            [In] UInt32 Index,
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 NameSize,
            [Out] out UInt64 Handle,
            [Out] out UInt32 Type);

        /* IDebugClient5 */

        [PreserveSig]
        new int AttachKernelWide(
            [In] DEBUG_ATTACH Flags,
            [In, MarshalAs(UnmanagedType.LPWStr)] string ConnectOptions);

        [PreserveSig]
        new int GetKernelConnectionOptionsWide(
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 OptionsSize);

        [PreserveSig]
        new int SetKernelConnectionOptionsWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string Options);

        [PreserveSig]
        new int StartProcessServerWide(
            [In] DEBUG_CLASS Flags,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Options,
            [In] IntPtr Reserved);

        [PreserveSig]
        new int ConnectProcessServerWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string RemoteOptions,
            [Out] out UInt64 Server);

        [PreserveSig]
        new int StartServerWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string Options);

        [PreserveSig]
        new int OutputServersWide(
            [In] DEBUG_OUTCTL OutputControl,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Machine,
            [In] DEBUG_SERVERS Flags);

        /* GetOutputCallbacks could a conversion thunk from the debugger engine so we can't specify a specific interface */

        [PreserveSig]
        new int GetOutputCallbacksWide(
            [Out] out IDebugOutputCallbacksWide Callbacks);

        /* We may have to pass a debugger engine conversion thunk back in so we can't specify a specific interface */

        [PreserveSig]
        new int SetOutputCallbacksWide(
            [In] IDebugOutputCallbacks2 Callbacks);

        [PreserveSig]
        new int GetOutputLinePrefixWide(
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 PrefixSize);

        [PreserveSig]
        new int SetOutputLinePrefixWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string Prefix);

        [PreserveSig]
        new int GetIdentityWide(
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 IdentitySize);

        [PreserveSig]
        new int OutputIdentityWide(
            [In] DEBUG_OUTCTL OutputControl,
            [In] UInt32 Flags,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Machine);

        /* GetEventCallbacks could a conversion thunk from the debugger engine so we can't specify a specific interface */

        [PreserveSig]
        new int GetEventCallbacksWide(
            [Out] out IDebugEventCallbacksWide Callbacks);

        /* We may have to pass a debugger engine conversion thunk back in so we can't specify a specific interface */

        [PreserveSig]
        new int SetEventCallbacksWide(
            [In] IDebugEventCallbacksWide Callbacks);

        [PreserveSig]
        new int CreateProcess2(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPStr)] string CommandLine,
            [In] ref DEBUG_CREATE_PROCESS_OPTIONS OptionsBuffer,
            [In] UInt32 OptionsBufferSize,
            [In, MarshalAs(UnmanagedType.LPStr)] string InitialDirectory,
            [In, MarshalAs(UnmanagedType.LPStr)] string Environment);

        [PreserveSig]
        new int CreateProcess2Wide(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPWStr)] string CommandLine,
            [In] ref DEBUG_CREATE_PROCESS_OPTIONS OptionsBuffer,
            [In] UInt32 OptionsBufferSize,
            [In, MarshalAs(UnmanagedType.LPWStr)] string InitialDirectory,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Environment);

        [PreserveSig]
        new int CreateProcessAndAttach2(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPStr)] string CommandLine,
            [In] ref DEBUG_CREATE_PROCESS_OPTIONS OptionsBuffer,
            [In] UInt32 OptionsBufferSize,
            [In, MarshalAs(UnmanagedType.LPStr)] string InitialDirectory,
            [In, MarshalAs(UnmanagedType.LPStr)] string Environment,
            [In] UInt32 ProcessId,
            [In] DEBUG_ATTACH AttachFlags);

        [PreserveSig]
        new int CreateProcessAndAttach2Wide(
            [In] UInt64 Server,
            [In, MarshalAs(UnmanagedType.LPWStr)] string CommandLine,
            [In] ref DEBUG_CREATE_PROCESS_OPTIONS OptionsBuffer,
            [In] UInt32 OptionsBufferSize,
            [In, MarshalAs(UnmanagedType.LPWStr)] string InitialDirectory,
            [In, MarshalAs(UnmanagedType.LPWStr)] string Environment,
            [In] UInt32 ProcessId,
            [In] DEBUG_ATTACH AttachFlags);

        [PreserveSig]
        new int PushOutputLinePrefix(
            [In, MarshalAs(UnmanagedType.LPStr)] string NewPrefix,
            [Out] out UInt64 Handle);

        [PreserveSig]
        new int PushOutputLinePrefixWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string NewPrefix,
            [Out] out UInt64 Handle);

        [PreserveSig]
        new int PopOutputLinePrefix(
            [In] UInt64 Handle);

        [PreserveSig]
        new int GetNumberInputCallbacks(
            [Out] out UInt32 Count);

        [PreserveSig]
        new int GetNumberOutputCallbacks(
            [Out] out UInt32 Count);

        [PreserveSig]
        new int GetNumberEventCallbacks(
            [In] DEBUG_EVENT Flags,
            [Out] out UInt32 Count);

        [PreserveSig]
        new int GetQuitLockString(
            [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 StringSize);

        [PreserveSig]
        new int SetQuitLockString(
            [In, MarshalAs(UnmanagedType.LPStr)] string LockString);

        [PreserveSig]
        new int GetQuitLockStringWide(
            [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder Buffer,
            [In] Int32 BufferSize,
            [Out] out UInt32 StringSize);

        [PreserveSig]
        new int SetQuitLockStringWide(
            [In, MarshalAs(UnmanagedType.LPWStr)] string LockString);

        /* IDebugClient6 */

        [PreserveSig]
        int SetEventContextCallbacks(
            [In] IDebugEventContextCallbacks Callbacks);
    }
}
