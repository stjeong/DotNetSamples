
#define USE_HOSTFXR_MAIN1
#define THROW_EX1

#define NETHOST_NORETURN __declspec(noreturn)

#define NETHOST_SUCCESS 0

struct get_hostfxr_parameters {
    size_t size;
    const wchar_t* assembly_path;
    const wchar_t* dotnet_root;
};

extern "C" int __stdcall get_hostfxr_path(
    wchar_t* buffer,
    size_t * buffer_size,
    const struct get_hostfxr_parameters* parameters);

#include <iostream>
#include <Windows.h>
#include <shlwapi.h>
#include <assert.h>
#include <pathcch.h>

#include <fstream>
#include <string>

#include "resource.h"

#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "pathcch.lib")

#define NETHOST_CALLTYPE __stdcall
#define CORECLR_DELEGATE_CALLTYPE __stdcall

enum failure_type
{
    failure_load_runtime = 1,
    failure_load_export,
};

enum hostfxr_delegate_type
{
    hdt_com_activation,
    hdt_load_in_memory_assembly,
    hdt_winrt_activation,
    hdt_com_register,
    hdt_com_unregister,
    hdt_load_assembly_and_get_function_pointer
};

typedef struct hostfxr_initialize_parameters
{
    size_t size;
    const wchar_t* host_path;
    const wchar_t* dotnet_root;
} hostfxr_initialize_parameters;

typedef void* hostfxr_handle;

typedef void (NETHOST_CALLTYPE* failure_fn)(enum failure_type type, int error_code);
typedef int32_t(NETHOST_CALLTYPE* hostfxr_initialize_for_runtime_config_fn)(
    const wchar_t* runtime_config_path,
    const hostfxr_initialize_parameters* parameters,
    /*out*/ hostfxr_handle* host_context_handle);
typedef int32_t(NETHOST_CALLTYPE* hostfxr_get_runtime_delegate_fn)(
    const hostfxr_handle host_context_handle,
    enum hostfxr_delegate_type type,
    /*out*/ void** delegate);
typedef int32_t(NETHOST_CALLTYPE* hostfxr_close_fn)(const hostfxr_handle host_context_handle);
typedef int (CORECLR_DELEGATE_CALLTYPE* load_assembly_and_get_function_pointer_fn)(
    const wchar_t* assembly_path      /* Fully qualified path to assembly */,
    const wchar_t* type_name          /* Assembly qualified type name */,
    const wchar_t* method_name        /* Public static method name compatible with delegateType */,
    const wchar_t* delegate_type_name /* Assembly qualified delegate type name or null
                                        or UNMANAGEDCALLERSONLY_METHOD if the method is marked with
                                        the UnmanagedCallersOnlyAttribute. */,
    void* reserved           /* Extensibility parameter (currently unused and must be 0) */,
    /*out*/ void** delegate          /* Pointer where to store the function pointer result */);

typedef int(*hostfxr_run_app_fn)(void* host_context_handle);
typedef INT(*hostfxr_main_fn) (DWORD argc, CONST PCWSTR argv[]);
typedef int32_t(*hostfxr_initialize_for_dotnet_command_line_fn)(
    int argc,
    const wchar_t** argv,
    const struct hostfxr_initialize_parameters* parameters,
    /*out*/ hostfxr_handle* host_context_handle);

static failure_fn failure_fptr;
static hostfxr_initialize_for_runtime_config_fn init_fptr;
static hostfxr_get_runtime_delegate_fn get_delegate_fptr;
static hostfxr_close_fn close_fptr;
static hostfxr_initialize_for_dotnet_command_line_fn hostfxr_initialize_for_dotnet_command_line_fptr;

static load_assembly_and_get_function_pointer_fn get_managed_export_fptr;

static hostfxr_run_app_fn hostfxr_run_app_fptr;
static hostfxr_main_fn hostfxr_main_fptr;

#pragma comment(lib, "C:\\Program Files\\dotnet\\packs\\Microsoft.NETCore.App.Host.win-x64\\5.0.0\\runtimes\\win-x64\\native\\libnethost.lib")

static bool is_failure(int rc)
{
    // The CLR hosting API uses the Win32 HRESULT scheme. This means
    // the high order bit indicates an error and S_FALSE (1) can be returned
    // and is _not_ a failure.
    return (rc < NETHOST_SUCCESS);
}

NETHOST_NORETURN static void noreturn_runtime_load_failure(int error_code)
{
    if (failure_fptr)
        failure_fptr(failure_load_runtime, error_code);

    // Nothing to do if the runtime failed to load.
    abort();
}

static void* load_library(const wchar_t* path)
{
    assert(path != NULL);
    HMODULE h = LoadLibraryW(path);
    assert(h != NULL);
    return (void*)h;
}

static void* get_export(void* h, const char* name)
{
    assert(h != NULL && name != NULL);
    void* f = GetProcAddress((HMODULE)h, name);
    assert(f != NULL);
    return f;
}

void write_and_get_config_path(wchar_t* config_path, size_t config_path_size, wchar_t* hostfxr_path)
{
    // TODO: It depends on hostfxr_path
    std::string txt = 
"{\n"
"    \"runtimeOptions\": {\n"
"        \"tfm\": \"net5.0\",\n"
"            \"framework\" : {\n"
"            \"name\": \"Microsoft.NETCore.App\",\n"
"            \"version\" : \"5.0.0\"\n"
"        }\n"
"    }\n"
"}\n";
    
    HMODULE hModule = ::GetModuleHandle(NULL);

    ::GetModuleFileName(hModule, config_path, (DWORD)config_path_size);
    ::PathCchRemoveFileSpec(config_path, config_path_size);
    ::PathCchAppendEx(config_path, config_path_size, L"SampleHostApp.runtimeconfig.json", 0);

    std::ofstream ofs;
    ofs.open(config_path);

    ofs << txt;

    ofs.close();
}

void write_and_get_assembly_path(wchar_t* assembly_path, size_t asm_path_size)
{
    HMODULE hModule = ::GetModuleHandle(NULL);

    ::GetModuleFileName(hModule, assembly_path, (DWORD)asm_path_size);
    ::PathCchRemoveFileSpec(assembly_path, asm_path_size);
    ::PathCchAppendEx(assembly_path, asm_path_size, L"SampleHostApp.dll", 0);

    HRSRC hResource = FindResource(hModule, MAKEINTRESOURCE(IDR_DLLFILE1), L"DLLFILE");
    HGLOBAL hMemory = LoadResource(hModule, hResource);
    DWORD dwSize = SizeofResource(hModule, hResource);
    LPVOID lpAddress = LockResource(hMemory);

    std::ofstream ofs;
    ofs.open(assembly_path, std::ios::binary);

    ofs.write((char*)lpAddress, dwSize);

    ofs.close();
}

void throw_failure(int rc, hostfxr_handle cxt)
{
    if (rc != 0)
    {
        if (cxt != nullptr)
        {
            close_fptr(cxt);
        }

        noreturn_runtime_load_failure(rc);
    }
}

int main()
{
    wchar_t hostfxr_path[MAX_PATH];

    {
        size_t buffer_size = sizeof(hostfxr_path) / sizeof(wchar_t);

        int rc = get_hostfxr_path(hostfxr_path, &buffer_size, nullptr);
        // buffer == C:\Program Files\dotnet\host\fxr\5.0.0\hostfxr.dll

        if (is_failure(rc))
            noreturn_runtime_load_failure(rc);

        void* lib = load_library(hostfxr_path);
        init_fptr = (hostfxr_initialize_for_runtime_config_fn)get_export(lib, "hostfxr_initialize_for_runtime_config");
        get_delegate_fptr = (hostfxr_get_runtime_delegate_fn)get_export(lib, "hostfxr_get_runtime_delegate");
        close_fptr = (hostfxr_close_fn)get_export(lib, "hostfxr_close");

        hostfxr_run_app_fptr = (hostfxr_run_app_fn)get_export(lib, "hostfxr_run_app");
        hostfxr_main_fptr = (hostfxr_main_fn)get_export(lib, "hostfxr_main");
        hostfxr_initialize_for_dotnet_command_line_fptr = (hostfxr_initialize_for_dotnet_command_line_fn)get_export(lib, "hostfxr_initialize_for_dotnet_command_line");
    }

    wchar_t config_path[MAX_PATH];

    {
        size_t config_path_size = sizeof(config_path) / sizeof(wchar_t);
        write_and_get_config_path(config_path, config_path_size, hostfxr_path);
    }

    // config_path == "...path_to_runtimeconfig.json..."
    /*
{
    "runtimeOptions": {
        "tfm": "net5.0",
        "framework": {
            "name": "Microsoft.NETCore.App",
            "version": "5.0.0"
        }
    }
}
    */

    wchar_t assembly_path[MAX_PATH];

    {
        size_t asm_path_size = sizeof(assembly_path) / sizeof(wchar_t);
        write_and_get_assembly_path(assembly_path, asm_path_size);
    }


    hostfxr_handle cxt = nullptr;
    int rc = 0;

    do
    {
#if defined(USE_HOSTFXR_MAIN)

        int nArgs = 0;
        LPWSTR* szArglist = CommandLineToArgvW(GetCommandLineW(), &nArgs);

        PCWSTR* argv = new PCWSTR[2 + ((short)nArgs) - 1];
        argv[0] = config_path;
        argv[1] = assembly_path;

        for (int i = 0; i < nArgs - 1; i++)
        {
            argv[2 + i] = szArglist[i + 1];
        }

        hostfxr_main_fptr(2 + (nArgs - 1), argv);

        delete[] argv;
        LocalFree(szArglist);

#else
        rc = init_fptr(config_path, nullptr, &cxt);
        if (is_failure(rc) || cxt == nullptr)
        {
#ifdef THROW_EX
            throw_failure(rc, cxt);
#endif
            break;
        }

        void* load_assembly_and_get_function_pointer = nullptr;

        rc = get_delegate_fptr(
            cxt,
            hdt_load_assembly_and_get_function_pointer,
            &load_assembly_and_get_function_pointer);
        if (is_failure(rc) || load_assembly_and_get_function_pointer == nullptr)
        {
#ifdef THROW_EX
            throw_failure(rc, cxt);
#endif
            break;
        }

        get_managed_export_fptr = (load_assembly_and_get_function_pointer_fn)load_assembly_and_get_function_pointer;

        
        const wchar_t* dotnet_type = L"SampleHostApp.Program, SampleHostApp";
        const wchar_t* dotnet_type_method = L"Main";
        // const wchar_t* dotnet_delegate_type = L"System.Action<System.String[]>, System.Runtime";
        const wchar_t* dotnet_delegate_type = L"SampleHostApp.MainDelegate, SampleHostApp";

        void* func = nullptr;
        rc = get_managed_export_fptr(
            assembly_path,
            dotnet_type,
            dotnet_type_method,
            dotnet_delegate_type,
            nullptr,
            &func);
        if (is_failure(rc) || func == nullptr)
        {
#ifdef THROW_EX
            throw_failure(rc, cxt);
#endif
            break;
        }

        typedef void (*main_ptr)(intptr_t);
        main_ptr mainFunc = (main_ptr)func;

        mainFunc(0);
#endif

    } while (false);

    if (cxt != nullptr)
    {
        close_fptr(cxt);
    }

    return 0;
}
