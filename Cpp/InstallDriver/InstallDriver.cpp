
#include "stdafx.h"

int InstallDriver(wchar_t *pDriverFilePath, wchar_t *pDriverName);
int UninstallDriver(wchar_t *pDriverName);
int StartDeviceDriver(wchar_t *pDriverName);
int StopDeviceDriver(wchar_t *pDriverName);

// HKLM\SYSTEM\CurrentControlSet\Services\[DriverName]

// [install]
//  InstallDriver 1 "[경로]" "[DriverName]"
//  or
//  sc create "[DriverName]" binPath= "[경로]" type= kernel start= demand
//
// [uninstall]
//  InstallDriver 0 "[DriverName]"
//
// [start_service]
//  InstallDriver 2 "[DriverName]"
//  or
//  net start "[DriverName]"
//
// [stop_service]
//  InstallDriver 3 "[DriverName]"
//  or
//  net stop "[DriverName]"

int _tmain(int argc, _TCHAR* argv[])
{
    wchar_t *mode = nullptr;
    wchar_t driverFullPath[MAX_PATH];
    wchar_t *driverName = nullptr;

    if (argc > 0)
    {
        mode = argv[1];
    }

    if (mode == nullptr)
    {
        return 1;
    }

    if (argc == 4)
    {
        wchar_t currentPath[MAX_PATH];
        ::GetCurrentDirectory(MAX_PATH, currentPath);

        ::PathCombine(driverFullPath, currentPath, argv[2]);

        driverName = argv[3];
    } 
    else if (argc == 3)
    {
        driverName = argv[2];
    }

    if (driverName == nullptr)
    {
        return 1;
    }

    if (wcscmp(mode, L"1") == 0)
    {
        return InstallDriver(driverFullPath, driverName) == 0;
    }
    else if (wcscmp(mode, L"0") == 0)
    {
        return UninstallDriver(driverName) == 0;
    }
    else if (wcscmp(mode, L"2") == 0)
    {
        return StartDeviceDriver(driverName) == 0;
    }
    else if (wcscmp(mode, L"3") == 0)
    {
        return StopDeviceDriver(driverName) == 0;
    }

    return 0;
}

int StartDeviceDriver(wchar_t *pDriverName)
{
    int result = 0;

    SC_HANDLE hSCManager = NULL;
    SC_HANDLE hService = NULL;

    do
    {
        hSCManager = OpenSCManager(NULL, NULL, SC_MANAGER_CONNECT);
        if (hSCManager == NULL)
        {
            wprintf(L"[StartDeviceDriver] hSCManager == NULL");
            break;
        }

        hService = OpenService(hSCManager, pDriverName, SERVICE_START);
        if (hService == NULL)
        {
            wprintf(L"[StartDeviceDriver] hService == NULL");
            break;
        }

        if (::StartService(hService, 0, NULL) == TRUE)
        {
            result = 1;
        }
    } while (false);

    if (hService != NULL)
    {
        CloseServiceHandle(hService);
        hService = NULL;
    }

    if (hSCManager != NULL)
    {
        CloseServiceHandle(hSCManager);
        hSCManager = NULL;
    }

    return result;
}

int StopDeviceDriver(wchar_t *pDriverName)
{
    int result = 0;

    SC_HANDLE hSCManager = NULL;
    SC_HANDLE hService = NULL;
    DWORD dwResult;

    do
    {
        hSCManager = OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
        if (hSCManager == NULL)
        {
            wprintf(L"[StopDeviceDriver] hSCManager == NULL");
            break;
        }

        hService = OpenService(hSCManager, pDriverName, SERVICE_STOP | SERVICE_QUERY_STATUS);
        if (hService == NULL)
        {
            wprintf(L"[StopDeviceDriver] hService == NULL");
            break;
        }

        SERVICE_STATUS st;
        if (::ControlService(hService, SERVICE_CONTROL_STOP, &st) == TRUE)
        {
            result = 1;
        }
        else
        {
            dwResult = ::GetLastError();
            wprintf(L"[StopDeviceDriver - %s] ControlService == FALSE, LastError = %d", pDriverName, dwResult);
        }
    } while (false);

    if (hService != NULL)
    {
        CloseServiceHandle(hService);
        hService = NULL;
    }

    if (hSCManager != NULL)
    {
        CloseServiceHandle(hSCManager);
        hSCManager = NULL;
    }

    return result;
}

int InstallDriver(wchar_t *pDriverFilePath, wchar_t *pDriverName)
{
    int result = 0;

    SC_HANDLE hSCManager = NULL;
    SC_HANDLE hService = NULL;

    do
    {
        hSCManager = OpenSCManager(NULL, NULL, SC_MANAGER_CREATE_SERVICE);
        if (hSCManager == NULL)
        {
            wprintf(L"[InstallDriver] hSCManager == NULL");
            break;
        }
         
        /* HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\[DirverName] */
        hService = CreateService(hSCManager, pDriverName, pDriverName,        
                                    GENERIC_READ, 
                                    SERVICE_KERNEL_DRIVER,             /* service type */
                                    SERVICE_DEMAND_START,              /* start type */
                                    SERVICE_ERROR_NORMAL,              /* error control type */
                                    pDriverFilePath, /* service's binary */
                                    NULL,                              /* no load ordering group */
                                    NULL,                              /* no tag identifier*/
                                    NULL,                              /* no dependencies */
                                    NULL,                              /* LocalSystem account*/
                                    NULL                               /* no password */
                                    );		

        if (hService == NULL) 
        {
            DWORD dwResult = ::GetLastError();
            wprintf(L"[InstallDriver] hService == NULL, LastError = %d", dwResult);
            break;
        }

        result = 1;

    } while (false);

    if (hService != NULL)
    {
        CloseServiceHandle(hService);
        hService = NULL;
    }

    if (hSCManager != NULL)
    {
        CloseServiceHandle(hSCManager);
        hSCManager = NULL;
    }

    return result;
}

int UninstallDriver(wchar_t *pDriverName)
{
    int result = 0;

    SC_HANDLE hSCManager = NULL;
    SC_HANDLE hService = NULL;

    do
    {
        hSCManager = OpenSCManager(NULL, NULL, SC_MANAGER_ALL_ACCESS);
        if (hSCManager == NULL)
        {
            wprintf(L"[UninstallDriver] hSCManager == NULL");
            break;
        }

        hService = OpenService(hSCManager, pDriverName, DELETE);
        if (hService == NULL)
        {
            wprintf(L"[UninstallDriver] hService == NULL");
            break;
        }

        if (::DeleteService(hService) == TRUE)
        {
            result = 1;
        }
    } while (false);

    if (hService != NULL)
    {
        CloseServiceHandle(hService);
        hService = NULL;
    }

    if (hSCManager != NULL)
    {
        CloseServiceHandle(hSCManager);
        hSCManager = NULL;
    }

    return result;
}
