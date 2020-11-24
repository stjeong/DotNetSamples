// ConsoleApplication1.cpp : Defines the entry point for the console application.
//

#include "resource.h"

#include <atlbase.h>
#include <atlsafe.h>

#include <MetaHost.h>
#pragma comment(lib, "MSCorEE.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only				\
    high_property_prefixes("_get","_put","_putref")		\
    rename("ReportEvent", "InteropServices_ReportEvent") 
	//rename("or", "_or_")
using namespace mscorlib;

int _tmain(int argc, _TCHAR* argv[])
{
	HRESULT hr;

	ICLRMetaHost* pMetaHost = nullptr;
	ICLRRuntimeInfo* pRuntimeInfo = nullptr;
	ICorRuntimeHost* pCorRuntimeHost = nullptr;

	do
	{
		hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
		if (FAILED(hr))
		{
			wprintf(L"CLRCreateInstance failed: 0x%x (.NET 4 not installed)\n", hr);

			// 실패한다면 .NET 2.0 로드를 시도
			PCWSTR pszFlavor = L"wks";
			PCWSTR pszVersion = L"v2.0.50727";

			hr = CorBindToRuntimeEx(
				pszVersion,                     // Runtime version 
				pszFlavor,                      // Flavor of the runtime to request 
				0,                              // Runtime startup flags 
				CLSID_CorRuntimeHost,           // CLSID of ICorRuntimeHost 
				IID_PPV_ARGS(&pCorRuntimeHost)  // Return ICorRuntimeHost 
			);

			if (FAILED(hr))
			{
				wprintf(L".NET 2.0 load failed\n");
				break;
			}
			else
			{
				wprintf(L".NET 2.0 loaded\n");
			}
		}
		else
		{
			// 우선 .NET 4.0 로드를 시도
			hr = pMetaHost->GetRuntime(L"v4.0.30319", IID_PPV_ARGS(&pRuntimeInfo));
			if (FAILED(hr))
			{
				wprintf(L".NET 4.0 load failed\n");
				break;
			}
			else
			{
				wprintf(L".NET 4.0 loaded\n");
			}

			hr = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&pCorRuntimeHost));
			if (FAILED(hr))
			{
				wprintf(L"GetInterface(CLSID_CorRuntimeHost) failed\n");
				break;
			}
		}

		hr = pCorRuntimeHost->Start(); // CLR 구동
		if (FAILED(hr))
		{
			wprintf(L"CLR failed to start\n");
			break;
		}

		{
			IUnknownPtr spAppDomainThunk = nullptr;
			_AssemblyPtr spAssembly = nullptr;

			hr = pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk);
			if (FAILED(hr))
			{
				wprintf(L"GetDefaultDomain failed\n");
				break;
			}

			_AppDomainPtr spDefaultAppDomain = spAppDomainThunk;
			if (spDefaultAppDomain == nullptr)
			{
				wprintf(L"Failed to get default _AppDomainPtr\n");
				break;
			}

			HMODULE hModule = ::GetModuleHandle(NULL);
			HRSRC hResource = FindResource(hModule, MAKEINTRESOURCE(IDR_EXEFILE1), L"EXEFILE");
			HGLOBAL hMemory = LoadResource(hModule, hResource);
			DWORD dwSize = SizeofResource(hModule, hResource);
			LPVOID lpAddress = LockResource(hMemory);

			{
				SAFEARRAYBOUND rgsabound[] = { dwSize, 0 };
				SAFEARRAY* pSafeArray = SafeArrayCreate(VT_UI1, 1, rgsabound);
				pSafeArray->pvData = lpAddress;

				hr = spDefaultAppDomain->Load_3(pSafeArray, &spAssembly);

				/*
				Critical error detected c0000374
				ConsoleApplication1.exe has triggered a breakpoint.
				*/
				// SafeArrayDestroy(pSafeArray);
			}

			if (FAILED(hr))
			{
				wprintf(L"Failed to load the assembly\n");
				break;
			}

			_MethodInfoPtr mainMethod;
			hr = spAssembly->get_EntryPoint(&mainMethod);

			if (FAILED(hr))
			{
				wprintf(L"No Entry method\n");
				break;
			}

			VARIANT vtEmpty;
			VariantInit(&vtEmpty);

			BindingFlags flags = (BindingFlags)(BindingFlags::BindingFlags_InvokeMethod | BindingFlags::BindingFlags_Static);
			hr = mainMethod->Invoke_2(vtEmpty, flags,
				nullptr, nullptr, nullptr, nullptr);

			if (FAILED(hr))
			{
				wprintf(L"Failed to call main method\n");
				break;
			}
		}

		pCorRuntimeHost->Stop();

	} while (false);

	if (pCorRuntimeHost != nullptr)
	{
		pCorRuntimeHost->Release();
	}

	if (pRuntimeInfo != nullptr)
	{
		pRuntimeInfo->Release();
	}

	if (pMetaHost != nullptr)
	{
		pMetaHost->Release();
	}

	return 0;
}

