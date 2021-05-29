// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "ClassFactory.h"
#include "ComRegister.h"

// A CLSID is required for registration of the extension in the registry.
// This extension's CLSID is: { e0b60a9b-9f69-4ca6-ac9b-edbb502a9fce }
const CLSID CLSID_OperaFSShellExtension =
{ 0xdeadbeef, 0xdead, 0xbeef, { 0xde, 0xad, 0xbe, 0xef, 0xde, 0xad, 0xbe, 0xef } };

// Instance of this DLL module.
HINSTANCE g_hInst = NULL;
// Refrence count of this DLL module.
long g_cDllRef = 0;

BOOL APIENTRY DllMain( HMODULE hModule,
                       DWORD  ul_reason_for_call,
                       LPVOID lpReserved
                     )
{
    switch (ul_reason_for_call)
    {
	case DLL_PROCESS_ATTACH:
		g_hInst = hModule;
		DisableThreadLibraryCalls(hModule);
		break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
	HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

	if (IsEqualCLSID(CLSID_OperaFSShellExtension, rclsid))
	{
		hr = E_OUTOFMEMORY;
		
		ClassFactory *pClassFactory = new ClassFactory();
		if (pClassFactory)
		{
			hr = pClassFactory->QueryInterface(riid, ppv);
			pClassFactory->Release();
		}
	}

	return hr;
}

STDAPI DllCanUnloadNow(void)
{
	return g_cDllRef > 0 ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer(void)
{
	HRESULT hr;

	wchar_t szModule[MAX_PATH];
	if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		return hr;
	}

	hr = RegisterInprocServer(szModule, CLSID_OperaFSShellExtension, L"OperaFS Tools", L"Apartment");
	if (SUCCEEDED(hr))
	{
		hr = RetisterShellExtIconHandler(L".operafs", CLSID_OperaFSShellExtension, L"3DO.OperaFS");		
	}

	return hr;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT hr = S_OK;

	wchar_t szModule[MAX_PATH];
	if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		return hr;
	}

	hr = UnregisterInprocServer(CLSID_OperaFSShellExtension);
	if (SUCCEEDED(hr))
	{
		hr = UnregisterShellExtIconHandler(L".operafs", L"3DO.OperaFS");
	}

	return hr;
}