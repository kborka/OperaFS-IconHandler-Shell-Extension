// dllmain.cpp : Defines the entry point for the DLL application.
#include "pch.h"
#include "ClassFactory.h"
#include "ComRegister.h"

// A CLSID is required for registration of the extension in the registry.
// This extension's CLSID is: { e0b60a9b-9f69-4ca6-ac9b-edbb502a9fce }
const CLSID CLSID_OperaFSShellExtension =
{ 0xe0b60a9b, 0x9f69, 0x4ca6, { 0xac, 0x9b, 0xed, 0xbb, 0x50, 0x2a, 0x9f, 0xce } };

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

/// <Summary>
/// Initializes an instance of the <see cref="ClassFactory"> class and query to the specific interface.
/// </Summary>
/// <param name="rclsid"> 
/// The CLSID that will be checked against and registered.
/// </param>
/// <param name="riid"> 
/// A reference to the identifier of the interface that the caller will use to communicate with the class object.
/// </param>
/// <param name="ppv"> 
/// The address of a pointer variable that gets the interface pointer requested in <see cref="riid">.
/// On success, *ppv contains the requested interface pointer. On error or fail, the pointer is NULL.
/// </param>
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

/// <Summary>
/// Check if the DLL can be unloaded from memory.
/// </Summary>
STDAPI DllCanUnloadNow(void)
{
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

/// <Summary>
/// Register the COM server and its components.
/// </Summary>
STDAPI DllRegisterServer(void)
{
    HRESULT hr;

    wchar_t szModule[MAX_PATH];
    if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        return hr;
    }

    // Register the component, creating the CLSID in the registry.
    hr = RegisterInprocServer(szModule, CLSID_OperaFSShellExtension, L"OperaFS Tools", L"Apartment");
    if (SUCCEEDED(hr))
    {
        // Register the IconHandler, creating the file extension, friendly name, and IconHandler registry keys.
        hr = RetisterShellExtIconHandler(L".operafs", CLSID_OperaFSShellExtension, L"3DO.OperaFS");		
    }

    return hr;
}

/// <Summary>
/// Unregister the COM server and its components.
/// </Summary>
STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;

    wchar_t szModule[MAX_PATH];
    if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        return hr;
    }

    // Delete the CLSID key tree.
    hr = UnregisterInprocServer(CLSID_OperaFSShellExtension);
    if (SUCCEEDED(hr))
    {
        // Delete the file extension and friendly name trees.
        hr = UnregisterShellExtIconHandler(L".operafs", L"3DO.OperaFS");
    }

    return hr;
}