#include "pch.h"
#include "ComRegister.h"
#include <combaseapi.h>
#include <strsafe.h>

/// <Summary>
/// Create a registry key under HKEY_CLASSES_ROOT and set a specific value.
/// </Summary>
/// <param name="pszSubKey">
/// The specific registry key to be affected. If the registry key doesn't exist, it will be created.
/// </param>
/// <param name="pszValueName">
/// The registry value to be set. A NULL value will set the default registry value.
/// </param>
/// <param name="pszData">
/// The data of the registry value.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
HRESULT SetHKCRRegistryKeyAndValue(PCWSTR pszSubKey, PCWSTR pszValueName,
    PCWSTR pszData)
{
    HRESULT hr;
    HKEY hKey = NULL;

    // Creates the specified registry key. If the key already exists, the 
    // function opens it. 
    hr = HRESULT_FROM_WIN32(RegCreateKeyEx(HKEY_CLASSES_ROOT, pszSubKey, 0,
        NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL));

    if (SUCCEEDED(hr))
    {
        if (pszData != NULL)
        {
            // Set the specified value of the key.
            DWORD cbData = lstrlen(pszData) * sizeof(*pszData);
            hr = HRESULT_FROM_WIN32(RegSetValueEx(hKey, pszValueName, 0,
                REG_SZ, reinterpret_cast<const BYTE *>(pszData), cbData));
        }

        RegCloseKey(hKey);
    }

    return hr;
}

/// <Summary>
/// Gets the data from a specified registry key and registry value name under HKEY_CLASSES_ROOT.
/// </Summary>
/// <param name="pszSubKey">
/// The specific registry key to be accessed. If the registry key doesn't exist, the function returns an error.
/// </param>
/// <param name="pszValueName">
/// The registry value to be retrieved. A NULL value will set the default registry value.
/// </param>
/// <param name="pszData">
/// A pointer to the buffer that will get the registry value's data.
/// </param>
/// <param name="cbData">
/// The size of the buffer in bytes.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
HRESULT GetHKCRRegistryKeyAndValue(PCWSTR pszSubKey, PCWSTR pszValueName,
    PWSTR pszData, DWORD cbData)
{
    HRESULT hr;
    HKEY hKey = NULL;

    // Try to open the specified registry key. 
    hr = HRESULT_FROM_WIN32(RegOpenKeyEx(HKEY_CLASSES_ROOT, pszSubKey, 0,
        KEY_READ, &hKey));

    if (SUCCEEDED(hr))
    {
        // Get the data for the specified value name.
        hr = HRESULT_FROM_WIN32(RegQueryValueEx(hKey, pszValueName, NULL,
            NULL, reinterpret_cast<LPBYTE>(pszData), &cbData));

        RegCloseKey(hKey);
    }

    return hr;
}

/// <Summary>
/// Register the in-process component in the registry.
/// </Summary>
/// <param name="pszModule">
/// Path of the module that contains the component.
/// </param>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Preferably a description instead of a key.
/// </param>
/// <param name="pszThreadModel">
/// The threading model of the component. In most cases it should be "Apartment".
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function creates the HKEY_CLASSES_ROOT\CLSID\{CLSID} key in the registry, as well as its InprocServer32 subkey.
/// </remarks>
HRESULT RegisterInprocServer(PCWSTR pszModule, const CLSID& clsid, PCWSTR pszFriendlyName, PCWSTR pszThreadModel)
{
    if (pszModule == NULL || pszThreadModel == NULL)
    {
        return E_INVALIDARG;
    }

    HRESULT hr;

    wchar_t szCLSID[MAX_PATH];
    StringFromGUID2(clsid, szCLSID, ARRAYSIZE(szCLSID));

    wchar_t szSubkey[MAX_PATH];

    // Create HKEY_CLASSES_ROOT\CLSID\{clsid} registry key.
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    if (SUCCEEDED(hr))
    {
        // Create friendly name under CLSID.
        hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, pszFriendlyName);
        if (SUCCEEDED(hr))
        {
            // Create InprocServer32 registry key under CLSID.
            hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s\\InprocServer32", szCLSID);
            if (SUCCEEDED(hr))
            {
                // Set the default value of the InprocServer32 key to the path of the COM module.
                hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, pszModule);
                if (SUCCEEDED(hr))
                {
                    // Set the threading model.
                    hr = SetHKCRRegistryKeyAndValue(szSubkey, L"ThreadingModel", pszThreadModel);
                }
            }
        }
    }

    return hr;
}

/// <Summary>
/// Unegister the in-process component in the registry.
/// </Summary>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function deletes the HKEY_CLASSES_ROOT\CLSID\{CLSID} key in the registry.
/// </remarks>
HRESULT UnregisterInprocServer(const CLSID& clsid)
{
    HRESULT hr = S_OK;

    wchar_t szCLSID[MAX_PATH];
    StringFromGUID2(clsid, szCLSID, ARRAYSIZE(szCLSID));

    wchar_t szSubkey[MAX_PATH];

    // Delete HKEY_CLASSES_ROOT\CLSID\{clsid} registry key.
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(RegDeleteTree(HKEY_CLASSES_ROOT, szSubkey));
    }

    return hr;
}

/// <Summary>
/// Register the IconHandler in the registry.
/// </Summary>
/// <param name="pszFileType">
/// The file type that this IconHandler will check against.
/// </param>
/// <param name="clsid">
/// The Class ID of the component.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Should be an intermediate registry key, such as {ProgramName}.{FileType}.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// This function creates the following registry key structure:
/// HKEY_CLASSES_ROOT
///   {pszFileType}
///       (Default) = {pszFriendlyName}
///   {pszFriendlyName}
///       (Default) = Description
///        DefaultIcon
///            (Default) = %1
///        Shellex
///            IconHandler
///                (Default) = {clsid}
/// </remarks>
HRESULT RetisterShellExtIconHandler(PCWSTR pszFileType, const CLSID& clsid, PCWSTR pszFriendlyName)
{
    if (pszFileType == NULL)
    {
        return E_INVALIDARG;
    }

    HRESULT hr;

    wchar_t szCLSID[MAX_PATH];
    StringFromGUID2(clsid, szCLSID, ARRAYSIZE(szCLSID));

    wchar_t szSubkey[MAX_PATH];

    
    // Always assume no keys exist for file type already.
    // Create the registry key for the file type and set the default value to the friendly name.
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s", pszFileType);
    if (SUCCEEDED(hr))
    {
        hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, pszFriendlyName);

        if (FAILED(hr))
        {
            return hr;
        }
    }

    // Since we have succeeded in associating the friendly name with the file type, let's generate
    // the friendly name registry key that will point to the IconHandler.
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s", pszFriendlyName);
    if (SUCCEEDED(hr))
    {
        // Register a description as the default value.
        hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, L"OperaFS Tools");
        if (SUCCEEDED(hr))
        {
            hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s\\DefaultIcon", pszFriendlyName);
            if (SUCCEEDED(hr))
            {
                // Set DefaultIcon value to %1 (this is reqired for IconHandler registration)
                hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, L"%1");
                if (SUCCEEDED(hr))
                {
                    //Set IconHandler CLSID
                    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s\\Shellex\\IconHandler", pszFriendlyName);
                    if (SUCCEEDED(hr))
                    {
                        hr = SetHKCRRegistryKeyAndValue(szSubkey, NULL, szCLSID);
                    }
                }
            }
        }
    }

    return hr;

}

/// <Summary>
/// Unregister the IconHandler from the registry.
/// </Summary>
/// <param name="pszFileType">
/// The file type this icon handler is associated with.
/// </param>
/// <param name="pszFriendlyName">
/// The friendly name of the component. Should be an intermediate registry key such as {ProgramName}.{FileType}.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
HRESULT UnregisterShellExtIconHandler(PCWSTR pszFileType, PCWSTR pszFriendlyName)
{
    if (pszFileType == NULL)
    {
        return E_INVALIDARG;
    }

    HRESULT hr;

    wchar_t szSubkey[MAX_PATH];
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s", pszFileType);
    if (SUCCEEDED(hr))
    {
        // Remove file type associations.
        hr = HRESULT_FROM_WIN32(RegDeleteTree(HKEY_CLASSES_ROOT, szSubkey));
        if (SUCCEEDED(hr))
        {
            hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"%s", pszFriendlyName);
            if (SUCCEEDED(hr))
            {
                // Remove friendly name registry tree.
                hr = HRESULT_FROM_WIN32(RegDeleteTree(HKEY_CLASSES_ROOT, szSubkey));
            }
        }
    }

    return hr;
}