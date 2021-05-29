#include "pch.h"
#include "ComRegister.h"
#include <combaseapi.h>
#include <strsafe.h>

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
	// Create the registry key for the file type and set the defautl value to the friendly name.
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