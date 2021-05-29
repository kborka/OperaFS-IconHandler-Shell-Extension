#pragma once
#include <Windows.h>

HRESULT RegisterInprocServer(PCWSTR pszModule, const CLSID& clsid, PCWSTR pszFriendlyName, PCWSTR pszThreadModel);

HRESULT UnregisterInprocServer(const CLSID& clsid);

HRESULT RetisterShellExtIconHandler(PCWSTR pszFileType, const CLSID& clsid, PCWSTR pszFriendlyName);

HRESULT UnregisterShellExtIconHandler(PCWSTR pszFileType, PCWSTR pszFriendlyName);