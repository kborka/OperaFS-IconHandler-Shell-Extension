#include "pch.h"
#include "IconHandler.h"
#include "resource.h"
#include <Shlwapi.h>
#include <strsafe.h>
#include <guiddef.h>
#pragma comment(lib, "shlwapi.lib")

extern long g_cDllRef;
extern HINSTANCE g_hInst;

// A CLSID is required for registration of the extension in the registry.
// This extension's CLSID is: { e0b60a9b-9f69-4ca6-ac9b-edbb502a9fce }
const CLSID CLSID_OperaFSShellExtension =
{ 0xdeadbeef, 0xdead, 0xbeef, { 0xde, 0xad, 0xbe, 0xef, 0xde, 0xad, 0xbe, 0xef } };

IconHandler::IconHandler() : m_cRef(1)
{
	InterlockedIncrement(&g_cDllRef);
}

IconHandler::~IconHandler()
{
	InterlockedDecrement(&g_cDllRef);
}

// IUnknown Implementation
IFACEMETHODIMP IconHandler::QueryInterface(REFIID riid, void **ppv)
{
	static const QITAB qit[] =
	{
		QITABENT(IconHandler, IPersistFile),
		QITABENT(IconHandler, IExtractIconW),
		{ 0 },
	};

	return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) IconHandler::AddRef()
{
	return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) IconHandler::Release()
{
	ULONG cRef = InterlockedDecrement(&m_cRef);

	if (0 == cRef)
	{
		delete this;
	}

	return cRef;
}

// IInitializeWithStream Implementation
// TODO: Figure this out.
//IFACEMETHODIMP IconHandler::Initialize(IStream *pstream, DWORD grfMode)
//{
//	HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
//
//	if (m_pStream == NULL)
//	{
//		hr = pstream->QueryInterface(&m_pStream);
//	}
//
//	return hr;
//}

// IPersistFile Implementation
IFACEMETHODIMP IconHandler::Load(LPCOLESTR pszFileName, DWORD dwMode)
{
	m_dwMode = dwMode;
	return StringCchCopy(m_szFileName, ARRAYSIZE(m_szFileName), pszFileName);
}

// IExtractIcon Implementation
IFACEMETHODIMP IconHandler::Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize)
{
	return S_FALSE;
}

IFACEMETHODIMP IconHandler::GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags)
{
	HANDLE hFile;
	hFile = CreateFile(m_szFileName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);

	if (hFile == INVALID_HANDLE_VALUE)
	{
		return S_FALSE;
	}

	

	char duck[] = { 'd', 'u', 'c', 'k' };
	char duckBytes[4];

	// Move to 0x84.
	OVERLAPPED ol{};
	ol.Offset = 132;
	ReadFile(hFile, duckBytes, ARRAYSIZE(duckBytes), NULL, &ol);

	//m_pStream->Seek({ 132 }, STREAM_SEEK_SET, NULL);
	//m_pStream->Read(duckBytes, ARRAYSIZE(duckBytes), NULL);

	for (int i = 0; i < ARRAYSIZE(duck); i++)
	{
		if (duckBytes[i] != duck[i])
		{
			return S_FALSE;
		}
	}

	
	GetModuleFileName(g_hInst, pszIconFile, cchMax);		

	*piIndex = 0;
	*pwFlags = 0;

	return S_OK;
}