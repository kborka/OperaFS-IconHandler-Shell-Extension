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

IconHandler::IconHandler() : m_cRef(1), m_pStream(NULL)
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
		QITABENT(IconHandler, IInitializeWithStream),
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
IFACEMETHODIMP IconHandler::Initialize(IStream *pstream, DWORD grfMode)
{
	HRESULT hr = HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);

	if (m_pStream == NULL)
	{
		hr = pstream->QueryInterface(&m_pStream);
	}

	return hr;
}

// IExtractIcon Implementation
IFACEMETHODIMP IconHandler::Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize)
{
	return S_FALSE;
}

IFACEMETHODIMP IconHandler::GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags)
{
	char duck[] = { 'd', 'u', 'c', 'k' };
	char duckBytes[4];

	m_pStream->Seek({ 132 }, STREAM_SEEK_SET, NULL);
	m_pStream->Read(duckBytes, ARRAYSIZE(duckBytes), NULL);

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