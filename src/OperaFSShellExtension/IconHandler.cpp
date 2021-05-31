#include "pch.h"
#include "IconHandler.h"
#include "resource.h"
#include <Shlwapi.h>
#include <strsafe.h>
#include <guiddef.h>
#pragma comment(lib, "shlwapi.lib")

extern long g_cDllRef;
extern HINSTANCE g_hInst;

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

/// <summary>
/// Initializes a handler with a stream.
/// </summary>
/// <param name="pstream">
/// A pointer to the stream source.
/// </param>
/// <param name="grfMode">
/// A STGM value that indicates the access mode of the pstream.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// Only one stream should be initialized per instance of IconHandler.
/// </remarks>
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

/// <summary>
/// Gets the location and index of an icon.
/// </summary>
/// <param name="uFlags">
/// Flag information to determine how to get the icon. Can be NULL.
/// </param>
/// <param name="pszIconFile">
/// A pointer to the buffer that gets the icon location.
/// The location is a null-terminated string that identifies the file that contains the icon.
/// </param>
/// <param name="cchMax">
/// The size of the buffer, in characters, pointed to by pszIconFile.
/// </param>
/// <param name="piIndex">
/// A pointer to an int that recieves the index of the icon in the fild pointed to by pszIconFile.
/// </param>
/// <param name="pwFlags">
/// A UINT value that gets a zero of a combination of flag values.
/// </param>
/// <returns>
/// The HRESULT value of setting the registry key.
/// On success, the value will be S_OK. On fail, the value will be an error code.
/// </returns>
/// <remarks>
/// For flag information, see https://docs.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-iextracticona-geticonlocation
/// </remarks>
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