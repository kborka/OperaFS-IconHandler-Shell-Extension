#pragma once
#include <Unknwn.h>
#include <ShlObj.h>

class IconHandler : 
	public IInitializeWithStream,
	public IExtractIconW
{
public:
	// IUnknown implementation
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();

	// IInitializeWithStream Implementation
	IFACEMETHODIMP Initialize(IStream *pstream, DWORD grfMode);

	// IExtractIcon implemenatation
	IFACEMETHODIMP Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize);
	IFACEMETHODIMP GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags);

	IconHandler();

protected:
	~IconHandler();

private:
	long m_cRef;
	IStream *m_pStream;
};

