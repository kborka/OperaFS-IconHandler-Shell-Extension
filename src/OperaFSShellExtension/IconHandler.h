#pragma once
#include <Unknwn.h>
#include <ShlObj.h>

class IconHandler : 
	public IPersistFile, 
	public IExtractIconW
{
public:
	// IUnknown implementation
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();

	// IPersistFile implemenation
	IFACEMETHODIMP GetClassID(CLSID *pclsid) { return E_NOTIMPL; };
	IFACEMETHODIMP GetCurFile(LPOLESTR *ppszFileName) { return E_NOTIMPL; };
	IFACEMETHODIMP IsDirty() { return E_NOTIMPL; };
	IFACEMETHODIMP Load(LPCOLESTR pszFileName, DWORD dwMode);
	IFACEMETHODIMP Save(LPCOLESTR pszFileName, BOOL fRemember) { return E_NOTIMPL; };
	IFACEMETHODIMP SaveCompleted(LPCOLESTR pszFileName) { return E_NOTIMPL; };

	// IInitializeWithStream Implementation
	//IFACEMETHODIMP Initialize(IStream *pstream, DWORD grfMode);

	// IExtractIcon implemenatation
	IFACEMETHODIMP Extract(PCWSTR pszFile, UINT nIconIndex, HICON *phiconLarge, HICON *phiconSmall, UINT nIconSize);
	IFACEMETHODIMP GetIconLocation(UINT uFlags, PWSTR pszIconFile, UINT cchMax, int *piIndex, UINT *pwFlags);

	IconHandler();

protected:
	~IconHandler();

private:
	long m_cRef;
	wchar_t m_szFileName[MAX_PATH];
	DWORD m_dwMode;
	//IStream *m_pStream;
};

