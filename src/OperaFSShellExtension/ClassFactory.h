#pragma once

#include <Unknwn.h>
#include <Windows.h>

class ClassFactory : public IClassFactory
{
public:
	// IUnknown implementation
	IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
	IFACEMETHODIMP_(ULONG) AddRef();
	IFACEMETHODIMP_(ULONG) Release();

	// IClassFactory implementation
	IFACEMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv);
	IFACEMETHODIMP LockServer(BOOL fLock);

	ClassFactory();

protected:
	~ClassFactory();

private:
	long m_cRef;
};

