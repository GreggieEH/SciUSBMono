#pragma once

class CMyFactory : public IClassFactory
{
public:
	CMyFactory(BOOL fPropPage);
	~CMyFactory(void);
	// IUknown methods
	STDMETHODIMP			QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();
	// IClassFactory methods
	STDMETHODIMP			CreateInstance(
								IUnknown	*	pUnkOuter,
								REFIID			riid,
								void		**	ppvObject);
	STDMETHODIMP			LockServer(
								BOOL			fLock);
private:
	ULONG					m_cRefs;
	BOOL					m_fPropPage;
};
