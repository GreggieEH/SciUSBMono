#include "stdafx.h"
#include "MyFactory.h"
#include "MyObject.h"
#include "MyPropPage.h"

CMyFactory::CMyFactory(BOOL fPropPage) :
	m_cRefs(0),
	m_fPropPage(fPropPage)
{
}

CMyFactory::~CMyFactory(void)
{
}

// IUnknown methods
STDMETHODIMP CMyFactory::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IClassFactory == riid)
	{
		*ppv = this;
		this->AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyFactory::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyFactory::Release()
{
	ULONG				cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		this->m_cRefs++;
		ObjectsDown();				// decrement the object count
		delete this;
	}
	return cRefs;
}

// IClassFactory methods
STDMETHODIMP CMyFactory::CreateInstance(
									IUnknown	*	pUnkOuter,
									REFIID			riid,
									void		**	ppv)
{
	// check the input parameters
	HRESULT				hr;
	CMyObject		*	pCob;		// object
	CMyPropPage		*	pPropPage;	// property page

	if (this->m_fPropPage)
	{
		// create a property page
		if (NULL != pUnkOuter) return CLASS_E_NOAGGREGATION;
		pPropPage	= new CMyPropPage();
		hr = pPropPage->QueryInterface(riid, ppv);
		if (FAILED(hr)) delete pPropPage;
	}
	else
	{
		if (NULL != pUnkOuter && riid != IID_IUnknown)
			hr = CLASS_E_NOAGGREGATION;
		else
		{
			// let's create the object
			pCob = new CMyObject(pUnkOuter);
			if (pCob != NULL)
			{
				// increment the object count
				ObjectsUp();
				// initialize the object
				hr = pCob->Init();
				if (SUCCEEDED(hr))
				{
					hr = pCob->QueryInterface(riid, ppv);
				}
				if (FAILED(hr))
				{
					ObjectsDown();
					delete pCob;
				}
			}
			else
				hr = E_OUTOFMEMORY;
		}
	}
	return hr;
}

STDMETHODIMP CMyFactory::LockServer(
									BOOL			fLock)
{
	if (fLock)
		ObjectsUp();
	else
		ObjectsDown();
	return S_OK;
}