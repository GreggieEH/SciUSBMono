#include "stdafx.h"
#include "MyEnumFormatetc.h"

CMyEnumFormatetc::CMyEnumFormatetc(IUnknown * pHostObj) :
	m_pHostObj(pHostObj),				// host object
	m_cRefs(0),							// reference count
	m_cFormats(0),						// number of formats
	m_pFormats(NULL),					// array of formats
	m_iCurIndex(0)						// current index of the array
{

}

CMyEnumFormatetc::~CMyEnumFormatetc()
{
	if (m_pFormats)
	{
		delete [] m_pFormats;
		m_pFormats = NULL;
	}
}

// initialization
HRESULT CMyEnumFormatetc::Init(
								ULONG					cFormats,
								FORMATETC			*	paFormats,
								ULONG					iEnumIndex)
{
	ULONG			i;

	m_cFormats = cFormats;
	if (m_cFormats > 0)
	{
		m_pFormats = new FORMATETC [m_cFormats];
		m_iCurIndex = iEnumIndex;
		// copy the array
		for (i=0; i<m_cFormats; i++)
		{
			m_pFormats[i].cfFormat	= paFormats[i].cfFormat;
			m_pFormats[i].dwAspect	= paFormats[i].dwAspect;
			m_pFormats[i].lindex	= paFormats[i].lindex;
			m_pFormats[i].ptd		= NULL;
			m_pFormats[i].tymed		= paFormats[i].tymed;
		}
	}
	else return E_UNEXPECTED;
	return S_OK;
}

// IUnknown methods
STDMETHODIMP CMyEnumFormatetc::QueryInterface(
								REFIID					riid,
								LPVOID*					ppv)
{
	if (IID_IUnknown == riid || IID_IEnumFORMATETC == riid)
	{
		*ppv = this;
		AddRef();
		return S_OK;
	}
	else
	{
		*ppv = NULL;
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyEnumFormatetc::AddRef()
{
	m_pHostObj->AddRef();
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyEnumFormatetc::Release()
{
	ULONG			cRefs;

	m_pHostObj->Release();
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		delete this;
	}
	return cRefs;
}

// IEnumConnections methods
STDMETHODIMP CMyEnumFormatetc::Next(
								ULONG					cFormats, 
								FORMATETC			*	pFormats, 
								ULONG				*	pcFetched)
{
	ULONG			i;			// array index

	if (NULL == pcFetched && 1 != cFormats)
		return E_POINTER;
	if (m_iCurIndex >= m_cFormats) return S_FALSE;
	// loop over the array
	i = 0;
	while (i < cFormats && m_iCurIndex < m_cFormats)
	{
		// copy the format structure
		pFormats[i].cfFormat	= m_pFormats[m_iCurIndex].cfFormat;
		pFormats[i].dwAspect	= m_pFormats[m_iCurIndex].dwAspect;
		pFormats[i].lindex		= m_pFormats[m_iCurIndex].lindex;
		pFormats[i].ptd			= NULL;
		pFormats[i].tymed		= m_pFormats[m_iCurIndex].tymed;
		// increment the indices
		m_iCurIndex++;
		i++;
	}
	if (pcFetched)
	{
		*pcFetched = i;
	}
	return i == cFormats ? S_OK : S_FALSE;
}

STDMETHODIMP CMyEnumFormatetc::Skip(
								ULONG					cSkip)
{
	if ((m_iCurIndex + cSkip) < m_cFormats)
	{
		m_iCurIndex += cSkip;
		return S_OK;
	}
	else return S_FALSE;
}

STDMETHODIMP CMyEnumFormatetc::Reset()
{
	m_iCurIndex = 0;
	return S_OK;
}

STDMETHODIMP CMyEnumFormatetc::Clone(
								IEnumFORMATETC			**	ppEnumFormats)
{
	HRESULT					hr;
	CMyEnumFormatetc			*	pEnum;	

	*ppEnumFormats = NULL;			// null the output
	pEnum = new CMyEnumFormatetc(m_pHostObj);
	if (pEnum)
	{
		hr = pEnum->Init(m_cFormats, m_pFormats, m_iCurIndex);
		if (SUCCEEDED(hr))
		{
			hr = pEnum->QueryInterface(IID_IEnumFORMATETC,
				(LPVOID*) ppEnumFormats);
		}
		if (FAILED(hr))
			delete pEnum;
	}
	else hr = E_OUTOFMEMORY;
	return hr;
}
