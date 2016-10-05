#pragma once

class CMyEnumFormatetc : public IEnumFORMATETC
{
public:
	CMyEnumFormatetc(IUnknown * pHostObj);
	~CMyEnumFormatetc(void);
	// IUnknown methods
	STDMETHODIMP			QueryInterface(
								REFIID					riid,
								LPVOID*					ppv);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();
	// IEnumConnections methods
	STDMETHODIMP			Next(
								ULONG					cFormats, 
								FORMATETC			*	pFormats, 
								ULONG				*	cFetched);
    STDMETHODIMP			Skip(
								ULONG					cSkip);
    STDMETHODIMP			Reset();
    STDMETHODIMP			Clone(
								IEnumFORMATETC		**	ppFormats);
	// initialization
	HRESULT					Init(
								ULONG					cFormats,
								FORMATETC			*	paFormats,
								ULONG					iEnumIndex);
private:
	IUnknown			*	m_pHostObj;				// host object
	ULONG					m_cRefs;				// reference count
	ULONG					m_cFormats;				// number of formats
	FORMATETC			*	m_pFormats;				// array of formats
	ULONG					m_iCurIndex;			// current index of the array
};
