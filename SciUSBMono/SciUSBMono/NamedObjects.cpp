#include "stdafx.h"
#include "NamedObjects.h"
#include "MyObject.h"
#include "dispids.h"

CNamedObjects::CNamedObjects(void) :
	// linked list of object names
	m_numNamedObjects(0),
	m_pStart(NULL),					// start of the list
	m_pEnd(NULL)					// end of the list
{
}

CNamedObjects::~CNamedObjects(void)
{
	// clear the list
	CLink		*	pLink	= this->m_pStart;
	CLink		*	pNext;
	while (NULL != pLink)
	{
		pNext	= pLink->GetNext();
		delete [] pLink;
		pLink	= pNext;
	}
	// make it official
	this->m_pStart	= NULL;
	this->m_pEnd	= NULL;
	this->m_numNamedObjects	= 0;
}


// add or remove objects from the list
void CNamedObjects::AddNamedObject(
							CMyObject*	pMyObject,
							BOOL		fAddObject)
{
	CLink			*	pLink;			// link element
	CLink			*	pPrev;			// previous link element
	CLink			*	pNext;			// next link element
	if (fAddObject)
	{
		// add the object to the list
		// make sure that the object doesn't currently exist
		if (!this->FindNamedObject(pMyObject, &pLink))
		{
			// create a new link
			pLink	= new CLink(pMyObject);
			// add to the list
			if (NULL == this->m_pStart)
			{
				// first list element
				this->m_pStart	= pLink;
				this->m_pEnd	= pLink;
				this->m_numNamedObjects	= 1;
			}
			else
			{
				// add to the end of the list
				this->m_pEnd->SetNext(pLink);
				pLink->SetPrev(this->m_pEnd);
				this->m_pEnd	= pLink;
				this->m_numNamedObjects++;
			}
		}
	}
	else
	{
		// remove the object from the list
		// find the element
		if (this->FindNamedObject(pMyObject, &pLink))
		{
			// previous and next elements
			pPrev = pLink->GetPrev();
			pNext = pLink->GetNext();
			// delete the element
			delete pLink;
			// decrement the list count
			this->m_numNamedObjects--;
			// recreate the list, 4 cases
			if (NULL != pNext)
			{
				if (NULL != pPrev)
				{
					// somewhere inside the list
					pPrev->SetNext(pNext);
					pNext->SetPrev(pPrev);
				}
				else
				{
					// at the start of the list
					pNext->SetPrev(NULL);
					this->m_pStart	 = pNext;
				}
			}
			else			
			{
				// pNext = NULL;
				if (NULL != pPrev)
				{
					// at the end of the list
					pPrev->SetNext(NULL);
					this->m_pEnd	= pPrev;
				}
				else
				{
					// no list is left, make it official
					this->m_pStart		= NULL;
					this->m_pEnd		= NULL;
					this->m_numNamedObjects	= 0;
				}
			}
		}
	}
}

// find a named object
BOOL CNamedObjects::FindNamedObject(
							CMyObject*			pMyObject,
							CNamedObjects::CLink**	ppLink)
{
	CLink			*	pLink;
	BOOL				fDone;
	*ppLink		= NULL;
	pLink	= this->m_pStart;
	fDone	= FALSE;
	while (!fDone && NULL != pLink)
	{
		// compare the pointers
		if (pLink->GetMyObject() == pMyObject)
		{
			fDone = TRUE;
			*ppLink	= pLink;
		}
		if (!fDone)
		{
			// try the next element
			pLink = pLink->GetNext();
		}
	}
	return fDone;
}

// get an object name
BOOL CNamedObjects::GetObjectName(
							CMyObject*	pMyObject,
							LPTSTR	*	szName)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	*szName		= NULL;
	hr = pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		hr = Utils_GetStringProperty(pdisp, DISPID_DisplayName, szName);
		pdisp->Release();
	}
//	hr = pMyObject->GetName(szName);
	return SUCCEEDED(hr);
}

// check if the page selected flag is set
BOOL CNamedObjects::GetPageSelected(
							CMyObject*	pMyObject)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	BOOL				fPageSelected	= FALSE;
	hr = pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		fPageSelected	= Utils_GetBoolProperty(pdisp, DISPID_PageSelected);
		pdisp->Release();
	}
	return fPageSelected;
}

// get the name for the first non selected page
BOOL CNamedObjects::GetNonSelectedName(
							LPTSTR	*	szName)
{
	HRESULT				hr;
	CLink			*	pLink;
	CMyObject		*	pMyObject;
	BOOL				fDone;
	IDispatch		*	pdisp;
	*szName		= NULL;
	pLink	= this->m_pStart;
	fDone	= FALSE;
	while (!fDone && NULL != pLink)
	{
		pMyObject = pLink->GetMyObject();
		if (!this->GetPageSelected(pMyObject))
		{
			fDone = this->GetObjectName(pMyObject, szName);
			// set the page selected
			hr = pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
			if (SUCCEEDED(hr))
			{
				Utils_SetBoolProperty(pdisp, DISPID_PageSelected, TRUE);
				pdisp->Release();
			}
		}
		if (!fDone)
			pLink = pLink->GetNext();
	}
	return fDone;
}

CNamedObjects::CLink::CLink(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pPrev(NULL),
	m_pNext(NULL)
{
	// Add ref the object
	this->m_pMyObject->AddRef();
}

CNamedObjects::CLink::~CLink()
{
	// release our added reference
	Utils_RELEASE_INTERFACE(this->m_pMyObject);
}

CMyObject* CNamedObjects::CLink::GetMyObject()
{
	return this->m_pMyObject;
}

CNamedObjects::CLink* CNamedObjects::CLink::GetPrev()
{
	return this->m_pPrev;
}

void CNamedObjects::CLink::SetPrev(
							CNamedObjects::CLink	*	pPrev)
{
	this->m_pPrev	= pPrev;
}

CNamedObjects::CLink* CNamedObjects::CLink::GetNext()
{
	return this->m_pNext;
}

void CNamedObjects::CLink::SetNext(
							CNamedObjects::CLink	*	pNext)
{
	this->m_pNext = pNext;
}