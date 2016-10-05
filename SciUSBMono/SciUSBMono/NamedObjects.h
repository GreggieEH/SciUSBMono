#pragma once

class CMyObject;

class CNamedObjects
{
public:
	CNamedObjects(void);
	~CNamedObjects(void);
	// add or remove objects from the list
	void				AddNamedObject(
							CMyObject*	pMyObject,
							BOOL		fAddObject);
	// get an object name
	BOOL				GetObjectName(
							CMyObject*	pMyObject,
							LPTSTR	*	szName);
	// check if the page selected flag is set
	BOOL				GetPageSelected(
							CMyObject*	pMyObject);
	// get the name for the first non selected page
	BOOL				GetNonSelectedName(
							LPTSTR	*	szName);
protected:
	// find a named object
	class CLink;
	BOOL				FindNamedObject(
							CMyObject*	pMyObject,
							CLink	**	ppLink);
private:
	// link list element
	class CLink
	{
	public:
		CLink(CMyObject * pMyObject);
		~CLink();
		CMyObject	*	GetMyObject();
		CLink		*	GetPrev();
		void			SetPrev(
							CLink	*	pPrev);
		CLink		*	GetNext();
		void			SetNext(
							CLink	*	pNext);
	private:
		CMyObject	*	m_pMyObject;
		CLink		*	m_pPrev;
		CLink		*	m_pNext;
	};
	// linked list of object names
	long				m_numNamedObjects;
	CLink			*	m_pStart;				// start of the list
	CLink			*	m_pEnd;					// end of the list
};
