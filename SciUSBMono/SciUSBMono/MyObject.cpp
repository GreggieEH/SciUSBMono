#include "stdafx.h"
#include "MyObject.h"
#include "MyGuids.h"
#include "dispids.h"
#include "MyWindow.h"
#include "MySciUsbMono.h"
#include "DlgDisplayConfigValues.h"
#include "NamedObjects.h"
#include "MyEnumFormatetc.h"

CMyObject::CMyObject(IUnknown * pUnkOuter) :
	m_pImpIDispatch(NULL),
	m_pImpIProvideClassInfo2(NULL),
	m_pImpIConnectionPointContainer(NULL),
	m_pImpIPersistStorage(NULL),
	m_pImpIPersistStreamInit(NULL),
	m_pImpIPersistPropertyBag(NULL),
	m_pImpIDataObject(NULL),
	m_pImpIOleObject(NULL),
	m_pImpIOleControl(NULL),
	m_pImpIOleInPlaceObject(NULL),
	m_pImpIOleInPlaceActiveObject(NULL),
	m_pImpISpecifyPropertyPages(NULL),
	m_pImpIViewObject2(NULL),
	m_pImp_clsIMono(NULL),
	m_pImpISciMono(NULL),
	// object reference count
	m_cRefs(0),
	// outer unknown for aggregation
	m_pUnkOuter(pUnkOuter),
	// flags
	m_fAmLoaded(FALSE),			// am loaded flag
	m_fFreezeEvents(FALSE),		// freeze events flag
	// data advise holder
	m_pDataAdviseHolder(NULL),
	// window manager
	m_pMyWindow(NULL),
	// storage objects
	m_pStg(NULL),				// storage object
	m_pStm(NULL),				// stream object
	// monochromator manager
	m_pMySciUsbMono(NULL),
	// event message handlers
	m_pdispOnError(NULL),
	m_pdispOnStatusMessage(NULL),
	m_pdispOnHaveNewPosition(NULL),
	m_pdispOnChangedGrating(NULL),
	m_pdispOnAmInitialized(NULL),
	// _clsIMono interface id
	m_iid_clsIMono(IID_NULL),
	// __clsIMono events
	m_iid__clsIMono(IID_NULL),
	m_dispidGratingChanged(DISPID_UNKNOWN),
	m_dispidBeforeMoveChange(DISPID_UNKNOWN),
	m_dispidHaveAborted(DISPID_UNKNOWN),
	m_dispidMoveCompleted(DISPID_UNKNOWN),
	m_dispidMoveError(DISPID_UNKNOWN),
	m_dispidRequestChangeGrating(DISPID_UNKNOWN),
	m_dispidRequestParentWindow(DISPID_UNKNOWN),
	m_dispidStatusMessage(DISPID_UNKNOWN),
	m_dispidAmInitPropChanged(DISPID_UNKNOWN),
	m_dispidAutoGratingPropChanged(DISPID_UNKNOWN),
	// page selected flag
	m_fPageSelected(FALSE),
	// ISciMono
	m_iid_ISciMono(IID_NULL),
	m_iid__SciMono(IID_NULL),
	// _SciMono events
	m_dispidPropChanged(DISPID_UNKNOWN),
	m_dispidError(DISPID_UNKNOWN),
	m_dispidBusyStatusChange(DISPID_UNKNOWN),
	m_dispidMonochromatorPropChanged(DISPID_UNKNOWN),
	m_dispidGratingPropChanged(DISPID_UNKNOWN)
{
	if (NULL == this->m_pUnkOuter) this->m_pUnkOuter = this;
	// connection point array
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		this->m_paConnPts[i]	= NULL;
}

CMyObject::~CMyObject(void)
{
	Utils_DELETE_POINTER(m_pImpIDispatch);
	Utils_DELETE_POINTER(m_pImpIProvideClassInfo2);
	Utils_DELETE_POINTER(m_pImpIConnectionPointContainer);
	Utils_DELETE_POINTER(m_pImpIPersistStorage);
	Utils_DELETE_POINTER(m_pImpIPersistStreamInit);
	Utils_DELETE_POINTER(m_pImpIPersistPropertyBag);
	Utils_DELETE_POINTER(m_pImpIDataObject);
	Utils_DELETE_POINTER(m_pImpIOleObject);
	Utils_DELETE_POINTER(m_pImpIOleControl);
	Utils_DELETE_POINTER(m_pImpIOleInPlaceObject);
	Utils_DELETE_POINTER(m_pImpIOleInPlaceActiveObject);
	Utils_DELETE_POINTER(m_pImpISpecifyPropertyPages);
	Utils_DELETE_POINTER(m_pImpIViewObject2);
	Utils_DELETE_POINTER(m_pImp_clsIMono);
	Utils_DELETE_POINTER(m_pImpISciMono);
	// connection point array
	for (ULONG i=0; i<MAX_CONN_PTS; i++)
		Utils_RELEASE_INTERFACE(this->m_paConnPts[i]);
	Utils_RELEASE_INTERFACE(m_pDataAdviseHolder);
	// window manager
	Utils_DELETE_POINTER(m_pMyWindow);
	Utils_DELETE_POINTER(m_pStg);				// storage object
	Utils_DELETE_POINTER(m_pStm);				// stream object
	Utils_DELETE_POINTER(m_pMySciUsbMono);
	// event message handlers
	Utils_RELEASE_INTERFACE(m_pdispOnError);
	Utils_RELEASE_INTERFACE(m_pdispOnStatusMessage);
	Utils_RELEASE_INTERFACE(m_pdispOnHaveNewPosition);
	Utils_RELEASE_INTERFACE(m_pdispOnChangedGrating);
	Utils_RELEASE_INTERFACE(m_pdispOnAmInitialized);
}

// IUnknown methods
STDMETHODIMP CMyObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	*ppv = NULL;
	if (IID_IUnknown == riid)
		*ppv = this;
	else if (IID_IDispatch == riid || IID_ISciUsbMono == riid)
		*ppv = this->m_pImpIDispatch;
	else if (IID_IProvideClassInfo == riid || IID_IProvideClassInfo2 == riid)
		*ppv = this->m_pImpIProvideClassInfo2;
	else if (IID_IConnectionPointContainer == riid)
		*ppv = this->m_pImpIConnectionPointContainer;
	else if (IID_IPersist == riid || IID_IPersistStorage == riid)
		*ppv = this->m_pImpIPersistStorage;
	else if (IID_IPersistStreamInit == riid)
		*ppv = this->m_pImpIPersistStreamInit;
	else if (IID_IPersistPropertyBag == riid)
		*ppv = this->m_pImpIPersistPropertyBag;
	else if (IID_IDataObject == riid)
		*ppv = this->m_pImpIDataObject;
	else if (IID_IOleObject == riid)
		*ppv = this->m_pImpIOleObject;
	else if (IID_IOleControl == riid)
		*ppv = this->m_pImpIOleControl;
	else if (IID_IOleWindow == riid || IID_IOleInPlaceObject == riid)
		*ppv = this->m_pImpIOleInPlaceObject;
	else if (IID_IOleInPlaceActiveObject == riid)
		*ppv = this->m_pImpIOleInPlaceActiveObject;
	else if (IID_ISpecifyPropertyPages == riid)
		*ppv = this->m_pImpISpecifyPropertyPages;
	else if (IID_IViewObject == riid || IID_IViewObject2 == riid)
		*ppv = this->m_pImpIViewObject2;
	else if (riid == this->m_iid_clsIMono)
		*ppv = this->m_pImp_clsIMono;
	else if (riid == this->m_iid_ISciMono)
		*ppv = this->m_pImpISciMono;
	if (NULL != *ppv)
	{
		((IUnknown*)*ppv)->AddRef();
		return S_OK;
	}
	else
	{
		return E_NOINTERFACE;
	}
}

STDMETHODIMP_(ULONG) CMyObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyObject::Release()
{
	ULONG				cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
	{
		ObjectsDown();
		this->m_cRefs++;
		delete this;
	}
	return cRefs;
}

// initialization
HRESULT CMyObject::Init()
{
	HRESULT					hr;
	this->m_pImpIDispatch					= new CImpIDispatch(this, this->m_pUnkOuter);
	this->m_pImpIProvideClassInfo2			= new CImpIProvideClassInfo2(this, this->m_pUnkOuter);
	this->m_pImpIConnectionPointContainer	= new CImpIConnectionPointContainer(this, this->m_pUnkOuter);
	this->m_pImpIPersistStorage				= new CImpIPersistStorage(this, this->m_pUnkOuter);
	this->m_pImpIPersistStreamInit			= new CImpIPersistStreamInit(this, this->m_pUnkOuter);
	this->m_pImpIPersistPropertyBag			= new CImpIPersistPropertyBag(this, this->m_pUnkOuter);
	this->m_pImpIDataObject					= new CImpIDataObject(this, this->m_pUnkOuter);
	this->m_pImpIOleObject					= new CImpIOleObject(this, this->m_pUnkOuter);
	this->m_pImpIOleControl					= new CImpIOleControl(this, this->m_pUnkOuter);
	this->m_pImpIOleInPlaceObject			= new CImpIOleInPlaceObject(this, this->m_pUnkOuter);
	this->m_pImpIOleInPlaceActiveObject		= new CImpIOleInPlaceActiveObject(this, this->m_pUnkOuter);
	this->m_pImpISpecifyPropertyPages		= new CImpISpecifyPropertyPages(this, this->m_pUnkOuter);
	this->m_pImpIViewObject2				= new CImpIViewObject2(this, this->m_pUnkOuter);
	this->m_pImp_clsIMono					= new CImp_clsIMono(this, this->m_pUnkOuter);
	this->m_pImpISciMono					= new CImpISciMono(this, this->m_pUnkOuter);
	// SourceSelector manager
	this->m_pMySciUsbMono					= new CMySciUsbMono(this);
	// window manager
	this->m_pMyWindow						= new CMyWindow(this);

	if (NULL != this->m_pImpIDispatch					&&
		NULL != this->m_pImpIProvideClassInfo2			&&
		NULL != this->m_pImpIConnectionPointContainer	&&
		NULL != this->m_pImpIPersistStorage				&&
		NULL != this->m_pImpIPersistStreamInit			&&
		NULL != this->m_pImpIPersistPropertyBag			&&
		NULL != this->m_pImpIDataObject					&&
		NULL != this->m_pImpIOleObject					&&
		NULL != this->m_pImpIOleControl					&&
		NULL != this->m_pImpIOleInPlaceObject			&&
		NULL != this->m_pImpIOleInPlaceActiveObject		&&
		NULL != this->m_pImpISpecifyPropertyPages		&&
		NULL != this->m_pImpIViewObject2				&&
		NULL != this->m_pImp_clsIMono					&&
		NULL != this->m_pImpISciMono					&&
		NULL != this->m_pMySciUsbMono					&&
		NULL != this->m_pMyWindow)
	{
		// create the connection points
		hr = Utils_CreateConnectionPoint(this, 
			IID__SciUsbMono, &(this->m_paConnPts[CONN_PT_CUSTOMSINK]));
		if (SUCCEEDED(hr))
		{
			hr = Utils_CreateConnectionPoint(this,
				IID_IPropertyNotifySink, &(this->m_paConnPts[CONN_PT_PROPNOTIFY]));
		}
		if (SUCCEEDED(hr))
		{
			hr = this->Init__clsIMono();
			if (SUCCEEDED(hr))
			{
				hr = Utils_CreateConnectionPoint(this, this->m_iid__clsIMono,
					&(this->m_paConnPts[CONN_PT__clsIMono]));
			}
		}
		if (SUCCEEDED(hr))
		{
			hr = this->Init_SciMono();
			if (SUCCEEDED(hr))
			{
				hr = Utils_CreateConnectionPoint(this, this->m_iid__SciMono, &(this->m_paConnPts[CONN_PT__SciMono]));
			}
		}
	}
	else
	{
		hr = E_OUTOFMEMORY;
	}
	return hr;
}


HRESULT CMyObject::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;
	*ppTI		= NULL;
	hr = GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(CLSID_SciUsbMono, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

HRESULT CMyObject::GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo)
{
	HRESULT			hr;
	ITypeInfo	*	pClassInfo;
	BOOL			fSuccess	= FALSE;
	*ppTypeInfo	= NULL;
	hr = this->GetClassInfo(&pClassInfo);
	if (SUCCEEDED(hr))
	{
		fSuccess = Utils_FindImplClassName(pClassInfo, szInterface, ppTypeInfo);
		pClassInfo->Release();
	}
	return fSuccess ? S_OK : E_FAIL;
}

HRESULT CMyObject::Init__clsIMono()
{
	HRESULT				hr;
//	ITypeLib		*	pTypeLib;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->GetRefTypeInfo(TEXT("__clsIMono") , &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_iid__clsIMono	= pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, TEXT("GratingChanged"), &m_dispidGratingChanged);
		Utils_GetMemid(pTypeInfo, TEXT("BeforeMoveChange"), &m_dispidBeforeMoveChange);
		Utils_GetMemid(pTypeInfo, TEXT("HaveAborted"), &m_dispidHaveAborted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveCompleted"), &m_dispidMoveCompleted);
		Utils_GetMemid(pTypeInfo, TEXT("MoveError"), &m_dispidMoveError);
		Utils_GetMemid(pTypeInfo, TEXT("RequestChangeGrating"), &m_dispidRequestChangeGrating);
		Utils_GetMemid(pTypeInfo, TEXT("RequestParentWindow"), &m_dispidRequestParentWindow);
		Utils_GetMemid(pTypeInfo, TEXT("StatusMessage"), &m_dispidStatusMessage);
		Utils_GetMemid(pTypeInfo, TEXT("AmInitPropChanged"), &m_dispidAmInitPropChanged);
		Utils_GetMemid(pTypeInfo, TEXT("AutoGratingPropChanged"), &m_dispidAutoGratingPropChanged);
		pTypeInfo->Release();
	}
	return hr;
}

HRESULT CMyObject::Init_SciMono()
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;
	hr = this->GetRefTypeInfo(L"_SciMono", &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_iid__SciMono = pTypeAttr->guid;
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(pTypeInfo, L"PropChanged", &m_dispidPropChanged);
		Utils_GetMemid(pTypeInfo, L"Error", &m_dispidError);
		Utils_GetMemid(pTypeInfo, L"BusyStatusChange", &m_dispidBusyStatusChange);
		Utils_GetMemid(pTypeInfo, L"MonochromatorPropChanged", &m_dispidMonochromatorPropChanged);
		Utils_GetMemid(pTypeInfo, L"GratingPropChanged", &m_dispidGratingPropChanged);
		pTypeInfo->Release();
	}
	return hr;
}

HWND CMyObject::FireRequestMainWindow()
{
	VARIANTARG			varg;
	long				lVal		= 0;
	VariantInit(&varg);
	varg.vt		= VT_BYREF | VT_I4;
	varg.plVal	= &lVal;
//	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_RequestMainWindow, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidRequestParentWindow,
		&varg, 1);
	return (HWND) lVal;
}

void CMyObject::FireError(
									LPCTSTR			Error)
{
	VARIANTARG			varg;
	InitVariantFromString(Error, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_Error, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidMoveError, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__SciMono, this->m_dispidError, &varg, 1);
	// call the event handler if it exists
	if (NULL != this->m_pdispOnError)
	{
		Utils_InvokeMethod(this->m_pdispOnError, DISPID_VALUE, &varg, 1, NULL);
	}
	VariantClear(&varg);

}

void CMyObject::FireStatusMessage(
									LPCTSTR			Status,
									BOOL			fAmBusy)
{
	VARIANTARG			avarg[2];
	InitVariantFromString(Status, &avarg[1]);
	InitVariantFromBoolean(fAmBusy, &avarg[0]);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_StatusMessage, avarg, 2);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidStatusMessage, avarg, 2);
	// call the event handler if it exists
	if (NULL != this->m_pdispOnStatusMessage)
	{
		Utils_InvokeMethod(this->m_pdispOnStatusMessage, DISPID_VALUE, avarg, 2, NULL);
	}
	VariantClear(&avarg[1]);
}

void CMyObject::FireHaveNewPosition(
									double			position)
{
	VARIANTARG			varg;
	InitVariantFromDouble(position, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_HaveNewPosition, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidMoveCompleted, &varg, 1);
	// call the event handler if it exists
	if (NULL != this->m_pdispOnHaveNewPosition)
	{
		Utils_InvokeMethod(this->m_pdispOnHaveNewPosition, DISPID_VALUE, &varg, 1, NULL);
	}
}

BOOL CMyObject::FireRequestChangeGrating(
									long			grating)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	InitVariantFromInt32(grating, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fAllow;
//	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_RequestChangeGrating, avarg, 2);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidRequestChangeGrating,
		avarg, 2);
	return VARIANT_TRUE == fAllow;
}


void CMyObject::FireChangedGrating(
									long			grating)
{
	VARIANTARG			varg;
	InitVariantFromInt32(grating, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_ChangedGrating, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidGratingChanged, &varg, 1);
	// call the event handler
	if (NULL != this->m_pdispOnChangedGrating)
	{
		Utils_InvokeMethod(this->m_pdispOnChangedGrating, DISPID_VALUE, &varg, 1, NULL);
	}
}

void CMyObject::FireBusyStatusChange(
									BOOL			fBusy)
{
	VARIANTARG				varg;
	InitVariantFromBoolean(fBusy, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_BusyStatusChange, &varg, 1);
	Utils_NotifySinks(this, this->m_iid__SciMono, this->m_dispidBusyStatusChange, &varg, 1);
}

BOOL CMyObject::FireQueryAllowChangePosition(
									double			newPosition)
{
	VARIANTARG			avarg[2];
	VARIANT_BOOL		fAllow		= VARIANT_TRUE;
	InitVariantFromDouble(newPosition, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_BOOL;
	avarg[0].pboolVal	= &fAllow;
//	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_QueryAllowChangePosition, avarg, 2);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidBeforeMoveChange, avarg, 2);
	return VARIANT_TRUE == fAllow;
}

BOOL CMyObject::FireQueryAllowChangeZeroOffset()
{
	VARIANT_BOOL		fAllow	= VARIANT_TRUE;
	VARIANTARG			varg;
	VariantInit(&varg);
	varg.vt			= VT_BYREF | VT_BOOL;
	varg.pboolVal	= &fAllow;
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_QueryAllowChangeZeroOffset, &varg, 1);
	return VARIANT_TRUE == fAllow;
}

void CMyObject::FireRapidScanStepped(
									double			wavelength)
{
	VARIANTARG			varg;
	InitVariantFromDouble(wavelength, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_RapidScanStepped, &varg, 1);
}

void CMyObject::FireRapidScanEnded(
									BOOL			fSuccess)
{
	VARIANTARG			varg;
	InitVariantFromBoolean(fSuccess, &varg);
	Utils_NotifySinks(this, IID__SciUsbMono, DISPID_RapidScanEnded, &varg, 1);
}

void CMyObject::FireAmInitPropChanged(
									BOOL			amInit)
{
	VARIANTARG				varg;
	InitVariantFromBoolean(amInit, &varg);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidAmInitPropChanged, &varg, 1);
}

void CMyObject::FireAutoGratingPropChanged(
									BOOL			autoGrating)
{
	VARIANTARG				varg;
	InitVariantFromBoolean(autoGrating, &varg);
	Utils_NotifySinks(this, this->m_iid__clsIMono, this->m_dispidAutoGratingPropChanged, &varg, 1);
}

// _SciMono sink events
void CMyObject::FirePropChanged(LPCTSTR PropName)
{
	VARIANTARG			 varg;
	InitVariantFromString(PropName, &varg);
	Utils_NotifySinks(this, this->m_iid__SciMono, this->m_dispidPropChanged, &varg, 1);
	VariantClear(&varg);
}
void CMyObject::FireMonochromatorPropChanged(LPCTSTR propName)
{
	VARIANTARG			varg;
	InitVariantFromString(propName, &varg);
	Utils_NotifySinks(this, this->m_iid__SciMono, this->m_dispidMonochromatorPropChanged, &varg, 1);
	VariantClear(&varg);
}
void CMyObject::FireGratingPropChanged(long gratingID, LPCTSTR propName)
{
	VARIANTARG			avarg[2];
	InitVariantFromInt32(gratingID, &avarg[1]);
	InitVariantFromString(propName, &avarg[0]);
	Utils_NotifySinks(this, this->m_iid__SciMono, this->m_dispidGratingPropChanged, avarg, 2);
	VariantClear(&avarg[0]);
}



// successful initialization
void CMyObject::OnHaveInitialized()
{
	Utils_OnPropChanged(this, DISPID_AmOpen);
	if (NULL != this->m_pdispOnAmInitialized)
	{
		Utils_InvokeMethod(this->m_pdispOnAmInitialized, DISPID_VALUE, NULL, 0, NULL);
	}
}

// get the client site
HRESULT CMyObject::GetClientSite(
									IOleClientSite**	ppClientSite)
{
	HRESULT					hr;
	IOleObject			*	pOleObject;

	*ppClientSite	= NULL;
	hr = this->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
	if (SUCCEEDED(hr))
	{
		hr = pOleObject->GetClientSite(ppClientSite);
		pOleObject->Release();
	}
	return hr;
}

// get the extended control
HRESULT CMyObject::GetExtendedControl(
									IDispatch		**	ppdispExtended)
{
	HRESULT					hr;
	IOleControlSite		*	pControlSite;
	*ppdispExtended	= NULL;
	hr = this->GetControlSite(&pControlSite);
	if (SUCCEEDED(hr))
	{
		hr = pControlSite->GetExtendedControl(ppdispExtended);
		pControlSite->Release();
	}
	return hr;
}

// send on data change
void CMyObject::SendOnDataChange(
									BOOL				fDataOnStop)
{
	HRESULT					hr;
	IDataObject			*	pDataObject;
	if (NULL != this->m_pDataAdviseHolder)
	{
		hr = this->QueryInterface(IID_IDataObject, (LPVOID*) &pDataObject);
		if (SUCCEEDED(hr))
		{
			this->m_pDataAdviseHolder->SendOnDataChange(pDataObject, 0, 
				fDataOnStop ? ADVF_DATAONSTOP : 0);
			pDataObject->Release();
		}
	}
}

// get the container
HRESULT CMyObject::GetContainer(
									IOleContainer	**	ppContainer)
{
	HRESULT					hr;
	IOleClientSite		*	pClientSite;
	*ppContainer	= NULL;
	hr = this->GetClientSite(&pClientSite);
	if (SUCCEEDED(hr))
	{
		hr = pClientSite->GetContainer(ppContainer);
		pClientSite->Release();
	}
	return hr;
}

// get the InPlaceFrame object
HRESULT CMyObject::GetInPlaceFrame(
									IOleInPlaceFrame**	ppInPlaceFrame)
{
	HRESULT					hr;
	IOleContainer		*	pContainer;
	*ppInPlaceFrame		= NULL;
	hr = this->GetContainer(&pContainer);
	if (SUCCEEDED(hr))
	{
		hr = pContainer->QueryInterface(IID_IOleInPlaceFrame, (LPVOID*) ppInPlaceFrame);
		pContainer->Release();
	}
	return hr;
}

// enable or disable modeless dialogs in the client
void CMyObject::EnableModeless(
									BOOL				fEnable)
{
	HRESULT					hr;
	IOleInPlaceFrame	*	pInPlaceFrame;

	hr = this->GetInPlaceFrame(&pInPlaceFrame);
	if (SUCCEEDED(hr))
	{
		pInPlaceFrame->EnableModeless(fEnable);
		pInPlaceFrame->Release();
	}
}

// get the name for this control - given by the client
HRESULT CMyObject::GetName(
									LPTSTR			*	pszName)
{
	HRESULT					hr;
	IDispatch			*	pdispExtended;
	*pszName	= NULL;
	hr = this->GetExtendedControl(&pdispExtended);
	if (SUCCEEDED(hr))
	{
		hr = Utils_GetStringProperty(pdispExtended, DISPID_Name, pszName);
		pdispExtended->Release();
	}
	return hr;
}

// get our control window
HWND CMyObject::GetControlWindow()
{
	HRESULT				hr;
	IOleWindow		*	pOleWindow;
	HWND				hwndControl		= NULL;
	hr = this->QueryInterface(IID_IOleWindow, (LPVOID*) &pOleWindow);
	if (SUCCEEDED(hr))
	{
		pOleWindow->GetWindow(&hwndControl);
		pOleWindow->Release();
	}
	return hwndControl;
}

// set this as an active object
void CMyObject::SetActiveObject()
{
	HRESULT						hr;
	IOleInPlaceFrame		*	pInPlaceFrame;
	IOleInPlaceActiveObject	*	pActiveObject;
	hr = this->GetInPlaceFrame(&pInPlaceFrame);
	if (SUCCEEDED(hr))
	{
		hr = this->QueryInterface(IID_IOleInPlaceActiveObject, (LPVOID*) &pActiveObject);
		if (SUCCEEDED(hr))
		{
			pInPlaceFrame->SetActiveObject(pActiveObject, NULL);
			pActiveObject->Release();
		}
		pInPlaceFrame->Release();
	}
}

// get the control site
HRESULT CMyObject::GetControlSite(
									IOleControlSite**	ppControlSite)
{
	HRESULT						hr;
	IOleClientSite			*	pClientSite;
	*ppControlSite		= NULL;
	hr = this->GetClientSite(&pClientSite);
	if (SUCCEEDED(hr))
	{
		hr = pClientSite->QueryInterface(IID_IOleControlSite, (LPVOID*) ppControlSite);
		pClientSite->Release();
	}
	return hr;
}

// get the ambient object
HRESULT CMyObject::GetAmbientObject(
									IDispatch		**	ppdispAmbient)
{
	HRESULT					hr;
	IOleClientSite		*	pClientSite;
	*ppdispAmbient	= NULL;
	hr = this->GetClientSite(&pClientSite);
	if (SUCCEEDED(hr))
	{
		hr = pClientSite->QueryInterface(IID_IDispatch, (LPVOID*) ppdispAmbient);
		pClientSite->Release();
	}
	return hr;
}

// get an ambient property
HRESULT CMyObject::GetAmbientProperty(
									DISPID				dispid,
									VARIANT			*	pValue)
{
	HRESULT					hr;
	IDispatch			*	pdispAmbient;
	hr = this->GetAmbientObject(&pdispAmbient);
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokePropertyGet(pdispAmbient, dispid, NULL, 0, pValue);
		pdispAmbient->Release();
	}
	return hr;
}


CMyObject::CImpIDispatch::CImpIDispatch(CMyObject * pMyObject, IUnknown * punkOuter) :
	m_pMyObject(pMyObject),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL)
{
}

CMyObject::CImpIDispatch::~CImpIDispatch()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDispatch::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDispatch::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT					hr;
	ITypeLib			*	pTypeLib;

	*ppTInfo	= NULL;
	hr = GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(IID_ISciUsbMono, ppTInfo);
		pTypeLib->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;

	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDispatch::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	HRESULT					hr;
	VARIANTARG				varg;
	VARIANT					varResult;
	UINT					uArgErr;
	LPTSTR					szString			= NULL;
	LPOLESTR				pOleStr				= NULL;
	// Make sure the wFlags are valid.
	if(wFlags & ~(DISPATCH_METHOD | DISPATCH_PROPERTYGET |
		DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF))
		return E_INVALIDARG;
	if (NULL == pVarResult) pVarResult = &varResult;
	if (NULL == puArgErr) puArgErr = &uArgErr;
	VariantInit(&varg);
	VariantInit(pVarResult);
	switch (dispIdMember)
	{
	case DISPID_position:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_R8;
			pVarResult->dblVal	= this->m_pMyObject->m_pMySciUsbMono->Getposition();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->SetPosition(varg.dblVal);
		}
		break;
	case DISPID_currentGrating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetcurrentGrating();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetcurrentGrating(varg.lVal);
		}
		break;
	case DISPID_AmOpen:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->GetAmOpen() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetAmOpen(VARIANT_TRUE == varg.boolVal);
		}
		break;
	case DISPID_NumberOfGratings:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetNumberOfGratings();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetNumberOfGratings(varg.lVal);
		}
		break;
	case DISPID_ConfigFile:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			this->m_pMyObject->m_pMySciUsbMono->GetConfigFile(&szString);
			if (NULL != szString)
			{
	//			Utils_AnsiToUnicode(szString, &pOleStr);
		//		CoTaskMemFree((LPVOID) szString);
				pVarResult->vt		= VT_BSTR;
				pVarResult->bstrVal	= SysAllocString(szString);
				CoTaskMemFree((LPVOID) szString);
			}
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, puArgErr);
			if (FAILED(hr)) return hr;
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
	//		VariantClear(&varg);
			this->m_pMyObject->m_pMySciUsbMono->SetConfigFile(varg.bstrVal);
			VariantClear(&varg);
	//		CoTaskMemFree((LPVOID) szString);
		}
		break;
	case DISPID_autoGrating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->GetautoGrating() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetautoGrating(VARIANT_TRUE == varg.boolVal);
		}
		break;
	case DISPID_INIFile:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			this->m_pMyObject->m_pMySciUsbMono->GetINIFile(&szString);
			if (NULL != szString)
			{
	//			Utils_AnsiToUnicode(szString, &pOleStr);
	//			CoTaskMemFree((LPVOID) szString);
				pVarResult->vt		= VT_BSTR;
				pVarResult->bstrVal	= SysAllocString(szString);
				CoTaskMemFree((LPVOID) szString);
			}
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, puArgErr);
			if (FAILED(hr)) return hr;
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
	//		VariantClear(&varg);
			this->m_pMyObject->m_pMySciUsbMono->SetINIFile(varg.bstrVal);
			VariantClear(&varg);
	//		CoTaskMemFree((LPVOID) szString);
		}
		break;
	case DISPID_AmBusy:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->GetAmBusy() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		break;
	case DISPID_DisplayName:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			this->m_pMyObject->m_pMySciUsbMono->GetDisplayName(&szString);
			if (NULL != szString)
			{
		//		Utils_AnsiToUnicode(szString, &pOleStr);
		//		CoTaskMemFree((LPVOID*) szString);
				pVarResult->vt		= VT_BSTR;
				pVarResult->bstrVal	= SysAllocString(szString);
				CoTaskMemFree((LPVOID) szString);
			}
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, puArgErr);
			if (FAILED(hr)) return hr;
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
			SHStrDup(varg.bstrVal, &szString);
			VariantClear(&varg);
			this->m_pMyObject->m_pMySciUsbMono->SetDisplayName(szString);
			CoTaskMemFree((LPVOID) szString);
			CNamedObjects	*	pNamedObjects	= GetNamedObjects();
			if (NULL != pNamedObjects)
				pNamedObjects->AddNamedObject(this->m_pMyObject, TRUE);
		}
		break;
	case DISPID_WaitForComplete:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->GetWaitForComplete() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetWaitForComplete(VARIANT_TRUE == varg.boolVal);
		}
		break;
	case DISPID_AllowChangeZeroOffset:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->FireQueryAllowChangeZeroOffset() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		break;
	// *************************************
	// message handlers
	// start
	// *************************************
	case DISPID_OnError:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			this->GetHandler(this->m_pMyObject->m_pdispOnError, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT)	||
				 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetHandler(pDispParams, &(this->m_pMyObject->m_pdispOnError));
		break;
	case DISPID_OnStatusMessage:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			this->GetHandler(this->m_pMyObject->m_pdispOnStatusMessage, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT)	||
				 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetHandler(pDispParams, &(this->m_pMyObject->m_pdispOnStatusMessage));
		break;
	case DISPID_OnHaveNewPosition:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			this->GetHandler(this->m_pMyObject->m_pdispOnHaveNewPosition, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT)	||
				 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetHandler(pDispParams, &(this->m_pMyObject->m_pdispOnHaveNewPosition));
		break;
	case DISPID_OnChangedGrating:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			this->GetHandler(this->m_pMyObject->m_pdispOnChangedGrating, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT)	||
				 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetHandler(pDispParams, &(this->m_pMyObject->m_pdispOnChangedGrating));
		break;
	case DISPID_OnAmInitialized:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			this->GetHandler(this->m_pMyObject->m_pdispOnAmInitialized, pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT)	||
				 0 != (wFlags & DISPATCH_PROPERTYPUTREF))
			return this->SetHandler(pDispParams, &(this->m_pMyObject->m_pdispOnAmInitialized));
		break;
	// *************************************
	// message handlers
	// end
	// *************************************
	case DISPID_SetupWindow:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			if (NULL == pVarResult) return E_INVALIDARG;
			InitVariantFromInt32((long) this->m_pMyObject->m_pMySciUsbMono->GetSetupWindow(),
				pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetSetupWindow((HWND) varg.lVal);
		}
		break;
	case DISPID_MonoInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
	//		if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
			this->m_pMyObject->m_pMySciUsbMono->GetMonoInfo(varg.iVal, pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			short int param;
			if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	//		if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			param = varg.iVal;
			VariantCopy(&varg, &(pDispParams->rgvarg[0]));
			this->m_pMyObject->m_pMySciUsbMono->SetMonoInfo(param, &varg);
			VariantClear(&varg);
		}
		break;
	case DISPID_GratingInfo:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			short int iGrating;
			short int param;
			if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			param = varg.iVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			iGrating = varg.iVal;
//
//			if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
//			if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
			this->m_pMyObject->m_pMySciUsbMono->GetGratingInfo(
				iGrating, param, pVarResult);
//				*(pDispParams->rgvarg[1].piVal), *(pDispParams->rgvarg[0].piVal), pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
			short int iGrating;
			short int param;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			param = varg.iVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[2]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			iGrating = varg.iVal;

//			if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[2].vt) return DISP_E_TYPEMISMATCH;
//			if ((VT_BYREF | VT_I2) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
			VariantCopy(&varg, &(pDispParams->rgvarg[0]));
			this->m_pMyObject->m_pMySciUsbMono->SetGratingInfo(iGrating, param, &varg);
//				*(pDispParams->rgvarg[2].piVal), *(pDispParams->rgvarg[1].piVal), &varg);
			VariantClear(&varg);
		}
		break;
	case DISPID_moveTime:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			hr = DispGetParam(pDispParams, 0, VT_R8, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			pVarResult->vt		= VT_R8;
			pVarResult->dblVal	= this->m_pMyObject->m_pMySciUsbMono->GetmoveTime(varg.dblVal);
		}
		break;
	case DISPID_Setup:
		this->m_pMyObject->m_pMySciUsbMono->Setup();
		break;
	case DISPID_SetGratingParams:
		pVarResult->vt		= VT_BOOL;
		pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->SetGratingParams() ? VARIANT_TRUE : VARIANT_FALSE;
		break;
	case DISPID_IsValidPosition:
		hr = DispGetParam(pDispParams, 0, VT_R8, &varg, puArgErr);
		if (FAILED(hr)) return hr;
		pVarResult->vt		= VT_BOOL;
		pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->IsValidPosition(varg.dblVal) ? VARIANT_TRUE : VARIANT_FALSE;
		break;
	case DISPID_ConvertStepsToNM:
		if (4 == pDispParams->cArgs)
		{
			BOOL			fToNM;
			short int		grating;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[3]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_BOOL);
			if (FAILED(hr)) return hr;
			fToNM	= VARIANT_TRUE == varg.boolVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[2]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			grating = varg.iVal;
			if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
			if ((VT_BYREF | VT_I4) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->ConvertStepsToNM(
				fToNM, grating, pDispParams->rgvarg[1].pdblVal, pDispParams->rgvarg[0].plVal) ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else return DISP_E_BADPARAMCOUNT;
		break;
	case DISPID_Abort:
		this->m_pMyObject->m_pMySciUsbMono->Abort();
		break;
	case DISPID_WriteConfig:
		if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
		hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
		if (FAILED(hr)) return hr;
		hr = VariantChangeType(&varg, &varg, 0, VT_BSTR);
		if (FAILED(hr)) return hr;
		SHStrDup(varg.bstrVal, &szString);
		VariantClear(&varg);
		this->m_pMyObject->m_pMySciUsbMono->WriteConfig(szString);
		CoTaskMemFree((LPVOID) szString);
		break;
	case DISPID_WriteINI:
		if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
		hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
		if (FAILED(hr)) return hr;
		hr = VariantChangeType(&varg, &varg, 0, VT_BSTR);
		if (FAILED(hr)) return hr;
		SHStrDup(varg.bstrVal, &szString);
		VariantClear(&varg);
		this->m_pMyObject->m_pMySciUsbMono->WriteINI(szString);
		CoTaskMemFree((LPVOID) szString);
		break;
	case DISPID_GetCounter:
		pVarResult->vt		= VT_BSTR;
		pVarResult->dblVal	= this->m_pMyObject->m_pMySciUsbMono->GetCounter();
		break;
	case DISPID_setInitialPositions:
		if (3 == pDispParams->cArgs)
		{
			long			currentGrating;
			BOOL			autoGrating;
			double			CurrentPos;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[2]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I4);
			if (FAILED(hr)) return hr;
			currentGrating = varg.lVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_BOOL);
			if (FAILED(hr)) return hr;
			autoGrating = VARIANT_TRUE == varg.boolVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_R8);
			if (FAILED(hr)) return hr;
			CurrentPos = varg.dblVal;
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->setInitialPositions(
				currentGrating, autoGrating, CurrentPos) ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else return DISP_E_BADPARAMCOUNT;
		break;
	case DISPID_SaveGratingZeroPosition:
		if (2 == pDispParams->cArgs)
		{
			short int	Grating;
			long		zeroPos;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[1]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I2);
			if (FAILED(hr)) return hr;
			Grating = varg.iVal;
			hr = VariantCopyInd(&varg, &(pDispParams->rgvarg[0]));
			if (FAILED(hr)) return hr;
			hr = VariantChangeType(&varg, &varg, 0, VT_I4);
			if (FAILED(hr)) return hr;
			zeroPos = varg.lVal;
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->SaveGratingZeroPosition(
				Grating, zeroPos) ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else return DISP_E_BADPARAMCOUNT;
		break;
	case DISPID_doInit:
		this->m_pMyObject->m_pMySciUsbMono->doInit();
		break;
	case DISPID_CanAutoSelect:
		pVarResult->vt		= VT_BOOL;
		pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->CanAutoSelect() ? VARIANT_TRUE : VARIANT_FALSE;
		break;
	// specific for Arcus Motors
	case DISPID_RunCurrent:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetRunCurrent();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetRunCurrent(varg.lVal);
		}
		break;
	case DISPID_IdleCurrent:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetIdleCurrent();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetIdleCurrent(varg.lVal);
		}
		break;
	case DISPID_GratingZeroOffset:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetGratingZeroOffset(varg.lVal);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			long			grating;
			long			zeroOffset;
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			grating = varg.lVal;
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			zeroOffset = varg.lVal;
			this->m_pMyObject->m_pMySciUsbMono->SetGratingZeroOffset(grating, zeroOffset);
		}
		break;
	case DISPID_GratingStepsPerNM:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			pVarResult->vt		= VT_R8;
			pVarResult->dblVal	= this->m_pMyObject->m_pMySciUsbMono->GetGratingStepsPerNM(varg.lVal);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			long			grating;
			double			stepsPerNM;
			hr = DispGetParam(pDispParams, 0, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			grating = varg.lVal;
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			stepsPerNM = varg.dblVal;
			this->m_pMyObject->m_pMySciUsbMono->SetGratingStepsPerNM(grating, stepsPerNM);
		}
		break;
	case DISPID_SetMotorControl:
		// not used but must exist
		break;
	case DISPID_ApplyBacklashCorrection:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_BOOL;
			pVarResult->boolVal	= this->m_pMyObject->m_pMySciUsbMono->GetApplyBacklashCorrection() ? VARIANT_TRUE : VARIANT_FALSE;
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetApplyBacklashCorrection(VARIANT_TRUE == varg.boolVal);
		}
		break;
	case DISPID_HighSpeed:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			pVarResult->vt		= VT_I4;
			pVarResult->lVal	= this->m_pMyObject->m_pMySciUsbMono->GetHighSpeed();
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, puArgErr);
			if (FAILED(hr)) return hr;
			this->m_pMyObject->m_pMySciUsbMono->SetHighSpeed(varg.lVal);
			Utils_OnPropChanged(this->m_pMyObject, DISPID_HighSpeed);
		}
		break;
	case DISPID_ConvertSteps:
		return this->ConvertSteps(pDispParams, pVarResult);
	case DISPID_ConvertNM:
		return this->ConvertNM(pDispParams, pVarResult);
	case DISPID_DisplayConfigValues:
		{
			CDlgDisplayConfigValues		dlg;
			dlg.SetOurObject(this);
			dlg.DoOpenDialog(this->m_pMyObject->FireRequestMainWindow());
		}
		return S_OK;
	/*
		Rapid scan items
	*/
	case DISPID_RapidScanRunning:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetRapidScanRunning(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetRapidScanRunning(pDispParams);
		else
			return DISP_E_MEMBERNOTFOUND;
	case DISPID_StartRapidScan:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->StartRapidScan(pDispParams, pVarResult);
 		else
			return DISP_E_MEMBERNOTFOUND;
	case DISPID_PageSelected:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetPageSelected(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetPageSelected(pDispParams);
		else
			return DISP_E_MEMBERNOTFOUND;
	case DISPID_RemoveNamedObject:
		{
			CNamedObjects	*	pNamedObjects	= GetNamedObjects();
			if (NULL != pNamedObjects)
				pNamedObjects->AddNamedObject(this->m_pMyObject, FALSE);
		}
		break;
	case DISPID_BacklashSteps:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetBacklashSteps(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetBacklashSteps(pDispParams);
		return DISP_E_MEMBERNOTFOUND;
	case DISPID_ReInitOnScanStart:
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetReInitOnScanStart(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetReInitOnScanStart(pDispParams);
		return DISP_E_MEMBERNOTFOUND;
	default:
		return DISP_E_MEMBERNOTFOUND;
	}
	return S_OK;
}


HRESULT CMyObject::CImpIDispatch::GetReInitOnScanStart(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMySciUsbMono->GetReInitOnScanStart(), pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetReInitOnScanStart(
	DISPPARAMS	*	pDispParams)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetReInitOnScanStart(VARIANT_TRUE == varg.boolVal);
	Utils_OnPropChanged(this->m_pMyObject, DISPID_ReInitOnScanStart);
	return S_OK;
}


// set position using backlash correction if needed
void CMyObject::CImpIDispatch::SetPosition(
									double			position)
{
//	double			currentPosition;
	double			setPosition;
//	long			backlashSteps = this->m_pMyObject->m_pMySciUsbMono->GetBacklashSteps();
//	long			steps;
//	long			grating;
//	// check the apply backlash correction flag
/*
	if (backlashSteps > 0 && this->m_pMyObject->m_pMySciUsbMono->GetApplyBacklashCorrection())
	{
		// check the current position
		currentPosition = this->m_pMyObject->m_pMySciUsbMono->Getposition();
		// apply backlash correction if the desired position is less than this
		if (position < currentPosition)
		{
			this->m_pMyObject->m_pMySciUsbMono->FindGrating(position, &grating);
			// find the steps for this desired position
			this->m_pMyObject->m_pMySciUsbMono->ConvertStepsToNM(FALSE,
				grating, &position, &steps);
			// subtract off the backlash steps
			steps -= backlashSteps;
			// find the new position
			this->m_pMyObject->m_pMySciUsbMono->ConvertStepsToNM(TRUE,
				grating, &setPosition, &steps);
			this->m_pMyObject->m_pMySciUsbMono->Setposition(setPosition);
		}
	}
*/
	// move to the new position
	this->m_pMyObject->m_pMySciUsbMono->Setposition(position);
	// property change notification
	Utils_OnPropChanged(this->m_pMyObject, DISPID_position);
}


void CMyObject::CImpIDispatch::GetHandler(
									IDispatch	*	pdispHandler,
									VARIANT		*	pVarResult)
{
	if (NULL != pdispHandler)
	{
		pVarResult->vt			= VT_DISPATCH;
		pVarResult->pdispVal	= pdispHandler;
		pVarResult->pdispVal->AddRef();
	}
}

HRESULT CMyObject::CImpIDispatch::SetHandler(
									DISPPARAMS	*	pDispParams,
									IDispatch	**	ppdispHandler)
{
	if (1 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	Utils_RELEASE_INTERFACE(*ppdispHandler);
	if (VT_DISPATCH == pDispParams->rgvarg[0].vt && NULL != pDispParams->rgvarg[0].pdispVal)
	{
		*ppdispHandler = pDispParams->rgvarg[0].pdispVal;
		pDispParams->rgvarg[0].pdispVal->AddRef();
	}
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ConvertSteps(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	long			Grating;
	long			Steps;
	double			position		= -100.0;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Grating = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Steps = varg.lVal;
	this->m_pMyObject->m_pMySciUsbMono->ConvertStepsToNM(TRUE, (short) Grating, &position, &Steps);
	InitVariantFromDouble(position, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::ConvertNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				Grating;
	double				nm;
	long				Steps	= 0;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	Grating = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	nm = varg.dblVal;
	this->m_pMyObject->m_pMySciUsbMono->ConvertStepsToNM(FALSE, (short) Grating, &nm, &Steps);
	InitVariantFromInt32(Steps, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::GetRapidScanRunning(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMySciUsbMono->GetRapidScanRunning(), pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpIDispatch::SetRapidScanRunning(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetRapidScanRunning(VARIANT_TRUE == varg.boolVal);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::StartRapidScan(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT			hr;
	VARIANTARG		varg;
	UINT			uArgErr;
	long			grating;
	double			startWavelength;
	double			endWavelength;
	double			NMPerSecond;
	BOOL			fSuccess		= FALSE;

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	grating = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	startWavelength = varg.dblVal;
	hr = DispGetParam(pDispParams, 2, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	endWavelength = varg.dblVal;
	hr = DispGetParam(pDispParams, 3, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	NMPerSecond = varg.dblVal;
	fSuccess = this->m_pMyObject->m_pMySciUsbMono->StartRapidScan(grating, startWavelength,
		endWavelength, NMPerSecond);
	if (NULL != pVarResult) InitVariantFromBoolean(fSuccess, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetPageSelected(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_fPageSelected, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::SetPageSelected(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_fPageSelected = VARIANT_TRUE == varg.boolVal;
	return S_OK;
}

HRESULT CMyObject::CImpIDispatch::GetBacklashSteps(
	VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pMyObject->m_pMySciUsbMono->GetBacklashSteps(),
		pVarResult);
	return S_OK;
}
HRESULT CMyObject::CImpIDispatch::SetBacklashSteps(
	DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetBacklashSteps(varg.lVal);
	Utils_OnPropChanged(this->m_pMyObject, DISPID_BacklashSteps);
	return S_OK;
}



CMyObject::CImpIConnectionPointContainer::CImpIConnectionPointContainer(CMyObject * pMyObject, IUnknown * punkOuter) :
	m_pBackObj(pMyObject),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIConnectionPointContainer::~CImpIConnectionPointContainer()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIConnectionPointContainer::Release()
{
	return this->m_punkOuter->Release();
}

// IConnectionPointContainer methods
STDMETHODIMP CMyObject::CImpIConnectionPointContainer::EnumConnectionPoints( 
									IEnumConnectionPoints** ppEnum)
{
	return Utils_CreateEnumConnectionPoints(this, MAX_CONN_PTS, this->m_pBackObj->m_paConnPts,
		ppEnum);
}

STDMETHODIMP CMyObject::CImpIConnectionPointContainer::FindConnectionPoint(
									REFIID				riid,
									IConnectionPoint ** ppCP)
{
	IConnectionPoint*	pConnPt	= NULL;
	HRESULT				hr		= CONNECT_E_NOCONNECTION;

	if (NULL == ppCP) return E_POINTER;
	*ppCP	= NULL;
	if (IID__SciUsbMono == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_CUSTOMSINK];
	else if (IID_IPropertyNotifySink == riid)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT_PROPNOTIFY];
	else if (riid == this->m_pBackObj->m_iid__clsIMono)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT__clsIMono];
	else if (riid == this->m_pBackObj->m_iid__SciMono)
		pConnPt = this->m_pBackObj->m_paConnPts[CONN_PT__SciMono];
	if (NULL != pConnPt)
	{
		*ppCP		= pConnPt;
		pConnPt->AddRef();
		hr			= S_OK;
	}
	return hr;
}

CMyObject::CImpIProvideClassInfo2::CImpIProvideClassInfo2(CMyObject * pMyObject, IUnknown * punkOuter) :
	m_pBackObj(pMyObject),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIProvideClassInfo2::~CImpIProvideClassInfo2()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIProvideClassInfo2::Release()
{
	return this->m_punkOuter->Release();
}

// IProvideClassInfo method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetClassInfo(
									ITypeInfo	**	ppTI)
{
	HRESULT				hr;
	ITypeLib		*	pTypeLib;
	if (NULL == ppTI) return E_POINTER;
	*ppTI		= NULL;
	hr = GetTypeLib(&pTypeLib);
	if (SUCCEEDED(hr))
	{
		hr = pTypeLib->GetTypeInfoOfGuid(CLSID_SciUsbMono, ppTI);
		pTypeLib->Release();
	}
	return hr;
}

// IProvideClassInfo2 method
STDMETHODIMP CMyObject::CImpIProvideClassInfo2::GetGUID(
									DWORD			dwGuidKind,
									GUID		*	pGUID)
{
	if (NULL == pGUID) return E_POINTER;
	if (GUIDKIND_DEFAULT_SOURCE_DISP_IID == dwGuidKind)
	{
		*pGUID		= IID__SciUsbMono;
		return S_OK;
	}
	else
	{
		*pGUID		= GUID_NULL;
		return E_INVALIDARG;
	}
}

CMyObject::CImpIPersistStorage::CImpIPersistStorage(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_fNoScribble(FALSE)		// no scribble mode
{
}

CMyObject::CImpIPersistStorage::~CImpIPersistStorage()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistStorage::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStorage::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStorage::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistStorage::GetClassID(
									CLSID			*	pClassID)
{
	*pClassID	= CLSID_SciUsbMono;
	return S_OK;
}

// IPersistStorage methods
STDMETHODIMP CMyObject::CImpIPersistStorage::IsDirty(void)
{
	return this->m_pBackObj->m_pMySciUsbMono->GetDirty() ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIPersistStorage::InitNew(
									IStorage		*	pStg)
{
	HRESULT					hr;
	IPersistStreamInit	*	pPersistStream;					// our persist stream interface
	IStream				*	pStm		= NULL;				// new stream

	if (this->m_pBackObj->m_fAmLoaded) return CO_E_ALREADYINITIALIZED;
	hr = this->m_pBackObj->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStream);
	if (SUCCEEDED(hr))
	{
		// initialize the persist stream object
		hr = pPersistStream->InitNew();
		pPersistStream->Release();
	}
	// create a new stream in the storage object
	hr = this->CreateNewStream(pStg, &pStm);
	if (SUCCEEDED(hr))
	{
		// save without clearing the dirty flag
		hr = this->SaveToStream(pStm, FALSE);
		// if succeeded, store references to the stream and the storage objects
		if (SUCCEEDED(hr))
		{
			this->m_pBackObj->m_pStm	= pStm;
			this->m_pBackObj->m_pStm->AddRef();
			this->m_pBackObj->m_pStg	= pStg;
			this->m_pBackObj->m_pStg->AddRef();
			// set the have loaded flag
			this->m_pBackObj->m_fAmLoaded	= TRUE;
		}
		pStm->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistStorage::Load(
									IStorage		*	pStg)
{
	HRESULT					hr;
	IPersistStreamInit	*	pPersistStream;					// our persist stream interface
//	ULARGE_INTEGER			uli;
	LPOLESTR				StreamName;						// stream name
	CLSID					clsid;							// our class id
	IStream				*	pStm		= NULL;				// new stream
	LARGE_INTEGER			liMove;							// stream positioning

	if (this->m_pBackObj->m_fAmLoaded) return CO_E_ALREADYINITIALIZED;
	// open the stream in the storage object
	StreamName = NULL;
	SHStrDup(STREAM_NAME, &StreamName);
	hr = pStg->OpenStream(StreamName, NULL, STGM_READWRITE, 0, &pStm);
	CoTaskMemFree((LPVOID) StreamName);
	if (SUCCEEDED(hr))
	{
		// get our persist stream interface
		hr = this->m_pBackObj->QueryInterface(IID_IPersistStreamInit,
			(LPVOID*) &pPersistStream);
		if (SUCCEEDED(hr))
		{
			// read the stream
			// first the class id
			hr = ReadClassStm(pStm, &clsid);
			if (SUCCEEDED(hr) && CLSID_SciUsbMono == clsid)
			{
				hr = pPersistStream->Load(pStm);
			}
			else hr = E_FAIL;
			if (SUCCEEDED(hr))
			{
				// reset the stream pointer to its start position
				liMove.QuadPart	= 0;
				hr = pStm->Seek(liMove, STREAM_SEEK_SET, 0);
			}
			pPersistStream->Release();
		}
		// if succeeded, store references to the stream and the storage objects
		if (SUCCEEDED(hr))
		{
			this->m_pBackObj->m_pStm	= pStm;
			this->m_pBackObj->m_pStm->AddRef();
			this->m_pBackObj->m_pStg	= pStg;
			this->m_pBackObj->m_pStg->AddRef();
			// set the have loaded flag
			this->m_pBackObj->m_fAmLoaded	= TRUE;
		}
		pStm->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistStorage::Save(
									IStorage		*	pStgSave,   
									BOOL				fSameAsLoad)
{
	HRESULT				hr;
	IStream			*	pStm;

	if (fSameAsLoad)
	{
		// save case
		if (NULL != this->m_pBackObj->m_pStm)
		{
			hr = this->SaveToStream(this->m_pBackObj->m_pStm, TRUE);
		}
		else hr = E_UNEXPECTED;
	}
	else
	{
		// save or save a copy of
		// create a new stream
		hr = this->CreateNewStream(pStgSave, &pStm);
		if (SUCCEEDED(hr))
		{
			hr = this->SaveToStream(pStm, FALSE);
			pStm->Release();						// release the new stream
		}
	}
	if (SUCCEEDED(hr)) this->m_fNoScribble = TRUE;
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistStorage::SaveCompleted(
									IStorage		*	pStgNew)
{
	HRESULT				hr		= S_OK;
	if (NULL == this->m_pBackObj->m_pStg)
	{
		// object is in hands off mode
		if (NULL != pStgNew)
		{
			hr = this->OpenStream(pStgNew, &(this->m_pBackObj->m_pStm));
			if (SUCCEEDED(hr))
			{
				this->m_pBackObj->m_pStg = pStgNew;
				this->m_pBackObj->m_pStg->AddRef();
			}
		}
		else hr = E_INVALIDARG;
	}
	else
	{
		if (NULL != pStgNew)
		{
			// set in hands off mode
			this->HandsOffStorage();			// close our current storage
			hr = this->OpenStream(pStgNew, &(this->m_pBackObj->m_pStm));
			if (SUCCEEDED(hr))
			{
				this->m_pBackObj->m_pStg = pStgNew;
				this->m_pBackObj->m_pStg->AddRef();
			}
		}
	}
	// clear the no scribble flag
	if (SUCCEEDED(hr))
		this->m_fNoScribble = FALSE;
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistStorage::HandsOffStorage(void)
{
	// release the storage objects
	Utils_RELEASE_INTERFACE(this->m_pBackObj->m_pStg);
	Utils_RELEASE_INTERFACE(this->m_pBackObj->m_pStm);
	return S_OK;
}

HRESULT	CMyObject::CImpIPersistStorage::SaveToStream(
									IStream			*	pStm,
									BOOL				fClearDirty)
{
	HRESULT					hr;
	IPersistStreamInit	*	pPersistStream;
	CLSID					clsid;
	LARGE_INTEGER			liMove;

	hr = this->m_pBackObj->QueryInterface(IID_IPersistStreamInit, 
		(LPVOID*) &pPersistStream);
	if (SUCCEEDED(hr))
	{
		// write the class id
		this->GetClassID(&clsid);
		hr = WriteClassStm(pStm, clsid);
		// save to the stream, without clearing the dirty flag
		if (SUCCEEDED(hr)) hr = pPersistStream->Save(pStm, fClearDirty);
		if (SUCCEEDED(hr))
		{
			// reset the stream pointer to its start position
			liMove.QuadPart	= 0;
			hr = pStm->Seek(liMove, STREAM_SEEK_SET, 0);
		}
		pPersistStream->Release();
	}
	return hr;
}

HRESULT	CMyObject::CImpIPersistStorage::CreateNewStream(
									IStorage		*	pStg,
									IStream			**	ppStm)
{
	HRESULT					hr;
	LPOLESTR				StreamName;
	IPersistStreamInit	*	pPersistStream;
	ULARGE_INTEGER			uli;

	*ppStm = NULL;
	// create a new stream in the storage object
	StreamName = NULL;
	SHStrDup(STREAM_NAME, &StreamName);
	hr = pStg->CreateStream(StreamName, STGM_READWRITE | STGM_CREATE,
		0, 0, ppStm);
	CoTaskMemFree((LPVOID) StreamName);
	if (SUCCEEDED(hr))
	{
		// initialize the object
		hr = this->m_pBackObj->QueryInterface(IID_IPersistStreamInit, 
			(LPVOID*) &pPersistStream);
		if (SUCCEEDED(hr))
		{
			// set the stream size
			pPersistStream->GetSizeMax(&uli);
			hr = (*ppStm)->SetSize(uli);
			pPersistStream->Release();
		}
	}
	return hr;
}

HRESULT	CMyObject::CImpIPersistStorage::OpenStream(
									IStorage		*	pStg,
									IStream			**	ppStm)
{
	HRESULT				hr;
	LPOLESTR			StreamName;

	*ppStm = NULL;
	// open the stream in the storage object
	StreamName = NULL;
	SHStrDup(STREAM_NAME, &StreamName);
	hr = pStg->OpenStream(StreamName, NULL, STGM_READWRITE, 0, ppStm);
	CoTaskMemFree((LPVOID) StreamName);
	return hr;
}

CMyObject::CImpIPersistStreamInit::CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIPersistStreamInit::~CImpIPersistStreamInit()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistStreamInit::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetClassID(
									CLSID			*	pClassID)
{
	*pClassID	= CLSID_SciUsbMono;
	return S_OK;
}

// IPersistStreamInit methods
STDMETHODIMP CMyObject::CImpIPersistStreamInit::IsDirty(void)
{
	return this->m_pBackObj->m_pMySciUsbMono->GetDirty() ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Load(
									LPSTREAM			pStm)
{
	HRESULT					hr;
	if (this->m_pBackObj->m_fAmLoaded) return E_UNEXPECTED;
	hr = this->m_pBackObj->m_pMySciUsbMono->Load(pStm);
	if (SUCCEEDED(hr)) this->m_pBackObj->m_fAmLoaded = TRUE;
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::Save(
									LPSTREAM			pStm ,  
									BOOL				fClearDirty)
{
	return this->m_pBackObj->m_pMySciUsbMono->Save(pStm, fClearDirty);
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::GetSizeMax(
									ULARGE_INTEGER	*	pcbSize)
{
	pcbSize->HighPart		= this->m_pBackObj->m_pMySciUsbMono->GetSaveSize();
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIPersistStreamInit::InitNew(void)
{
	HRESULT				hr;
	if (this->m_pBackObj->m_fAmLoaded) return E_UNEXPECTED;
	hr = this->m_pBackObj->m_pMySciUsbMono->InitNew();
	if (SUCCEEDED(hr)) this->m_pBackObj->m_fAmLoaded = TRUE;
	return hr;
}

CMyObject::CImpIPersistPropertyBag::CImpIPersistPropertyBag(CMyObject * pMyObject, IUnknown * punkOuter) :
	m_pMyObject(pMyObject),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpIPersistPropertyBag::~CImpIPersistPropertyBag()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIPersistPropertyBag::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistPropertyBag::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIPersistPropertyBag::Release()
{
	return this->m_punkOuter->Release();
}

// IPersist method
STDMETHODIMP CMyObject::CImpIPersistPropertyBag::GetClassID(
									CLSID			*	pClassID)
{
	*pClassID	= CLSID_SciUsbMono;
	return S_OK;
}

// IPersistPropertyBag methods
STDMETHODIMP CMyObject::CImpIPersistPropertyBag::InitNew()
{
	HRESULT				hr;
	if (this->m_pMyObject->m_fAmLoaded) return E_UNEXPECTED;
	hr = this->m_pMyObject->m_pMySciUsbMono->InitNew();
	if (SUCCEEDED(hr)) this->m_pMyObject->m_fAmLoaded = TRUE;
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistPropertyBag::Load(          
									IPropertyBag	*	pPropBag,
									IErrorLog		*	pErrorLog)
{
	HRESULT				hr;
	if (this->m_pMyObject->m_fAmLoaded) return E_UNEXPECTED;
	hr = this->m_pMyObject->m_pMySciUsbMono->Load(pPropBag, pErrorLog);
	if (SUCCEEDED(hr)) this->m_pMyObject->m_fAmLoaded = TRUE;
	return hr;
}

STDMETHODIMP CMyObject::CImpIPersistPropertyBag::Save(
									IPropertyBag	*	pPropBag,
									BOOL				fClearDirty,
									BOOL				fSaveAllProperties)
{
	return this->m_pMyObject->m_pMySciUsbMono->Save(pPropBag, fClearDirty,
		fSaveAllProperties);
}

CMyObject::CImpIDataObject::CImpIDataObject(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pMyObject(pBackObj),
	m_punkOuter(punkOuter)
{}

CMyObject::CImpIDataObject::~CImpIDataObject()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIDataObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDataObject::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIDataObject::Release()
{
	return this->m_punkOuter->Release();
}

// IDataObject methods
STDMETHODIMP CMyObject::CImpIDataObject::GetData(
									FORMATETC		*	pFormatetc,  
									STGMEDIUM		*	pmedium)
{
	HRESULT					hr;

	// check if can handle this
	hr = this->QueryGetData(pFormatetc);
	if (SUCCEEDED(hr))
	{
		if (0 != (pFormatetc->tymed & TYMED_HGLOBAL))
		{
			pmedium->tymed			= TYMED_HGLOBAL;
			pmedium->pUnkForRelease	= NULL;
			hr = this->GetDataToHGlobal(&(pmedium->hGlobal));
		}
		if (FAILED(hr) || 0 != (pFormatetc->tymed & TYMED_FILE))
		{
			/*
			// need a temporary file name
			LPTSTR				szTempFile;

			hr = pUtils->FormTempFileName(&szTempFile);
			if (SUCCEEDED(hr))
			{
				pmedium->tymed			= TYMED_FILE;
				pmedium->pUnkForRelease	= NULL;
				hr = this->GetDataToFile(szTempFile);
				if (SUCCEEDED(hr))
				{
					pmedium->lpszFileName = NULL;
					pUtils->AnsiToOleStr(szTempFile, &(pmedium->lpszFileName));
				}
				CoTaskMemFree((LPVOID) szTempFile);
			}
			*/
		}
		if (FAILED(hr) || 0 != (pFormatetc->tymed & TYMED_ISTREAM))
		{
			pmedium->tymed			= TYMED_ISTREAM;
			pmedium->pUnkForRelease	= NULL;
			hr = this->GetDataToStream(&(pmedium->pstm));
		}
		if (FAILED(hr) || 0 != (pFormatetc->tymed & TYMED_GDI))
		{
			pmedium->tymed			= TYMED_GDI;
			pmedium->pUnkForRelease	= NULL;
			hr = this->GetDataToBitmap(&(pmedium->hBitmap));
		}
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDataObject::GetDataHere(
									FORMATETC		*	pFormatetc, 
									STGMEDIUM		*	pmedium)
{
	return E_NOTIMPL;
}

// get data
HRESULT CMyObject::CImpIDataObject::GetDataToHGlobal(
									HGLOBAL			*	phglobal)
{
	HRESULT					hr;
	IStream				*	pstm;

	*phglobal = NULL;
	// write to a stream
	hr = CreateStreamOnHGlobal(NULL, FALSE, &pstm);
	if (SUCCEEDED(hr))
	{
		hr = this->SaveToStream(pstm);
		if (SUCCEEDED(hr))
		{
			hr = GetHGlobalFromStream(pstm, phglobal);
		}
		pstm->Release();
	}
	return hr;
}

// write data to a stream
HRESULT CMyObject::CImpIDataObject::SaveToStream(
									IStream			*	pStm)
{
	HRESULT					hr;
	IPersistStreamInit	*	pPersistStream;
	ULARGE_INTEGER			uli;
	CLSID					clsid;

	hr = this->m_pMyObject->QueryInterface(
		IID_IPersistStreamInit, (LPVOID*) &pPersistStream);
	if (SUCCEEDED(hr))
	{
		// set the stream size
		pPersistStream->GetSizeMax(&uli);
		pStm->SetSize(uli);
		// write the class id
		pPersistStream->GetClassID(&clsid);
		WriteClassStm(pStm, clsid);
		// write to the stream
		hr = pPersistStream->Save(pStm, FALSE);
		pPersistStream->Release();
	}
	return hr;
}

HRESULT CMyObject::CImpIDataObject::GetDataToFile(
									LPCTSTR				szFile)
{
	return E_NOTIMPL;
}

HRESULT CMyObject::CImpIDataObject::GetDataToStream(
									IStream			**	ppStm)
{
	HRESULT					hr;
//	IStream				*	pstm;

	// write to a stream
	hr = CreateStreamOnHGlobal(NULL, TRUE, ppStm);
	if (SUCCEEDED(hr))
	{
		hr = this->SaveToStream(*ppStm);
	}
	return hr;
}

HRESULT CMyObject::CImpIDataObject::GetDataToBitmap(
									HBITMAP			*	phbitmap)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyObject::CImpIDataObject::QueryGetData(
									FORMATETC		*	pFormatetc)
{
	HRESULT					hr;
	IEnumFORMATETC		*	pEnum;
	BOOL					fDone	= FALSE;
	FORMATETC				format;					// format from the enumeration

	// enumerate the available getting formats
	hr = this->EnumFormatEtc(DATADIR_GET, &pEnum);
	if (SUCCEEDED(hr))
	{
		fDone = FALSE;
		while (!fDone && S_OK == pEnum->Next(1, &format, NULL))
		{
			// check the format
			if (format.lindex != pFormatetc->lindex)
				hr = DV_E_LINDEX;
			else if (format.dwAspect != pFormatetc->dwAspect)
				hr = DV_E_DVASPECT;
			else if (format.cfFormat != pFormatetc->cfFormat)
				hr = DV_E_FORMATETC;
			else if (0 == (format.tymed & pFormatetc->tymed))
				hr = DV_E_TYMED;
			else hr = S_OK;
			fDone = SUCCEEDED(hr);
		}
		pEnum->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDataObject::GetCanonicalFormatEtc(
									FORMATETC		*	pFormatetcIn,  
									FORMATETC		*	pFormatetcOut)
{
	pFormatetcOut->ptd		= NULL;
	pFormatetcOut->cfFormat	= pFormatetcIn->cfFormat;
	pFormatetcOut->dwAspect	= pFormatetcIn->dwAspect;
	pFormatetcOut->lindex	= pFormatetcIn->lindex;
	pFormatetcOut->tymed	= pFormatetcIn->tymed;
	return DATA_S_SAMEFORMATETC;
}

STDMETHODIMP CMyObject::CImpIDataObject::SetData(
									FORMATETC		*	pFormatetc,  
									STGMEDIUM		*	pmedium,     
									BOOL				fRelease)
{
	HRESULT					hr;
	IEnumFORMATETC		*	pEnum;
	FORMATETC				format;
	BOOL					fDone;

	// enumerate the available setting formats
	hr = this->EnumFormatEtc(DATADIR_SET, &pEnum);
	// loop through this to find the entered format
	if (SUCCEEDED(hr))
	{
		fDone	= FALSE;
		while (1 == pEnum->Next(1, &format, NULL))
		{
			// check the format
			if (format.lindex != pFormatetc->lindex)
				hr = DV_E_LINDEX;
			else if (format.dwAspect != pFormatetc->dwAspect)
				hr = DV_E_DVASPECT;
			else if (format.cfFormat != pFormatetc->cfFormat)
				hr = DV_E_FORMATETC;
			else if (format.tymed != pFormatetc->tymed)
				hr = DV_E_TYMED;
			else hr = S_OK;
			fDone = SUCCEEDED(hr);
		}
		pEnum->Release();
	}
	if (SUCCEEDED(hr))
	{
		// succeeded in finding the format
		switch (pFormatetc->tymed)
		{
		case TYMED_HGLOBAL:
			hr = this->SetDataFromHGlobal(pmedium->hGlobal);
			break;
		case TYMED_FILE:
			{
				LPTSTR			szFile	= NULL;
				SHStrDup(pmedium->lpszFileName, &szFile);
				hr = this->SetDataFromFile(szFile);
				CoTaskMemFree((LPVOID) szFile);
			}
			break;
		case TYMED_ISTREAM:
			hr = this->SetDataFromStream(pmedium->pstm);
			break;
		case TYMED_GDI:
			hr = this->SetDataFromBitmap(pmedium->hBitmap);
			break;
		default:
			hr = E_FAIL;
		}
	}
	if (fRelease) ReleaseStgMedium(pmedium);
	return hr;
}

// set data
HRESULT CMyObject::CImpIDataObject::SetDataFromHGlobal(
									HGLOBAL				hglobal)
{
	HRESULT					hr;
	IStream				*	pStm;

	hr = CreateStreamOnHGlobal(hglobal, FALSE, &pStm);
	if (SUCCEEDED(hr))
	{
		hr = this->SetDataFromStream(pStm);
		pStm->Release();
	}
	return hr;
}

HRESULT CMyObject::CImpIDataObject::SetDataFromFile(
									LPCTSTR				szFile)
{
	// TODO handle this
	return E_FAIL;
}

HRESULT CMyObject::CImpIDataObject::SetDataFromStream(
									IStream			*	pStm)
{
	HRESULT					hr;
	IPersistStreamInit	*	pPersistStream;
	CLSID					clsid;

	hr = this->m_pMyObject->QueryInterface(IID_IPersistStreamInit, (LPVOID*) &pPersistStream);
	if (SUCCEEDED(hr))
	{
		ReadClassStm(pStm, &clsid);
		hr = pPersistStream->Load(pStm);
		pPersistStream->Release();
	}
	return hr;
}

HRESULT CMyObject::CImpIDataObject::SetDataFromBitmap(
									HBITMAP				hbitmap)
{
	// TODO handle this
	return E_FAIL;
}

STDMETHODIMP CMyObject::CImpIDataObject::EnumFormatEtc(
									DWORD				dwDirection,  
									IEnumFORMATETC	**	ppenumFormatetc)
{
	HRESULT					hr;
	ULONG					cFormats;
	FORMATETC			*	pFormats	= NULL;
	ULONG					i;
	CMyEnumFormatetc	*	pEnum;

	*ppenumFormatetc	= NULL;
	// form the array
	switch (dwDirection)
	{
	case DATADIR_GET:
		cFormats	= 3;
		pFormats	= new FORMATETC [cFormats];
		for (i=0; i<2; i++) pFormats[i].cfFormat	= CF_TEXT;
		pFormats[0].tymed		= TYMED_HGLOBAL;
		pFormats[1].tymed		= TYMED_ISTREAM;
		pFormats[2].cfFormat	= CF_BITMAP;
		pFormats[2].tymed		= TYMED_GDI;
		break;
	case DATADIR_SET:
		cFormats	= 4;
		pFormats	= new FORMATETC [cFormats];
		for (i=0; i<3; i++) pFormats[i].cfFormat	= CF_TEXT;
		pFormats[0].tymed		= TYMED_HGLOBAL;
		pFormats[1].tymed		= TYMED_ISTREAM;
		pFormats[2].tymed		= TYMED_FILE;
		pFormats[3].cfFormat	= CF_BITMAP;
		pFormats[3].tymed		= TYMED_GDI;
		break;
	default:
		return E_INVALIDARG;
	}
	// set the common properties
	for (i=0; i<cFormats; i++)
	{
		pFormats[i].dwAspect		= DVASPECT_CONTENT;
		pFormats[i].lindex			= -1;
		pFormats[i].ptd				= NULL;
	}
	pEnum = new CMyEnumFormatetc(this);
	pEnum->Init(cFormats, pFormats, 0);
	hr = pEnum->QueryInterface(IID_IEnumFORMATETC, (LPVOID*)ppenumFormatetc);
	if (FAILED(hr)) delete pEnum;
	delete [] pFormats;
	return hr;
}

STDMETHODIMP CMyObject::CImpIDataObject::DAdvise(
									FORMATETC		*	pFormatetc,  
									DWORD				advf,              
									IAdviseSink		*	pAdvSink,  
									DWORD			*	pdwConnection)
{
	HRESULT				hr;

	if (NULL == this->m_pMyObject->m_pDataAdviseHolder)
	{
		hr = CreateDataAdviseHolder(&(this->m_pMyObject->m_pDataAdviseHolder));
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		hr = this->m_pMyObject->m_pDataAdviseHolder->Advise(
			this, pFormatetc, advf, pAdvSink, pdwConnection);
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIDataObject::DUnadvise(
									DWORD				dwConnection)
{
	HRESULT				hr;

	if (NULL != this->m_pMyObject->m_pDataAdviseHolder)
	{
		hr = this->m_pMyObject->m_pDataAdviseHolder->Unadvise(dwConnection);
	}
	else hr = OLE_E_NOCONNECTION;
	return hr;
}

STDMETHODIMP CMyObject::CImpIDataObject::EnumDAdvise(
									IEnumSTATDATA	**	ppenumAdvise)
{
	HRESULT				hr;

	if (NULL != this->m_pMyObject->m_pDataAdviseHolder)
	{
		hr = this->m_pMyObject->m_pDataAdviseHolder->EnumAdvise(ppenumAdvise);
	}
	else
	{
		*ppenumAdvise = NULL;
		hr = OLE_E_NOCONNECTION;
	}
	return hr;
}

CMyObject::CImpIOleObject::CImpIOleObject(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pAdviseHolder(NULL),
	m_pClientSite(NULL)
{}

CMyObject::CImpIOleObject::~CImpIOleObject()
{
	Utils_RELEASE_INTERFACE(m_pAdviseHolder);
	Utils_RELEASE_INTERFACE(m_pClientSite);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIOleObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleObject::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleObject::Release()
{
	return this->m_punkOuter->Release();
}

// IOleObject methods
STDMETHODIMP CMyObject::CImpIOleObject::SetClientSite(
									IOleClientSite	*	pClientSite)
{
	Utils_RELEASE_INTERFACE(this->m_pClientSite);
	if (NULL != pClientSite)
	{
		this->m_pClientSite = pClientSite;
		this->m_pClientSite->AddRef();
	}
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetClientSite(
									IOleClientSite	**	ppClientSite)
{
	if (NULL != this->m_pClientSite)
	{
		*ppClientSite = this->m_pClientSite;
		this->m_pClientSite->AddRef();
		return S_OK;
	}
	else
	{
		*ppClientSite = NULL;
		return E_FAIL;
	}
}

STDMETHODIMP CMyObject::CImpIOleObject::SetHostNames(
									LPCOLESTR			szContainerApp,  
									LPCOLESTR			szContainerObj)
{
	LPTSTR					containerApp	= NULL;
	LPTSTR					containerObj	= NULL;

	SHStrDup(szContainerApp, &containerApp);
	SHStrDup(szContainerObj, &containerObj);
	this->m_pBackObj->m_pMyWindow->SetHostNames(containerApp, containerObj);
	CoTaskMemFree((LPVOID) containerApp);
	CoTaskMemFree((LPVOID) containerObj);
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::Close(
									DWORD				dwSaveOption)
{
	HRESULT					hr		= S_OK;
	IPersistStorage		*	pPersistStorage;
	BOOL					fDirty;					// dirty flag
	HWND					hwnd;					// control window
	CNamedObjects		*	pNamedObjects		= GetNamedObjects();

	// remove this object to the list of named objects held by the server
	if (NULL != pNamedObjects)
		pNamedObjects->AddNamedObject(this->m_pBackObj, FALSE);

	// check if this is dirty
	hr = this->m_pBackObj->QueryInterface(IID_IPersistStorage, (LPVOID*) &pPersistStorage);
	if (SUCCEEDED(hr))
	{
		fDirty = S_OK == pPersistStorage->IsDirty();
		if (fDirty)
		{
			// don't prompt the user 
			if (OLECLOSE_PROMPTSAVE == dwSaveOption || OLECLOSE_SAVEIFDIRTY == dwSaveOption)
			{
				hr = pPersistStorage->Save(NULL, TRUE);
				if (SUCCEEDED(hr))
					pPersistStorage->SaveCompleted(NULL);
			}
		}
		pPersistStorage->Release();
	}
	if (SUCCEEDED(hr))
	{
		// check if the control window exists
		hwnd		= this->m_pBackObj->m_pMyWindow->GetMyWindow();
		if (NULL != hwnd) DestroyWindow(hwnd);
		this->m_pBackObj->SendOnDataChange(TRUE);
		// close the window
		if (NULL != this->m_pClientSite)
			this->m_pClientSite->OnShowWindow(FALSE);
		// close advises
		if (NULL != this->m_pAdviseHolder)
			this->m_pAdviseHolder->SendOnClose();
		// release everyone
		Utils_RELEASE_INTERFACE(this->m_pAdviseHolder);
		Utils_RELEASE_INTERFACE(this->m_pClientSite);
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::SetMoniker(
									DWORD				dwWhichMoniker,  
									IMoniker		*	pmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetMoniker(
									DWORD				dwAssign,  
									DWORD				dwWhichMoniker,            
									IMoniker		**	ppmk)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyObject::CImpIOleObject::InitFromData(
									IDataObject		*	pDataObject,  
									BOOL				fCreation,            
									DWORD				dwReserved)
{
	HRESULT					hr;
	IDataObject			*	pourData;			// our data object

	// can only create the object
	if (!fCreation) return S_FALSE;
	if (NULL != pDataObject)
	{
		// attempt to initialize from the data object
		hr = this->m_pBackObj->QueryInterface(IID_IDataObject, (LPVOID*) &pourData);
		if (SUCCEEDED(hr))
		{
			hr = this->DoCopyData(pDataObject, pourData);
			pourData->Release();
		}
	}
	else
		hr = S_OK;
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetClipboardData(
									DWORD				dwReserved,  
									IDataObject		**	ppDataObject)
{
	HRESULT					hr;
	IOleObject			*	pOleObject;
	IDataObject			*	pDataObject;			// our data object

	*ppDataObject = NULL;
	// create a new object
	hr = CoCreateInstance(CLSID_SciUsbMono, NULL, CLSCTX_INPROC_SERVER, IID_IOleObject, (LPVOID*) &pOleObject);
	if (SUCCEEDED(hr))
	{
		// initialize from our data
		hr = this->m_pBackObj->QueryInterface(IID_IDataObject, (LPVOID*) &pDataObject);
		if (SUCCEEDED(hr))
		{
			hr = pOleObject->InitFromData(pDataObject, TRUE, 0);
			pDataObject->Release();
		}
		if (SUCCEEDED(hr))
		{
			hr = pOleObject->QueryInterface(IID_IDataObject, (LPVOID*) ppDataObject);
		}
		pOleObject->Release();
	}
	return hr;
}

// copy data objects
HRESULT CMyObject::CImpIOleObject::DoCopyData(
									IDataObject		*	pDataSource,
									IDataObject		*	pDataDest)
{
	HRESULT					hr;
	IEnumFORMATETC		*	pEnum;				// enumeration of available setting formats
	FORMATETC				format;				// data format
	STGMEDIUM				medium;				// storage medium
	BOOL					fDone		= FALSE;	// completed flag

	// obtain the enumeration of available setting formats
	hr = pDataDest->EnumFormatEtc(DATADIR_SET, &pEnum);
	if (SUCCEEDED(hr))
	{
		fDone = FALSE;
		while (!fDone && S_OK == pEnum->Next(1, &format, NULL))
		{
			// check the format
			hr = pDataSource->QueryGetData(&format);
			if (SUCCEEDED(hr))
			{
				hr = pDataSource->GetData(&format, &medium);
				if (SUCCEEDED(hr))
				{
					hr = pDataDest->SetData(&format, &medium, TRUE);
					fDone = SUCCEEDED(hr);
				}
			}
		}
		pEnum->Release();
	}
	return fDone ? S_OK : E_FAIL;
}

STDMETHODIMP CMyObject::CImpIOleObject::DoVerb(
									LONG				iVerb,          
									LPMSG				lpmsg,         
									IOleClientSite	*	pActiveSite,                      
									LONG				lindex,         
									HWND				hwndParent,     
									LPCRECT				lprcPosRect)
{
	HRESULT hr = E_NOTIMPL;
	switch (iVerb)
	{
	case OLEIVERB_PRIMARY:
		hr = DoVerbPrimary(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_SHOW:
		hr = DoVerbShow(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_INPLACEACTIVATE:
		hr = DoVerbInPlaceActivate(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_UIACTIVATE:
		hr = DoVerbUIActivate(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_HIDE:
		hr = DoVerbHide(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_OPEN:
		hr = DoVerbOpen(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_DISCARDUNDOSTATE:
		hr = DoVerbDiscardUndo(lprcPosRect, hwndParent);
		break;
	case OLEIVERB_PROPERTIES:
		hr = DoVerbProperties(lprcPosRect, hwndParent);
		break;
	default:
		break;
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::EnumVerbs(
									IEnumOLEVERB	**	ppEnumOLEVERB)
{
	return OleRegEnumVerbs(CLSID_SciUsbMono, ppEnumOLEVERB);
}

STDMETHODIMP CMyObject::CImpIOleObject::Update()
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::IsUpToDate()
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetUserClassID(
									CLSID			*	pClsid)
{
	*pClsid = CLSID_SciUsbMono;
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetUserType(
									DWORD				dwFormOfType,  
									LPOLESTR		*	pszUserType)
{
	return OleRegGetUserType(CLSID_SciUsbMono, dwFormOfType, pszUserType);
}

STDMETHODIMP CMyObject::CImpIOleObject::SetExtent(
									DWORD				dwDrawAspect,  
									SIZEL			*	psizel)
{
//	HRESULT				hr;
	HWND				hwndControl = this->m_pBackObj->m_pMyWindow->GetMyWindow();
	RECT				rc;			// window rectangle
	HWND				hwndParent;				// parent window

	if (NULL != hwndControl)
	{
		GetWindowRect(hwndControl, &rc);
		// convert to parent window client coordinates
		hwndParent = GetParent(hwndControl);
		if (NULL != hwndParent)
			MapWindowPoints(NULL, hwndParent, (LPPOINT) &rc, 2);
		// convert the size to pixels
		Utils_HimetricToPixels(psizel);
		// move the control window
		MoveWindow(hwndControl, rc.left, rc.top, psizel->cx, psizel->cy, TRUE);
	}
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetExtent(
									DWORD				dwDrawAspect,  
									SIZEL			*	psizel)
{
	HRESULT					hr;
	IViewObject2		*	pViewObject;

	hr = this->m_pBackObj->QueryInterface(IID_IViewObject2, (LPVOID*) &pViewObject);
	if (SUCCEEDED(hr))
	{
		hr = pViewObject->GetExtent(dwDrawAspect, -1, NULL, psizel);
		pViewObject->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::Advise(
									IAdviseSink		*	pAdvSink,  
									DWORD			*	pdwConnection)
{
	HRESULT					hr;

	if (NULL == this->m_pAdviseHolder)
	{
		hr = CreateOleAdviseHolder(&(this->m_pAdviseHolder));
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
		hr = this->m_pAdviseHolder->Advise(pAdvSink, pdwConnection);
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::Unadvise(
									DWORD				dwConnection)
{
	HRESULT					hr;

	if (NULL != this->m_pAdviseHolder)
		hr = this->m_pAdviseHolder->Unadvise(dwConnection);
	else
		hr = OLE_E_NOCONNECTION;
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::EnumAdvise(
									IEnumSTATDATA	**	ppenumAdvise)
{
	HRESULT					hr;

	if (NULL != this->m_pAdviseHolder)
		hr = this->m_pAdviseHolder->EnumAdvise(ppenumAdvise);
	else
	{
		*ppenumAdvise = NULL;
		hr = E_FAIL;
	}
	return hr;
}

STDMETHODIMP CMyObject::CImpIOleObject::GetMiscStatus(
									DWORD				dwAspect,  
									DWORD			*	pdwStatus)
{
	return OLE_S_USEREG;
}

STDMETHODIMP CMyObject::CImpIOleObject::SetColorScheme(
									LOGPALETTE		*	pLogpal)
{
	return E_NOTIMPL;
}

// helpers for DoVerb
HRESULT CMyObject::CImpIOleObject::DoVerbPrimary(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT			hr;
	BOOL			bDesignMode = FALSE;
	VARIANT			Value;
	BOOL			fUserMode;

	hr = this->m_pBackObj->GetAmbientProperty(DISPID_AMBIENT_USERMODE, &Value);
	if (SUCCEEDED(hr))
	{
		VariantToBoolean(Value, &fUserMode);
		bDesignMode = !fUserMode;
	}
	// if container doesn't support this property
//	bDesignMode = !m_pBackObj->m_fUserMode;
	if (bDesignMode)
		return DoVerbProperties(prcPosRect, hwndParent);
	else
		return DoVerbInPlaceActivate(prcPosRect, hwndParent);
	return S_OK;
}

HRESULT CMyObject::CImpIOleObject::DoVerbShow(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT hr;
	hr = InPlaceActivate(OLEIVERB_SHOW, prcPosRect);
	return hr;
}

HRESULT CMyObject::CImpIOleObject::DoVerbInPlaceActivate(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT hr;
	hr = InPlaceActivate(OLEIVERB_UIACTIVATE, prcPosRect);
//	if (SUCCEEDED(hr))
//		m_pBackObj->ViewChange();
	return hr;
}

HRESULT CMyObject::CImpIOleObject::DoVerbUIActivate(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT hr = S_OK;
//	if (!m_pBackObj->m_fUIActive)
//	{
		hr = InPlaceActivate(OLEIVERB_UIACTIVATE, prcPosRect);
//	}
	return hr;
}

HRESULT CMyObject::CImpIOleObject::DoVerbHide(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT					hr;
	IOleInPlaceObject	*	pInPlaceObject;
	HWND					hwndControl;

	hr = this->m_pBackObj->QueryInterface(IID_IOleInPlaceObject,
		(LPVOID*) &pInPlaceObject);
	if (SUCCEEDED(hr))
	{
		pInPlaceObject->UIDeactivate();
		pInPlaceObject->GetWindow(&hwndControl);
		if (hwndControl)
			ShowWindow(hwndControl, SW_HIDE);
		pInPlaceObject->Release();
	}
	return hr;
}

HRESULT CMyObject::CImpIOleObject::DoVerbOpen(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	return S_OK;
}

HRESULT CMyObject::CImpIOleObject::DoVerbDiscardUndo(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	return S_OK;
}

HRESULT CMyObject::CImpIOleObject::DoVerbProperties(
									LPCRECT				prcPosRect, 
									HWND				hwndParent)
{
	HRESULT						hr;
	IOleControlSite			*	pControlSite;
	IUnknown				*	punk;
	ISpecifyPropertyPages	*	pSpecifyPropPages;
	LPOLESTR					szTitle = NULL;
	CAUUID						pages;				// pages structure
	LCID						lcid;
	VARIANT						Value;
	long						lval;

	// have the site display their property pages
	if (NULL != this->m_pClientSite)
	{
		hr = this->m_pClientSite->QueryInterface(IID_IOleControlSite, 
			(LPVOID*) &pControlSite);
		if (SUCCEEDED(hr))
		{
			hr = pControlSite->ShowPropertyFrame();
			pControlSite->Release();
			if (SUCCEEDED(hr)) return hr;
		}
	}
	// the client could not display the property pages, we will need to do it ourselvels
	hr = this->GetUserType(USERCLASSTYPE_SHORT, &szTitle);
	if (SUCCEEDED(hr))
	{
		hr = this->m_pBackObj->QueryInterface(IID_IUnknown, (LPVOID*) &punk);
		if (SUCCEEDED(hr))
		{
			hr = this->m_pBackObj->QueryInterface(IID_ISpecifyPropertyPages,
				(LPVOID*) &pSpecifyPropPages);
			if (SUCCEEDED(hr))
			{
				hr = pSpecifyPropPages->GetPages(&pages);
				pSpecifyPropPages->Release();
			}
			if (SUCCEEDED(hr))
			{
				// locale id
				lcid	= LOCALE_SYSTEM_DEFAULT;
				hr = this->m_pBackObj->GetAmbientProperty(DISPID_AMBIENT_LOCALEID, &Value);
				if (SUCCEEDED(hr))
				{
					VariantToInt32(Value, &lval);
					lcid = (LCID) lval;
				}
				hr = OleCreatePropertyFrame(
					hwndParent,
					0, 0,
					szTitle,
					1, &punk,
					pages.cElems, pages.pElems,
					lcid,
					0, NULL);
				CoTaskMemFree((LPVOID) pages.pElems);
			}
			punk->Release();
		}
		CoTaskMemFree((LPVOID) szTitle);
	}
	return hr;
}

HRESULT CMyObject::CImpIOleObject::InPlaceActivate(
									LONG				iVerb, 
									const RECT		*	prcPosRect)
{
	HRESULT							hr;
	IOleInPlaceObject			*	pInPlaceObject;
	IOleInPlaceSite				*	pInPlaceSite;
	IOleInPlaceFrame			*	pInPlaceFrame;
	OLEINPLACEFRAMEINFO				frameInfo;
	RECT							rcPos, rcClip;
	HWND							hwndParent;			// parent window
	IOleInPlaceUIWindow			*	pInPlaceUIWindow	= NULL;
	HWND							hwndControl;
	IOleControl					*	pOleControl;		// Ole control
	CNamedObjects				*	pNamedObjects		= GetNamedObjects();

	if (NULL == this->m_pClientSite) return S_OK;
	hr = this->m_pBackObj->QueryInterface(IID_IOleInPlaceObject, (LPVOID*) &pInPlaceObject);
	if (SUCCEEDED(hr))
	{
		hr = this->m_pClientSite->QueryInterface(IID_IOleInPlaceSite, 
			(LPVOID*) &pInPlaceSite);
		if (SUCCEEDED(hr))
		{
			// CanInPlaceActivate returns S_FALSE or S_OK
			hr = pInPlaceSite->CanInPlaceActivate();
			if (S_OK == hr)
			{
				pInPlaceSite->OnInPlaceActivate();
				hwndControl = this->m_pBackObj->m_pMyWindow->GetMyWindow();
				if (NULL != hwndControl)
				{

				}
				else
				{
					// create the window
					frameInfo.cb = sizeof(OLEINPLACEFRAMEINFO);
					if (pInPlaceSite->GetWindow(&hwndParent) == S_OK)
					{
						hr = pInPlaceSite->GetWindowContext(
								&pInPlaceFrame,
								&pInPlaceUIWindow, 
								&rcPos, 
								&rcClip, 
								&frameInfo);
						if (SUCCEEDED(hr))
						{
							hwndControl = m_pBackObj->m_pMyWindow->CreateInPlaceWindow(
								GetOurInstance(), hwndParent, &rcPos);
							Utils_RELEASE_INTERFACE(pInPlaceFrame);
							Utils_RELEASE_INTERFACE(pInPlaceUIWindow);
							pInPlaceObject->SetObjectRects(&rcPos, &rcClip);
							// update the ambients
							hr = this->m_pBackObj->QueryInterface(IID_IOleControl,
								(LPVOID*) &pOleControl);
							if (SUCCEEDED(hr))
							{
								pOleControl->OnAmbientPropertyChange(DISPID_UNKNOWN);
								pOleControl->Release();
								// add this object to the list of named objects held by the server
								if (NULL != pNamedObjects)
									pNamedObjects->AddNamedObject(this->m_pBackObj, TRUE);
							}
						}
					}
				}
			}
			else
			{
				// CanInPlaceActivate returned S_FALSE.
				Utils_RELEASE_INTERFACE(pInPlaceSite);
				hr = SUCCEEDED(hr) ? E_FAIL : hr;
			}
			pInPlaceSite->Release();
		}
		pInPlaceObject->Release();
	}
	return hr;
}

BOOL CMyObject::CImpIOleObject::DoesVerbUIActivate(
									LONG				iVerb)
{
	BOOL b = FALSE;
	switch (iVerb)
	{
		case OLEIVERB_UIACTIVATE:
		case OLEIVERB_PRIMARY:
			b = TRUE;
			break;
	}
//	// if no ambient dispatch then in old style OLE container
//	if (DoesVerbActivate(iVerb) && m_spAmbientDispatch == NULL)
//		b = TRUE;
	return b;
}

BOOL CMyObject::CImpIOleObject::SetControlFocus(
									IOleInPlaceSite	*	pIOleInPlaceSite,
									BOOL				bGrab)
{
	return FALSE;
}


CMyObject::CImpIOleControl::CImpIOleControl(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{}

CMyObject::CImpIOleControl::~CImpIOleControl()
{}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIOleControl::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleControl::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleControl::Release()
{
	return this->m_punkOuter->Release();
}

// IOleControl methods
STDMETHODIMP CMyObject::CImpIOleControl::GetControlInfo(
									CONTROLINFO		*	pCI)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyObject::CImpIOleControl::OnMnemonic(
									LPMSG				pMsg)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyObject::CImpIOleControl::OnAmbientPropertyChange(
									DISPID				dispID)
{
	HRESULT					hr;
	IDispatch			*	pdispAmbient;

	hr = this->m_pBackObj->GetAmbientObject(&pdispAmbient);
	if (SUCCEEDED(hr))
	{
		switch (dispID)
		{
		case DISPID_AMBIENT_FONT:
			this->OnAmbientChangeFont(pdispAmbient);
			break;
		case DISPID_AMBIENT_LOCALEID:
			this->OnAmbientChangeLCID(pdispAmbient);
			break;
		case DISPID_AMBIENT_SHOWGRABHANDLES:
			this->OnAmbientChangeShowGrabHandles(pdispAmbient);
			break;
		case DISPID_AMBIENT_SHOWHATCHING:
			this->OnAmbientChangeShowHatching(pdispAmbient);
			break;
		case DISPID_AMBIENT_SUPPORTSMNEMONICS:
			this->OnAmbientChangeSupportsMnemonics(pdispAmbient);
			break;
		case DISPID_AMBIENT_UIDEAD:
			this->OnAmbientChangeUIDead(pdispAmbient);
			break;
		case DISPID_AMBIENT_USERMODE:
			this->OnAmbientChangeUserMode(pdispAmbient);
			break;
		case DISPID_UNKNOWN:
			this->OnAmbientChangeFont(pdispAmbient);
			this->OnAmbientChangeLCID(pdispAmbient);
			this->OnAmbientChangeShowGrabHandles(pdispAmbient);
			this->OnAmbientChangeShowHatching(pdispAmbient);
			this->OnAmbientChangeSupportsMnemonics(pdispAmbient);
			this->OnAmbientChangeUIDead(pdispAmbient);
			this->OnAmbientChangeUserMode(pdispAmbient);
			break;
		default:
			break;
		}
		pdispAmbient->Release();
	}
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleControl::FreezeEvents(
									BOOL				bFreeze)
{
	this->m_pBackObj->m_fFreezeEvents = bFreeze;
	return S_OK;
}

// ambient property changes
HRESULT CMyObject::CImpIOleControl::OnAmbientChangeFont(
									IDispatch		*	pdispAmbient)
{
	HRESULT				hr;
	VARIANT				var;

	hr = Utils_InvokePropertyGet(pdispAmbient, DISPID_AMBIENT_FONT, NULL, 0, &var);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == var.vt)
		{
//			this->m_pBackObj->SetFont(var.pdispVal);
		}
		VariantClear(&var);
	}
	return S_OK;
}

HRESULT	CMyObject::CImpIOleControl::OnAmbientChangeLCID(
									IDispatch		*	pdispAmbient)
{
	HRESULT				hr;
	VARIANT				var;
	long				lval;

	hr = Utils_InvokePropertyGet(pdispAmbient, DISPID_AMBIENT_LOCALEID, NULL, 0, &var);
	if (SUCCEEDED(hr))
	{
		VariantToInt32(var, &lval);
//		this->m_pBackObj->m_lcid = (LCID) lval;
	}
	return hr;
}

HRESULT CMyObject::CImpIOleControl::OnAmbientChangeUIDead(
									IDispatch		*	pdispAmbient)
{
	return S_OK;
}

HRESULT CMyObject::CImpIOleControl::OnAmbientChangeShowGrabHandles(
									IDispatch		*	pdispAmbient)
{
	return S_OK;
}

HRESULT CMyObject::CImpIOleControl::OnAmbientChangeShowHatching(
									IDispatch		*	pdispAmbient)
{
	return S_OK;
}

HRESULT CMyObject::CImpIOleControl::OnAmbientChangeUserMode(
									IDispatch		*	pdispAmbient)
{
	HRESULT					hr;
	VARIANT					var;

	hr = Utils_InvokePropertyGet(pdispAmbient, DISPID_AMBIENT_USERMODE, NULL, 0, &var);
	if (SUCCEEDED(hr))
	{
	//	VariantToBoolean(var, &(this->m_pBackObj->m_fUserMode));
	}
	return S_OK;
}

HRESULT CMyObject::CImpIOleControl::OnAmbientChangeSupportsMnemonics(
									IDispatch		*	pdispAmbient)
{
	return S_OK;
}


CMyObject::CImpIOleInPlaceObject::CImpIOleInPlaceObject(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{}

CMyObject::CImpIOleInPlaceObject::~CImpIOleInPlaceObject()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIOleInPlaceObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleInPlaceObject::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleInPlaceObject::Release()
{
	return this->m_punkOuter->Release();
}

// IOleWindow methods
STDMETHODIMP CMyObject::CImpIOleInPlaceObject::GetWindow(
									HWND			*	phwnd)
{
	*phwnd = this->m_pBackObj->m_pMyWindow->GetMyWindow();
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceObject::ContextSensitiveHelp(
									BOOL				fEnterMode)
{
	return S_OK;
}

// IOleInPlaceObject methods
STDMETHODIMP CMyObject::CImpIOleInPlaceObject::InPlaceDeactivate()
{
	HRESULT				hr;
	HWND				hwndControl;
	IOleClientSite	*	pClientSite		= NULL;
	IOleInPlaceSite	*	pInPlaceSite	= NULL;

	hr = this->m_pBackObj->GetClientSite(&pClientSite);
	if (SUCCEEDED(hr))
	{
		hr = pClientSite->QueryInterface(IID_IOleInPlaceSite, (LPVOID*) &pInPlaceSite);
		pClientSite->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pInPlaceSite->OnInPlaceDeactivate();
		pInPlaceSite->Release();
	}
	this->GetWindow(&hwndControl);
	if (NULL != hwndControl)
	{
		this->UIDeactivate();
		DestroyWindow(hwndControl);
	}
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceObject::UIDeactivate()
{
	HWND				hwndControl;
	this->GetWindow(&hwndControl);
	if (NULL != hwndControl) ShowWindow(hwndControl, SW_HIDE);
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceObject::SetObjectRects(
									LPCRECT				lprcPosRect,  
									LPCRECT				lprcClipRect)
{
	RECT				rcPos;
	RECT				rcClip;

	CopyRect(&rcPos, lprcPosRect);
	CopyRect(&rcClip, lprcClipRect);
	this->m_pBackObj->m_pMyWindow->SetObjectRects(&rcPos, &rcClip);
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceObject::ReactivateAndUndo()
{
	HWND			hwndControl;

	this->GetWindow(&hwndControl);
	if (NULL != hwndControl)
	{
		ShowWindow(hwndControl, SW_SHOW);
		return S_OK;
	}
	else
		return INPLACE_E_NOTUNDOABLE;
}

CMyObject::CImpIOleInPlaceActiveObject::CImpIOleInPlaceActiveObject(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{}

CMyObject::CImpIOleInPlaceActiveObject::~CImpIOleInPlaceActiveObject()
{}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleInPlaceActiveObject::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIOleInPlaceActiveObject::Release()
{
	return this->m_punkOuter->Release();
}

// IOleWindow methods
STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::GetWindow(
									HWND			*	phwnd)
{
	*phwnd = this->m_pBackObj->m_pMyWindow->GetMyWindow();
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::ContextSensitiveHelp(
									BOOL				fEnterMode)
{
	return S_OK;
}

// IOleInPlaceActiveObject methods
STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::TranslateAccelerator(
									LPMSG				lpmsg)
{
	return this->m_pBackObj->m_pMyWindow->MyTranslateAccelerator(lpmsg) ? S_OK : S_FALSE;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::OnFrameWindowActivate(
									BOOL				fActivate)
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::OnDocWindowActivate(
									BOOL				fActivate)
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::ResizeBorder(
									LPCRECT				prcBorder,              
									IOleInPlaceUIWindow *pUIWindow,                                  
									BOOL				fFrameWindow)
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIOleInPlaceActiveObject::EnableModeless(
									BOOL				fEnable)
{
	return S_OK;
}

CMyObject::CImpISpecifyPropertyPages::CImpISpecifyPropertyPages(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter)
{
}

CMyObject::CImpISpecifyPropertyPages::~CImpISpecifyPropertyPages()
{
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpISpecifyPropertyPages::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpISpecifyPropertyPages::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpISpecifyPropertyPages::Release()
{
	return this->m_punkOuter->Release();
}

// ISpecifyPropertyPages method
STDMETHODIMP CMyObject::CImpISpecifyPropertyPages::GetPages(
									CAUUID			*	pPages)
{
	pPages->cElems		= 1;
	pPages->pElems		= (GUID*) CoTaskMemAlloc(1 * sizeof(GUID));
	pPages->pElems[0]	= CLSID_PropPageSciUsbMono;
	return S_OK;
}

CMyObject::CImpIViewObject2::CImpIViewObject2(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pBackObj(pBackObj),
	m_punkOuter(punkOuter),
	m_pIAdviseSink(NULL),
	m_advf(0)
{}

CMyObject::CImpIViewObject2::~CImpIViewObject2()
{
	Utils_RELEASE_INTERFACE(this->m_pIAdviseSink);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImpIViewObject2::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImpIViewObject2::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImpIViewObject2::Release()
{
	return this->m_punkOuter->Release();
}

// IViewObject methods
STDMETHODIMP CMyObject::CImpIViewObject2::Draw(
									DWORD				dwAspect,   //Aspect to be drawn
									LONG				lindex,     
									void *				pvAspect,  //Pointer to DVASPECTINFO structure or NULL
									DVTARGETDEVICE *	ptd,
									HDC					hicTargetDev, //Information context for the target device
									HDC					hdcDraw,      //Device context on which to draw
									LPCRECTL			lprcBounds,
									LPCRECTL			lprcWBounds,
									BOOL (__stdcall *	pfnContinue) (DWORD),
									DWORD				dwContinue)
{
	HRESULT					hr;
	RECT					rc;			// plotting rectangle

	// make sure that I like the aspect
	if (DVASPECT_CONTENT != dwAspect) return DV_E_DVASPECT;
	// only proceed if the device context exists - this will have to be adjusted to
	// allow for printing
	if (NULL == hdcDraw) return VIEW_E_DRAW;
	// copy the painting rectangle
	rc.left		= lprcBounds->left;
	rc.top		= lprcBounds->top;
	rc.right	= lprcBounds->right;
	rc.bottom	= lprcBounds->bottom;
	if (IsRectEmpty(&rc)) return OLE_E_INVALIDRECT;
	// paint the bitmap
	hr = this->m_pBackObj->m_pMySciUsbMono->PaintToRect(hdcDraw, &rc) ? S_OK : E_FAIL;
	return hr;
}

void CMyObject::CImpIViewObject2::FireViewChange()
{
	if (NULL != this->m_pIAdviseSink)
	{
		this->m_pIAdviseSink->OnViewChange(DVASPECT_CONTENT, -1);
		if (0 != (ADVF_ONLYONCE & this->m_advf))
		{
			Utils_RELEASE_INTERFACE(this->m_pIAdviseSink);
		}
	}
}

STDMETHODIMP CMyObject::CImpIViewObject2::GetColorSet(
									DWORD				dwAspect,   
									LONG				lindex,      
									void			*	pvAspect,  
									DVTARGETDEVICE	*	ptd,
									HDC					hicTargetDev, 
									LOGPALETTE		**	ppColorSet)
{
	return S_FALSE;
}

STDMETHODIMP CMyObject::CImpIViewObject2::Freeze(
									DWORD				dwAspect,   
									LONG				lindex,      
									void			*	pvAspect,  
									DWORD			*	pdwFreeze)
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIViewObject2::Unfreeze(
								DWORD				dwFreeze)
{
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIViewObject2::SetAdvise(
									DWORD				dwAspect,  
									DWORD				advf,      
									IAdviseSink		*	pAdvSink)
{	
	if (DVASPECT_CONTENT != dwAspect) return DV_E_DVASPECT;
	// release the current sink
	Utils_RELEASE_INTERFACE(this->m_pIAdviseSink);
	if (NULL != pAdvSink)
	{
		this->m_pIAdviseSink	= pAdvSink;
		this->m_pIAdviseSink->AddRef();
		this->m_advf	= advf;
		if (0 != (this->m_advf & ADVF_PRIMEFIRST))
		{
			this->FireViewChange();
		}
	}
	return S_OK;
}

STDMETHODIMP CMyObject::CImpIViewObject2::GetAdvise(
									DWORD			*	pdwAspect,  
									DWORD			*	padvf,      
									IAdviseSink		**	ppAdvSink)
{
	*pdwAspect	= DVASPECT_CONTENT;
	if (NULL != padvf)
		*padvf		= this->m_advf;
	if (NULL != this->m_pIAdviseSink)
	{
		*ppAdvSink = this->m_pIAdviseSink;
		this->m_pIAdviseSink->AddRef();
	}
	else
		*ppAdvSink = NULL;
	return S_OK;
}

// IViewObject2 method
STDMETHODIMP CMyObject::CImpIViewObject2::GetExtent(
									DWORD				dwAspect,  
									LONG				lindex,    //Part of the object to draw
									DVTARGETDEVICE*		ptd,
									LPSIZEL				lpsizel)
{
	SIZE					sz;
	HWND					hwndControl		= NULL;
	RECT					rc;

	// get the window size in pixels
	hwndControl		= this->m_pBackObj->m_pMyWindow->GetMyWindow();
	if (NULL != hwndControl)
	{
		GetWindowRect(hwndControl, &rc);
		sz.cx		= rc.right - rc.left;
		sz.cy		= rc.bottom - rc.top;
	}
	else
	{
		// default window size
		this->m_pBackObj->m_pMyWindow->GetDefaultExtent(&sz.cx, &sz.cy);
	}
	lpsizel->cx		= sz.cx;
	lpsizel->cy		= sz.cy;
	// convert to HIMETRIC
	Utils_PixelsToHimetric(lpsizel);
	return S_OK;
}


CMyObject::CImp_clsIMono::CImp_clsIMono(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pMyObject(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_dispidCurrentWavelength(DISPID_UNKNOWN),
	m_dispidAutoGrating(DISPID_UNKNOWN),
	m_dispidCurrentGrating(DISPID_UNKNOWN),
	m_dispidAmBusy(DISPID_UNKNOWN),
	m_dispidAmInitialized(DISPID_UNKNOWN),
	m_dispidGratingParams(DISPID_UNKNOWN),
	m_dispidMonoParams(DISPID_UNKNOWN),
	m_dispidIsValidPosition(DISPID_UNKNOWN),
	m_dispidConvertStepsToNM(DISPID_UNKNOWN),
	m_dispidDisplayConfigValues(DISPID_UNKNOWN),
	m_dispidDoSetup(DISPID_UNKNOWN),
	m_dispidReadConfig(DISPID_UNKNOWN),
	m_dispidWriteConfig(DISPID_UNKNOWN),
	m_dispidWaitForComplete(DISPID_UNKNOWN),
	m_dispidGetGratingDispersion(DISPID_UNKNOWN),
	m_dispidScanStart(DISPID_UNKNOWN)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pMyObject->GetRefTypeInfo(TEXT("_clsIMono"), &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// store the type info
		this->m_pTypeInfo	= pTypeInfo;
		this->m_pTypeInfo->AddRef();
		// get the interface ID
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pMyObject->m_iid_clsIMono	= pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(this->m_pTypeInfo, TEXT("CurrentWavelength"), &m_dispidCurrentWavelength);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AutoGrating"), &m_dispidAutoGrating);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("CurrentGrating"), &m_dispidCurrentGrating);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AmBusy"), &m_dispidAmBusy);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("AmInitialized"), &m_dispidAmInitialized);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("GratingParams"), &m_dispidGratingParams);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("MonoParams"), &m_dispidMonoParams);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("IsValidPosition"), &m_dispidIsValidPosition);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ConvertStepsToNM"), &m_dispidConvertStepsToNM);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("DisplayConfigValues"), &m_dispidDisplayConfigValues);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("DoSetup"), &m_dispidDoSetup);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("ReadConfig"), &m_dispidReadConfig);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("WriteConfig"), &m_dispidWriteConfig);
		Utils_GetMemid(this->m_pTypeInfo, TEXT("WaitForComplete"), &m_dispidWaitForComplete);
		Utils_GetMemid(this->m_pTypeInfo, L"GetGratingDispersion", &m_dispidGetGratingDispersion);
		Utils_GetMemid(this->m_pTypeInfo, L"ScanStart", &m_dispidScanStart);
		pTypeInfo->Release();
	}
}

CMyObject::CImp_clsIMono::~CImp_clsIMono()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMyObject::CImp_clsIMono::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIMono::AddRef()
{
	return this->m_punkOuter->AddRef();
}

STDMETHODIMP_(ULONG) CMyObject::CImp_clsIMono::Release()
{
	return this->m_punkOuter->Release();
}

// IDispatch methods
STDMETHODIMP CMyObject::CImp_clsIMono::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMyObject::CImp_clsIMono::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	if (NULL != this->m_pTypeInfo)
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
		return S_OK;
	}
	else
	{
		*ppTInfo	= NULL;
		return E_FAIL;
	}
}

STDMETHODIMP CMyObject::CImp_clsIMono::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_SYSTEM_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMyObject::CImp_clsIMono::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	if (dispIdMember == m_dispidCurrentWavelength)
	{
		return this->MyInvoke(DISPID_position, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidAutoGrating)
	{
		return this->MyInvoke(DISPID_autoGrating, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidCurrentGrating)
	{
		return this->MyInvoke(DISPID_currentGrating, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidAmBusy)
	{
		return this->MyInvoke(DISPID_AmBusy, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidAmInitialized)
	{
		return this->MyInvoke(DISPID_AmOpen, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidGratingParams)
	{
		return this->MyInvoke(DISPID_GratingInfo, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidMonoParams)
	{
		return this->MyInvoke(DISPID_MonoInfo, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidIsValidPosition)
	{
		return this->MyInvoke(DISPID_IsValidPosition, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidConvertStepsToNM)
	{
		return this->ConvertStepsToNM(pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidDisplayConfigValues)
	{
		return this->MyInvoke(DISPID_DisplayConfigValues, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidDoSetup)
	{
		return this->MyInvoke(DISPID_Setup, wFlags, pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidReadConfig)
	{
		return this->ReadConfig(pDispParams, pVarResult);
	}
	else if (dispIdMember == m_dispidWriteConfig)
	{
		return this->WriteConfig(pDispParams);
	}
	else if (dispIdMember == this->m_dispidWaitForComplete)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
			return this->GetWaitForComplete(pVarResult);
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
			return this->SetWaitForComplete(pDispParams);
	}
	else if (dispIdMember == this->m_dispidGetGratingDispersion)
	{
		return this->GetGratingDispersion(pDispParams, pVarResult);
	}
	else if (dispIdMember == this->m_dispidScanStart)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->ScanStart();
		}
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImp_clsIMono::MyInvoke(
									DISPID			dispIdMember,      
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		hr = pdisp->Invoke(dispIdMember, IID_NULL, 0x0409, wFlags, pDispParams, pVarResult,
			NULL, NULL);
		pdisp->Release();
	}
	return hr;
}

HRESULT CMyObject::CImp_clsIMono::ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	VARIANTARG			avarg[4];
	VARIANTARG			varg;
	VARIANT_BOOL		fStepsToNM;
	long				gratingID;
	UINT				uArgErr;

	if (4 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	if ((VT_BYREF | VT_R8) != pDispParams->rgvarg[0].vt) return DISP_E_TYPEMISMATCH;
	if ((VT_BYREF | VT_I4) != pDispParams->rgvarg[1].vt) return DISP_E_TYPEMISMATCH;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	fStepsToNM = varg.boolVal;
	hr = DispGetParam(pDispParams, 1, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	gratingID = varg.lVal;
	VariantInit(&avarg[3]);
	avarg[3].vt			= VT_BYREF | VT_BOOL;
	avarg[3].pboolVal	= &fStepsToNM;
	VariantInit(&avarg[2]);
	avarg[2].vt			= VT_BYREF | VT_I4;
	avarg[2].plVal		= &gratingID;
	VariantInit(&avarg[1]);
	avarg[1].vt			= VT_BYREF | VT_I4;
	avarg[1].plVal		= pDispParams->rgvarg[1].plVal;
	VariantInit(&avarg[0]);
	avarg[0].vt			= VT_BYREF | VT_R8;
	avarg[0].pdblVal	= pDispParams->rgvarg[0].pdblVal;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokeMethod(pdisp, DISPID_ConvertStepsToNM, avarg, 4, pVarResult);
		pdisp->Release();
	}
	return hr;
}

HRESULT CMyObject::CImp_clsIMono::ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	IDispatch		*	pdisp;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		hr = Utils_InvokePropertyPut(pdisp, DISPID_ConfigFile, &varg, 1); 
		pdisp->Release();
	}
	VariantClear(&varg);
	if (NULL != pVarResult) InitVariantFromBoolean(SUCCEEDED(hr), pVarResult);
	return hr;
}

HRESULT CMyObject::CImp_clsIMono::WriteConfig(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VARIANTARG			varSend;
	IDispatch		*	pdisp;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	if (SUCCEEDED(hr))
	{
		VariantInit(&varSend);
		varSend.vt			= VT_BYREF | VT_BSTR;
		varSend.pbstrVal	= &(varg.bstrVal);
		hr = Utils_InvokeMethod(pdisp, DISPID_WriteConfig, &varg, 1, NULL);
		pdisp->Release();
	}
	VariantClear(&varg);
	return hr;
}

// wait for complete is always true
HRESULT CMyObject::CImp_clsIMono::GetWaitForComplete(
									VARIANT		*	pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(TRUE, pVarResult);
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::SetWaitForComplete(
									DISPPARAMS	*	pDispParams)
{
	return S_OK;
}

HRESULT CMyObject::CImp_clsIMono::GetGratingDispersion(
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	if (NULL == pVarResult) return E_INVALIDARG;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromDouble(this->m_pMyObject->m_pMySciUsbMono->GetGratingDispersion(varg.lVal), pVarResult);
	return S_OK;
}


HRESULT CMyObject::CImp_clsIMono::ScanStart()
{
	DoLogString(L"In ScanStart");
	this->m_pMyObject->m_pMySciUsbMono->ScanStart();
	return S_OK;
}



CMyObject::CImpISciMono::CImpISciMono(CMyObject * pBackObj, IUnknown * punkOuter) :
	m_pMyObject(pBackObj),
	m_punkOuter(punkOuter),
	m_pTypeInfo(NULL),
	m_dispidCurrentWavelength(DISPID_UNKNOWN),
	m_dispidCurrentGrating(DISPID_UNKNOWN),
	m_dispidAutoGrating(DISPID_UNKNOWN),
	m_dispidAmInitialized(DISPID_UNKNOWN),
	m_dispidModel(DISPID_UNKNOWN),
	m_dispidSerialNumber(DISPID_UNKNOWN),
	m_dispidNumGratings(DISPID_UNKNOWN),
	m_dispidAmBusy(DISPID_UNKNOWN),
	m_dispidMonochromatorProperties(DISPID_UNKNOWN),
	m_dispidGratingProperties(DISPID_UNKNOWN),
	m_dispidConfigFile(DISPID_UNKNOWN),
	m_dispidDriveType(DISPID_UNKNOWN),
	m_dispidMonochromatorProperty(DISPID_UNKNOWN),
	m_dispidGratingProperty(DISPID_UNKNOWN),
	m_dispidIsValidPosition(DISPID_UNKNOWN),
	m_dispidGetGratingDispersion(DISPID_UNKNOWN)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;
	TYPEATTR		*	pTypeAttr;

	hr = this->m_pMyObject->GetRefTypeInfo(TEXT("ISciMono"), &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		// store the type info
		this->m_pTypeInfo = pTypeInfo;
		this->m_pTypeInfo->AddRef();
		// get the interface ID
		hr = this->m_pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			this->m_pMyObject->m_iid_ISciMono = pTypeAttr->guid;
			this->m_pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		Utils_GetMemid(this->m_pTypeInfo, L"CurrentWavelength", &m_dispidCurrentWavelength);
		Utils_GetMemid(this->m_pTypeInfo, L"CurrentGrating", &m_dispidCurrentGrating);
		Utils_GetMemid(this->m_pTypeInfo, L"AutoGrating", &m_dispidAutoGrating);
		Utils_GetMemid(this->m_pTypeInfo, L"AmInitialized", &m_dispidAmInitialized);
		Utils_GetMemid(this->m_pTypeInfo, L"Model", &m_dispidModel);
		Utils_GetMemid(this->m_pTypeInfo, L"SerialNumber", &m_dispidSerialNumber);
		Utils_GetMemid(this->m_pTypeInfo, L"NumGratings", &m_dispidNumGratings);
		Utils_GetMemid(this->m_pTypeInfo, L"AmBusy", &m_dispidAmBusy);
		Utils_GetMemid(this->m_pTypeInfo, L"MonochromatorProperties", &m_dispidMonochromatorProperties);
		Utils_GetMemid(this->m_pTypeInfo, L"GratingProperties", &m_dispidGratingProperties);
		Utils_GetMemid(this->m_pTypeInfo, L"ConfigFile", &m_dispidConfigFile);
		Utils_GetMemid(this->m_pTypeInfo, L"DriveType", &m_dispidDriveType);
		Utils_GetMemid(this->m_pTypeInfo, L"MonochromatorProperty", &m_dispidMonochromatorProperty);
		Utils_GetMemid(this->m_pTypeInfo, L"GratingProperty", &m_dispidGratingProperty);
		Utils_GetMemid(this->m_pTypeInfo, L"IsValidPosition", &m_dispidIsValidPosition);
		Utils_GetMemid(this->m_pTypeInfo, L"GetGratingDispersion", &m_dispidGetGratingDispersion);
		pTypeInfo->Release();
	}
}
CMyObject::CImpISciMono::~CImpISciMono()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}
// IUnknown methods
STDMETHODIMP CMyObject::CImpISciMono::QueryInterface(
	REFIID			riid,
	LPVOID		*	ppv)
{
	return this->m_punkOuter->QueryInterface(riid, ppv);
}
STDMETHODIMP_(ULONG) CMyObject::CImpISciMono::AddRef()
{
	return this->m_punkOuter->AddRef();
}
STDMETHODIMP_(ULONG) CMyObject::CImpISciMono::Release()
{
	return this->m_punkOuter->Release();
}
// IDispatch methods
STDMETHODIMP CMyObject::CImpISciMono::GetTypeInfoCount(
	PUINT			pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}
STDMETHODIMP CMyObject::CImpISciMono::GetTypeInfo(
	UINT			iTInfo,
	LCID			lcid,
	ITypeInfo	**	ppTInfo)
{
	if (NULL != this->m_pTypeInfo)
	{
		*ppTInfo = this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
		return S_OK;
	}
	else
	{
		*ppTInfo = NULL;
		return E_FAIL;
	}
}
STDMETHODIMP CMyObject::CImpISciMono::GetIDsOfNames(
	REFIID			riid,
	OLECHAR		**  rgszNames,
	UINT			cNames,
	LCID			lcid,
	DISPID		*	rgDispId)
{
	HRESULT			hr;
	ITypeInfo	*	pTypeInfo;
	hr = this->GetTypeInfo(0, LOCALE_USER_DEFAULT, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}
STDMETHODIMP CMyObject::CImpISciMono::Invoke(
	DISPID			dispIdMember,
	REFIID			riid,
	LCID			lcid,
	WORD			wFlags,
	DISPPARAMS	*	pDispParams,
	VARIANT		*	pVarResult,
	EXCEPINFO	*	pExcepInfo,
	PUINT			puArgErr)
{
	if (dispIdMember == m_dispidCurrentWavelength)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetCurrentWavelength(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetCurrentWavelength(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidCurrentGrating)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetCurrentGrating(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetCurrentGrating(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidAutoGrating)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAutoGrating(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetAutoGrating(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidAmInitialized)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAmInitialized(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetAmInitialized(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidModel)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetModel(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidSerialNumber)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetSerialNumber(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidNumGratings)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetNumGratings(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidAmBusy)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetAmBusy(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidMonochromatorProperties)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetMonochromatorProperties(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidGratingProperties)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetGratingProperties(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidConfigFile)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetConfigFile(pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetConfigFile(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidDriveType)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetDriveType(pVarResult);
		}
	}
	else if (dispIdMember == m_dispidMonochromatorProperty)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetMonochromatorProperty(pDispParams, pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetMonochromatorProperty(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidGratingProperty)
	{
		if (0 != (wFlags & DISPATCH_PROPERTYGET))
		{
			return this->GetGratingProperty(pDispParams, pVarResult);
		}
		else if (0 != (wFlags & DISPATCH_PROPERTYPUT))
		{
			return this->SetGratingProperty(pDispParams);
		}
	}
	else if (dispIdMember == m_dispidIsValidPosition)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->IsValidPosition(pDispParams, pVarResult);
		}
	}
	else if (dispIdMember == m_dispidGetGratingDispersion)
	{
		if (0 != (wFlags & DISPATCH_METHOD))
		{
			return this->GetGratingDispersion(pDispParams, pVarResult);
		}
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMyObject::CImpISciMono::GetCurrentWavelength(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromDouble(this->m_pMyObject->m_pMySciUsbMono->Getposition(), pVarResult);
	return S_OK;
}
HRESULT CMyObject::CImpISciMono::SetCurrentWavelength(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->Setposition(varg.dblVal);
	this->m_pMyObject->FirePropChanged(L"CurrentWavelength");
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetCurrentGrating(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pMyObject->m_pMySciUsbMono->GetcurrentGrating(), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetCurrentGrating(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetcurrentGrating(varg.lVal);
	if (this->m_pMyObject->m_pMySciUsbMono->GetcurrentGrating() == varg.lVal)
	{
		this->m_pMyObject->FirePropChanged(L"CurrentGrating");
	}
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetAutoGrating(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMySciUsbMono->GetautoGrating(), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetAutoGrating(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetautoGrating(VARIANT_TRUE == varg.boolVal);
//	this->m_pMyObject->FirePropChanged(L"AutoGrating");
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetAmInitialized(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(this->m_pMyObject->m_pMySciUsbMono->GetAmOpen(), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetAmInitialized(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	BOOL				init;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BOOL, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	init = VARIANT_TRUE == varg.boolVal;
	this->m_pMyObject->m_pMySciUsbMono->SetAmOpen(init);
	if (this->m_pMyObject->m_pMySciUsbMono->GetAmOpen() == init)
	{
		this->m_pMyObject->FirePropChanged(L"AmInitialized");
	}
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetModel(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	this->m_pMyObject->m_pMySciUsbMono->GetMonoInfo(0, pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetSerialNumber(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	this->m_pMyObject->m_pMySciUsbMono->GetMonoInfo(1, pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetNumGratings(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromInt32(this->m_pMyObject->m_pMySciUsbMono->GetNumberOfGratings(), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetAmBusy(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	InitVariantFromBoolean(FALSE, pVarResult);
	return S_OK;
}

HRESULT	CMyObject::CImpISciMono::GetMonochromatorProperties(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	PCWSTR  rgStrings[] = { L"InputAngle", L"OutputAngle", L"GearTeeth", L"IdleCurrent", L"RunCurrent", L"HighSpeed", L"StepsPerRev" };
	InitVariantFromStringArray(rgStrings, ARRAYSIZE(rgStrings), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetGratingProperties(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	PCWSTR rgStrings[] = { L"Pitch", L"MaxEffectiveWavelength", L"MinEffectiveWavelength", L"Resolution", L"PhaseError",
		L"OffsetFactor", L"LinearFactor", L"QuadFactor", L"ZeroPosition" };
	InitVariantFromStringArray(rgStrings, ARRAYSIZE(rgStrings), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetConfigFile(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	LPTSTR			szConfig = NULL;
	this->m_pMyObject->m_pMySciUsbMono->GetConfigFile(&szConfig);
	if (NULL != szConfig)
	{
		InitVariantFromString(szConfig, pVarResult);
		CoTaskMemFree((LPVOID)szConfig);
	}
//	InitVariantFromString(this->m_pMyObject->m_szConfigFile, pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetConfigFile(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->m_pMyObject->m_pMySciUsbMono->SetConfigFile(varg.bstrVal);
	this->m_pMyObject->FirePropChanged(L"ConfigFile");
	VariantClear(&varg);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetDriveType(VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	// 14 is drive type as string
	VARIANT			Value;
	this->m_pMyObject->m_pMySciUsbMono->GetMonoInfo(14, &Value);
//	{
		VariantChangeType(&Value, &Value, 0, VT_BSTR);
		InitVariantFromInt16(0 == lstrcmpi(Value.bstrVal, L"SineDrive") ? 0 : 1, pVarResult);
		VariantClear(&Value);
//	}
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetMonochromatorProperty(DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	short				index = -1;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	//{L"InputAngle", L"OutputAngle", L"GearTeeth", L"IdleCurrent", L"RunCurrent", L"HighSpeed", L"StepsPerRev"};
	if (0 == lstrcmpi(varg.bstrVal, L"InputAngle"))
	{
		index = 7;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"OutputAngle"))
	{
		index = 8;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"GearTeeth"))
	{
		index = 15;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"IdleCurrent"))
	{
		index = 17;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"RunCurrent"))
	{
		index = 18;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"HighSpeed"))
	{
		index = 19;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"StepsPerRev"))
	{
		index = 20;
	}
	VariantClear(&varg);
	if (index >= 0)
		this->m_pMyObject->m_pMySciUsbMono->GetMonoInfo(index, pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetMonochromatorProperty(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	short				index;
	VARIANT				varCopy;
	if (2 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, DISPID_PROPERTYPUT, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	if (0 == lstrcmpi(varg.bstrVal, L"InputAngle"))
	{
		index = 7;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"OutputAngle"))
	{
		index = 8;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"GearTeeth"))
	{
		index = 15;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"IdleCurrent"))
	{
		index = 17;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"RunCurrent"))
	{
		index = 18;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"HighSpeed"))
	{
		index = 19;
	}
	else if (0 == lstrcmpi(varg.bstrVal, L"StepsPerRev"))
	{
		index = 20;
	}
	if (index >= 0)
	{
		VariantInit(&varCopy);
		VariantCopy(&varCopy, &(pDispParams->rgvarg[0]));
		this->m_pMyObject->m_pMySciUsbMono->SetMonoInfo(index, &varCopy);
		VariantClear(&varCopy);
		this->m_pMyObject->FireMonochromatorPropChanged(varg.bstrVal);
	}
	VariantClear(&varg);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetGratingProperty(DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				gratingID;
	TCHAR				szProperty[MAX_PATH];
	short				index = -1;

	//	PCWSTR rgStrings[] = { L"Pitch", L"MaxEffectiveWavelength", L"MinEffectiveWavelength", L"Resolution", L"PhaseError", L"OffsetFactor", L"LinearFactor", L"QuadFactor", L"ZeroPosition", L"", L"", L"", L"", L"", L"" };

	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	gratingID = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szProperty, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	if (0 == lstrcmpi(szProperty, L"Pitch"))
	{
		index = 0;
	}
	else if (0 == lstrcmpi(szProperty, L"MaxEffectiveWavelength"))
	{
		index = 2;
	}
	else if (0 == lstrcmpi(szProperty, L"MinEffectiveWavelength"))
	{
		index = 3;
	}
	else if (0 == lstrcmpi(szProperty, L"Resolution"))
	{
		index = 7;
	}
	else if (0 == lstrcmpi(szProperty, L"PhaseError"))
	{
		index = 6;
	}
	else if (0 == lstrcmpi(szProperty, L"OffsetFactor"))
	{
		index = 20;
	}
	else if (0 == lstrcmpi(szProperty, L"LinearFactor"))
	{
		index = 21;
	}
	else if (0 == lstrcmpi(szProperty, L"QuadFactor"))
	{
		index = 22;
	}
	else if (0 == lstrcmpi(szProperty, L"ZeroPosition"))
	{
		index = 4;
	}
	if (index >= 0)
	{
		this->m_pMyObject->m_pMySciUsbMono->GetGratingInfo(gratingID, index, pVarResult);
	}

	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::SetGratingProperty(DISPPARAMS* pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	long				gratingID;
	TCHAR				szProperty[MAX_PATH];
	short				index = -1;
	VARIANT				varCopy;

	if (3 != pDispParams->cArgs) return DISP_E_BADPARAMCOUNT;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	gratingID = varg.lVal;
	hr = DispGetParam(pDispParams, 1, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	StringCchCopy(szProperty, MAX_PATH, varg.bstrVal);
	VariantClear(&varg);
	if (0 == lstrcmpi(szProperty, L"Pitch"))
	{
		index = 0;
	}
	else if (0 == lstrcmpi(szProperty, L"MaxEffectiveWavelength"))
	{
		index = 2;
	}
	else if (0 == lstrcmpi(szProperty, L"MinEffectiveWavelength"))
	{
		index = 3;
	}
	else if (0 == lstrcmpi(szProperty, L"Resolution"))
	{
		index = 7;
	}
	else if (0 == lstrcmpi(szProperty, L"PhaseError"))
	{
		index = 6;
	}
	else if (0 == lstrcmpi(szProperty, L"OffsetFactor"))
	{
		index = 20;
	}
	else if (0 == lstrcmpi(szProperty, L"LinearFactor"))
	{
		index = 21;
	}
	else if (0 == lstrcmpi(szProperty, L"QuadFactor"))
	{
		index = 22;
	}
	else if (0 == lstrcmpi(szProperty, L"ZeroPosition"))
	{
		index = 4;
	}
	if (index >= 0)
	{
		VariantInit(&varCopy);
		VariantCopy(&varCopy, &(pDispParams->rgvarg[0]));
		this->m_pMyObject->m_pMySciUsbMono->SetGratingInfo(gratingID, index, &varCopy);
		VariantClear(&varCopy);
		this->m_pMyObject->FireGratingPropChanged(gratingID, szProperty);
	}
	return S_OK;
}

HRESULT	CMyObject::CImpISciMono::IsValidPosition(DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_R8, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromBoolean(this->m_pMyObject->m_pMySciUsbMono->IsValidPosition(varg.dblVal), pVarResult);
	return S_OK;
}
HRESULT	CMyObject::CImpISciMono::GetGratingDispersion(DISPPARAMS* pDispParams, VARIANT* pVarResult)
{
	if (NULL == pVarResult) return E_INVALIDARG;
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	double				dispersion = 5.0;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	InitVariantFromDouble(dispersion, pVarResult);
	return S_OK;
}
