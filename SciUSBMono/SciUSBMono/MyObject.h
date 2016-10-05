#pragma once

class CMyWindow;
class CMySciUsbMono;

class CMyObject : public IUnknown
{
public:
	CMyObject(IUnknown * pUnkOuter);
	~CMyObject(void);
	// IUnknown methods
	STDMETHODIMP				QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)		AddRef();
	STDMETHODIMP_(ULONG)		Release();
	// initialization
	HRESULT						Init();
	// sink events
	HWND						FireRequestMainWindow();
	void						FireError(
									LPCTSTR			Error);
	void						FireStatusMessage(
									LPCTSTR			Status,
									BOOL			fAmBusy);
	void						FireHaveNewPosition(
									double			position);
	BOOL						FireRequestChangeGrating(
									long			grating);
	void						FireChangedGrating(
									long			grating);
	void						FireBusyStatusChange(
									BOOL			fBusy);
	BOOL						FireQueryAllowChangePosition(
									double			newPosition);
	BOOL						FireQueryAllowChangeZeroOffset();
	void						FireRapidScanStepped(
									double			wavelength);
	void						FireRapidScanEnded(
									BOOL			fSuccess);
	void						FireAmInitPropChanged(
									BOOL			amInit);
	void						FireAutoGratingPropChanged(
									BOOL			autoGrating);
	// get our control window
	HWND						GetControlWindow();
	// successful initialization
	void						OnHaveInitialized();
	// get the name for this control - given by the client
	HRESULT						GetName(
									LPTSTR			*	pszName);
protected:
	HRESULT						GetClassInfo(
									ITypeInfo	**	ppTI);
	HRESULT						GetRefTypeInfo(
									LPCTSTR			szInterface,
									ITypeInfo	**	ppTypeInfo);
	HRESULT						Init__clsIMono();
	// get the client site
	HRESULT						GetClientSite(
									IOleClientSite**	ppClientSite);
	// get the extended control
	HRESULT						GetExtendedControl(
									IDispatch		**	ppdispExtended);
	// send on data change
	void						SendOnDataChange(
									BOOL				fDataOnStop);
	// get the container
	HRESULT						GetContainer(
									IOleContainer	**	ppContainer);
	// get the InPlaceFrame object
	HRESULT						GetInPlaceFrame(
									IOleInPlaceFrame**	ppInPlaceFrame);
	// enable or disable modeless dialogs in the client
	void						EnableModeless(
									BOOL				fEnable);
	// set this as an active object
	void						SetActiveObject();
	// get the control site
	HRESULT						GetControlSite(
									IOleControlSite**	ppControlSite);
	// get the ambient object
	HRESULT						GetAmbientObject(
									IDispatch		**	ppdispAmbient);
	// get an ambient property
	HRESULT						GetAmbientProperty(
									DISPID				dispid,
									VARIANT			*	pValue);
private:
	class CImpIDispatch : public IDispatch
	{
	public:
		CImpIDispatch(CMyObject * pMyObject, IUnknown * punkOuter);
		~CImpIDispatch();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		void					GetHandler(
									IDispatch	*	pdispHandler,
									VARIANT		*	pVarResult);
		HRESULT					SetHandler(
									DISPPARAMS	*	pDispParams,
									IDispatch	**	ppdispHandler);
		void					SetPosition(
									double			position);
		HRESULT					ConvertSteps(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ConvertNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetRapidScanRunning(
									VARIANT		*	pVarResult);
		HRESULT					SetRapidScanRunning(
									DISPPARAMS	*	pDispParams);
		HRESULT					StartRapidScan(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					GetPageSelected(
									VARIANT		*	pVarResult);
		HRESULT					SetPageSelected(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetBacklashSteps(
									VARIANT		*	pVarResult);
		HRESULT					SetBacklashSteps(
									DISPPARAMS	*	pDispParams);
	private:
		CMyObject			*	m_pMyObject;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
	};
	class CImpIProvideClassInfo2 : public IProvideClassInfo2
	{
	public:
		CImpIProvideClassInfo2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIProvideClassInfo2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IProvideClassInfo method
		STDMETHODIMP			GetClassInfo(
									ITypeInfo	**	ppTI);  
		// IProvideClassInfo2 method
		STDMETHODIMP			GetGUID(
									DWORD			dwGuidKind,  //Desired GUID
									GUID		*	pGUID);       
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIConnectionPointContainer : public IConnectionPointContainer
	{
	public:
		CImpIConnectionPointContainer(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIConnectionPointContainer();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IConnectionPointContainer methods
		STDMETHODIMP			EnumConnectionPoints(
									IEnumConnectionPoints **ppEnum);
		STDMETHODIMP			FindConnectionPoint(
									REFIID			riid,  //Requested connection point's interface identifier
									IConnectionPoint **ppCP);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIPersistStorage : public IPersistStorage
	{
	public:
		CImpIPersistStorage(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistStorage();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID			*	pClassID);
		// IPersistStorage methods
		STDMETHODIMP			IsDirty(void);
		STDMETHODIMP			InitNew(
									IStorage		*	pStg);
		STDMETHODIMP			Load(
									IStorage		*	pStg);
		STDMETHODIMP			Save(
									IStorage		*	pStgSave,   
									BOOL				fSameAsLoad);
		STDMETHODIMP			SaveCompleted(
									IStorage		*	pStgNew);
		STDMETHODIMP			HandsOffStorage(void);
	protected:
		HRESULT					SaveToStream(
									IStream			*	pStm,
									BOOL				fClearDirty);
		HRESULT					CreateNewStream(
									IStorage		*	pStg,
									IStream			**	ppStm);
		HRESULT					OpenStream(
									IStorage		*	pStg,
									IStream			**	ppStm);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		BOOL					m_fNoScribble;		// no scribble mode
	};
	class CImpIPersistStreamInit : public IPersistStreamInit
	{
	public:
		CImpIPersistStreamInit(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIPersistStreamInit();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID			*	pClassID);
		// IPersistStreamInit methods
		STDMETHODIMP			IsDirty(void);
		STDMETHODIMP			Load(
									LPSTREAM			pStm);
		STDMETHODIMP			Save(
									LPSTREAM			pStm ,  
									BOOL				fClearDirty);
		STDMETHODIMP			GetSizeMax(
									ULARGE_INTEGER	*	pcbSize);
		STDMETHODIMP			InitNew(void);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIPersistPropertyBag : public IPersistPropertyBag
	{
	public:
		CImpIPersistPropertyBag(CMyObject * pMyObject, IUnknown * punkOuter);
		~CImpIPersistPropertyBag();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPersist method
		STDMETHODIMP			GetClassID(
									CLSID			*	pClassID);
		// IPersistPropertyBag methods
		STDMETHODIMP			InitNew();
		STDMETHODIMP			Load(          
									IPropertyBag	*	pPropBag,
									IErrorLog		*	pErrorLog);
		STDMETHODIMP			Save(
									IPropertyBag	*	pPropBag,
									BOOL				fClearDirty,
									BOOL				fSaveAllProperties);
	private:
		CMyObject			*	m_pMyObject;
		IUnknown			*	m_punkOuter;
	};
	class CImpIDataObject : public IDataObject
	{
	public:
		CImpIDataObject(CMyObject * pMyObject, IUnknown * punkOuter);
		~CImpIDataObject();
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDataObject methods
		STDMETHODIMP			GetData(
									FORMATETC	*	pFormatetc,  //Pointer to the FORMATETC structure
									STGMEDIUM	*	pmedium);
		STDMETHODIMP			GetDataHere(
									FORMATETC	*	pFormatetc,  //Pointer to the FORMATETC structure
									STGMEDIUM	*	pmedium);
		STDMETHODIMP			QueryGetData(
									FORMATETC	*	pFormatetc);
		STDMETHODIMP			GetCanonicalFormatEtc(
									FORMATETC	*	pFormatetcIn,  
									FORMATETC	*	pFormatetcOut);
		STDMETHODIMP			SetData(
									FORMATETC	*	pFormatetc,  
									STGMEDIUM	*	pmedium,     
									BOOL			fRelease);
		STDMETHODIMP			EnumFormatEtc(
									DWORD			dwDirection,  
									IEnumFORMATETC** ppenumFormatetc);
		STDMETHODIMP			DAdvise(
									FORMATETC	*	pFormatetc,
									DWORD			advf,
									IAdviseSink	*	pAdvSink,
									DWORD		*	pdwConnection);
		STDMETHODIMP			DUnadvise(
									DWORD			dwConnection);
		STDMETHODIMP			EnumDAdvise(
									IEnumSTATDATA** ppenumAdvise);  
	protected:
		// set data
		HRESULT					SetDataFromHGlobal(
									HGLOBAL				hglobal);
		HRESULT					SetDataFromFile(
									LPCTSTR				szFile);
		HRESULT					SetDataFromStream(
									IStream			*	pStm);
		HRESULT					SetDataFromBitmap(
									HBITMAP				hbitmap);
		// get data
		HRESULT					GetDataToHGlobal(
									HGLOBAL			*	phglobal);
		HRESULT					GetDataToFile(
									LPCTSTR				szFile);
		HRESULT					GetDataToStream(
									IStream			**	ppStm);
		HRESULT					GetDataToBitmap(
									HBITMAP			*	phbitmap);
		// write data to a stream
		HRESULT					SaveToStream(
									IStream			*	pStm);
	private:
		CMyObject			*	m_pMyObject;
		IUnknown			*	m_punkOuter;
	};
	class CImpIOleObject : public IOleObject
	{
	public:
		CImpIOleObject(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIOleObject();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleObject methods
		STDMETHODIMP			SetClientSite(
									IOleClientSite	*	pClientSite);
		STDMETHODIMP			GetClientSite(
									IOleClientSite	**	ppClientSite);
		STDMETHODIMP			SetHostNames(
									LPCOLESTR			szContainerApp,  
									LPCOLESTR			szContainerObj);
		STDMETHODIMP			Close(
									DWORD				dwSaveOption);
		STDMETHODIMP			SetMoniker(
									DWORD				dwWhichMoniker,  
									IMoniker		*	pmk);
		STDMETHODIMP			GetMoniker(
									DWORD				dwAssign,  
									DWORD				dwWhichMoniker,            
									IMoniker		**	ppmk);
		STDMETHODIMP			InitFromData(
									IDataObject		*	pDataObject,  
									BOOL				fCreation,            
									DWORD				dwReserved);
		STDMETHODIMP			GetClipboardData(
									DWORD				dwReserved,  
									IDataObject		**	ppDataObject);
		STDMETHODIMP			DoVerb(
									LONG				iVerb,          
									LPMSG				lpmsg,         
									IOleClientSite	*	pActiveSite,                      
									LONG				lindex,         
									HWND				hwndParent,     
									LPCRECT				lprcPosRect);
		STDMETHODIMP			EnumVerbs(
									IEnumOLEVERB	**	ppEnumOLEVERB);
		STDMETHODIMP			Update();
		STDMETHODIMP			IsUpToDate();
		STDMETHODIMP			GetUserClassID(
									CLSID			*	pClsid);
		STDMETHODIMP			GetUserType(
									DWORD				dwFormOfType,  
									LPOLESTR		*	pszUserType);
		STDMETHODIMP			SetExtent(
									DWORD				dwDrawAspect,  
									SIZEL			*	psizel);
		STDMETHODIMP			GetExtent(
									DWORD				dwDrawAspect,  
									SIZEL			*	psizel);
		STDMETHODIMP			Advise(
									IAdviseSink		*	pAdvSink,  
									DWORD			*	pdwConnection);
		STDMETHODIMP			Unadvise(
									DWORD				dwConnection);
		STDMETHODIMP			EnumAdvise(
									IEnumSTATDATA	**	ppenumAdvise);
		STDMETHODIMP			GetMiscStatus(
									DWORD				dwAspect,  
									DWORD			*	pdwStatus);
		STDMETHODIMP			SetColorScheme(
									LOGPALETTE		*	pLogpal);
	protected:
		// helpers for DoVerb
		HRESULT					DoVerbPrimary(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbShow(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbInPlaceActivate(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbUIActivate(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbHide(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbOpen(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbDiscardUndo(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					DoVerbProperties(
									LPCRECT				prcPosRect, 
									HWND				hwndParent);
		HRESULT					InPlaceActivate(
									LONG				iVerb, 
									const RECT		*	prcPosRect);
		BOOL					DoesVerbUIActivate(
									LONG				iVerb);
		BOOL					SetControlFocus(
									IOleInPlaceSite	*	pIOleInPlaceSite,
									BOOL				bGrab);
		// copy data objects
		HRESULT					DoCopyData(
									IDataObject		*	pDataSource,
									IDataObject		*	pDataDest);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		IOleAdviseHolder	*	m_pAdviseHolder;
		IOleClientSite		*	m_pClientSite;
		ULONG					m_crefs;
	};
	class CImpIOleControl : public IOleControl
	{
	public:
		CImpIOleControl(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIOleControl();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleControl methods
		STDMETHODIMP			GetControlInfo(
									CONTROLINFO		*	pCI);
		STDMETHODIMP			OnMnemonic(
									LPMSG				pMsg);
		STDMETHODIMP			OnAmbientPropertyChange(
									DISPID				dispID);
		STDMETHODIMP			FreezeEvents(
									BOOL				bFreeze);
	protected:
		// ambient property changes
		HRESULT					OnAmbientChangeFont(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeLCID(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeUIDead(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeShowGrabHandles(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeShowHatching(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeUserMode(
									IDispatch		*	pdispAmbient);
		HRESULT					OnAmbientChangeSupportsMnemonics(
									IDispatch		*	pdispAmbient);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIOleInPlaceObject : public IOleInPlaceObject
	{
	public:
		CImpIOleInPlaceObject(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIOleInPlaceObject();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleWindow methods
		STDMETHODIMP			GetWindow(
									HWND			*	phwnd);
		STDMETHODIMP			ContextSensitiveHelp(
									BOOL				fEnterMode);
		// IOleInPlaceObject methods
		STDMETHODIMP			InPlaceDeactivate();
		STDMETHODIMP			UIDeactivate();
		STDMETHODIMP			SetObjectRects(
									LPCRECT				lprcPosRect,  
									LPCRECT				lprcClipRect);
		STDMETHODIMP			ReactivateAndUndo();
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIOleInPlaceActiveObject : public IOleInPlaceActiveObject
	{
	public:
		CImpIOleInPlaceActiveObject(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIOleInPlaceActiveObject();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IOleWindow methods
		STDMETHODIMP			GetWindow(
									HWND			*	phwnd);
		STDMETHODIMP			ContextSensitiveHelp(
									BOOL				fEnterMode);
		// IOleInPlaceActiveObject methods
		STDMETHODIMP			TranslateAccelerator(
									LPMSG				lpmsg);
		STDMETHODIMP			OnFrameWindowActivate(
									BOOL				fActivate);
		STDMETHODIMP			OnDocWindowActivate(
									BOOL				fActivate);
		STDMETHODIMP			ResizeBorder(
									LPCRECT				prcBorder,              
									IOleInPlaceUIWindow *pUIWindow,                                  
									BOOL				fFrameWindow);
		STDMETHODIMP			EnableModeless(
									BOOL				fEnable);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpISpecifyPropertyPages : public ISpecifyPropertyPages 
	{
	public:
		CImpISpecifyPropertyPages(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpISpecifyPropertyPages();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// ISpecifyPropertyPages method
		STDMETHODIMP			GetPages(
									CAUUID			*	pPages);
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
	};
	class CImpIViewObject2 : public IViewObject2
	{
	public:
		CImpIViewObject2(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImpIViewObject2();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IViewObject methods
		STDMETHODIMP			Draw(
									DWORD				dwAspect,   //Aspect to be drawn
									LONG				lindex,     
									void *				pvAspect,  //Pointer to DVASPECTINFO structure or NULL
									DVTARGETDEVICE *	ptd,
									HDC					hicTargetDev, //Information context for the target device
									HDC					hdcDraw,      //Device context on which to draw
									LPCRECTL			lprcBounds,
									LPCRECTL			lprcWBounds,
									BOOL (__stdcall *	pfnContinue) (DWORD),
									DWORD				dwContinue);
		STDMETHODIMP			GetColorSet(
									DWORD				dwAspect,   
									LONG				lindex,      
									void			*	pvAspect,  
									DVTARGETDEVICE	*	ptd,
									HDC					hicTargetDev, 
									LOGPALETTE		**	ppColorSet);
		STDMETHODIMP			Freeze(
									DWORD				dwAspect,   
									LONG				lindex,      
									void			*	pvAspect,  
									DWORD			*	pdwFreeze);
		STDMETHODIMP			Unfreeze(
									DWORD				dwFreeze);
		STDMETHODIMP			SetAdvise(
									DWORD				dwAspect,  
									DWORD				advf,      
									IAdviseSink		*	pAdvSink);
		STDMETHODIMP			GetAdvise(
									DWORD			*	pdwAspect,  
									DWORD			*	padvf,      
									IAdviseSink		**	ppAdvSink);
		// IViewObject2 method
		STDMETHODIMP			GetExtent(
									DWORD				dwAspect,  
									LONG				lindex,    //Part of the object to draw
									DVTARGETDEVICE*		ptd,
									LPSIZEL				lpsizel);
	protected:
		void					FireViewChange();
	private:
		CMyObject			*	m_pBackObj;
		IUnknown			*	m_punkOuter;
		IAdviseSink			*	m_pIAdviseSink;
		DWORD					m_advf;
	};
	class CImp_clsIMono : public IDispatch
	{
	public:
		CImp_clsIMono(CMyObject * pBackObj, IUnknown * punkOuter);
		~CImp_clsIMono();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					MyInvoke(
									DISPID			dispIdMember,      
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult);
		HRESULT					ConvertStepsToNM(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					ReadConfig(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					WriteConfig(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetWaitForComplete(
									VARIANT		*	pVarResult);
		HRESULT					SetWaitForComplete(
									DISPPARAMS	*	pDispParams);
		HRESULT					GetGratingDispersion(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
	private:
		CMyObject			*	m_pMyObject;
		IUnknown			*	m_punkOuter;
		ITypeInfo			*	m_pTypeInfo;
		DISPID					m_dispidCurrentWavelength;
		DISPID					m_dispidAutoGrating;
		DISPID					m_dispidCurrentGrating;
		DISPID					m_dispidAmBusy;
		DISPID					m_dispidAmInitialized;
		DISPID					m_dispidGratingParams;
		DISPID					m_dispidMonoParams;
		DISPID					m_dispidIsValidPosition;
		DISPID					m_dispidConvertStepsToNM;
		DISPID					m_dispidDisplayConfigValues;
		DISPID					m_dispidDoSetup;
		DISPID					m_dispidReadConfig;
		DISPID					m_dispidWriteConfig;
		DISPID					m_dispidWaitForComplete;
		DISPID					m_dispidGetGratingDispersion;
	};
	// make the nested classes friends
	friend CImpIDispatch;
	friend CImpIProvideClassInfo2;
	friend CImpIConnectionPointContainer;
	friend CImpIPersistStorage;
	friend CImpIPersistStreamInit;
	friend CImpIPersistPropertyBag;
	friend CImpIDataObject;
	friend CImpIOleObject;
	friend CImpIOleControl;
	friend CImpIOleInPlaceObject;
	friend CImpIOleInPlaceActiveObject;
	friend CImpISpecifyPropertyPages;
	friend CImpIViewObject2;
	friend CImp_clsIMono;
	// data members
	// implementation classes
	CImpIDispatch				*	m_pImpIDispatch;
	CImpIProvideClassInfo2		*	m_pImpIProvideClassInfo2;
	CImpIConnectionPointContainer*	m_pImpIConnectionPointContainer;
	CImpIPersistStorage			*	m_pImpIPersistStorage;
	CImpIPersistStreamInit		*	m_pImpIPersistStreamInit;
	CImpIPersistPropertyBag		*	m_pImpIPersistPropertyBag;
	CImpIDataObject				*	m_pImpIDataObject;
	CImpIOleObject				*	m_pImpIOleObject;
	CImpIOleControl				*	m_pImpIOleControl;
	CImpIOleInPlaceObject		*	m_pImpIOleInPlaceObject;
	CImpIOleInPlaceActiveObject	*	m_pImpIOleInPlaceActiveObject;
	CImpISpecifyPropertyPages	*	m_pImpISpecifyPropertyPages;
	CImpIViewObject2			*	m_pImpIViewObject2;
	CImp_clsIMono				*	m_pImp_clsIMono;
	// object reference count
	ULONG							m_cRefs;
	// outer unknown for aggregation
	IUnknown					*	m_pUnkOuter;
	// connection point array
	IConnectionPoint			*	m_paConnPts[MAX_CONN_PTS];
	// flags
	BOOL							m_fAmLoaded;			// am loaded flag
	BOOL							m_fFreezeEvents;		// freeze events flag
	// data advise holder
	IDataAdviseHolder			*	m_pDataAdviseHolder;
	// window manager
	CMyWindow					*	m_pMyWindow;
	// storage objects
	IStorage					*	m_pStg;				// storage object
	IStream						*	m_pStm;				// stream object
	// monochromator manager
	CMySciUsbMono				*	m_pMySciUsbMono;
	// event message handlers
	IDispatch					*	m_pdispOnError;
	IDispatch					*	m_pdispOnStatusMessage;
	IDispatch					*	m_pdispOnHaveNewPosition;
	IDispatch					*	m_pdispOnChangedGrating;
	IDispatch					*	m_pdispOnAmInitialized;
	// _clsIMono interface id
	IID								m_iid_clsIMono;
	// __clsIMono events
	IID								m_iid__clsIMono;
	DISPID							m_dispidGratingChanged;
	DISPID							m_dispidBeforeMoveChange;
	DISPID							m_dispidHaveAborted;
	DISPID							m_dispidMoveCompleted;
	DISPID							m_dispidMoveError;
	DISPID							m_dispidRequestChangeGrating;
	DISPID							m_dispidRequestParentWindow;
	DISPID							m_dispidStatusMessage;
	DISPID							m_dispidAmInitPropChanged;
	DISPID							m_dispidAutoGratingPropChanged;
	// page selected flag
	BOOL							m_fPageSelected;
};
