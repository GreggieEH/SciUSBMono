#pragma once

class CMyObject;

// structure sent to the subclass procedures
class CMyWindow;
struct SUBCLASS_DATA
{
	CMyWindow		*	pMyWindow;
	WNDPROC				wpOrig;			// original window procedure
	UINT				nControlID;		// edit box control ID
};
// struct

class CMyWindow
{
public:
	CMyWindow(CMyObject * pMyObject);
	~CMyWindow(void);
	BOOL				MyTranslateAccelerator(
							LPMSG			lpmsg);
	HWND				GetMyWindow();
	void				SetHostNames(
							LPCTSTR			szContainerApp,  
							LPCTSTR			szContainerObj);
	HWND				CreateInPlaceWindow(
							HINSTANCE		hInst,
							HWND			hwndParent,
							LPCRECT			prcPos);
	void				SetObjectRects(
							LPCRECT			prcPos, 
							LPCRECT			prcClip);
	// get the default window extent
	void				GetDefaultExtent(
							long		*	pX,
							long		*	pY);
	// create a font
	HFONT				myGetFont();
protected:
	BOOL				doRegWindowClass();				// register the window class
	// map dialog units to pixels
	void				DialogUnitsToPixels(
							HDC				hdc,
							LPRECT			prc);
	// create a child window
	HWND				CreateChildWindow(
							LPCTSTR			szWindowClass,
							UINT			nID,
							LPCRECT			prcDialogUnits,
							LPCTSTR			szName,
							DWORD			dwStyle,
							DWORD			dwExtendedStyle);
	LRESULT				WinProc(
							UINT			uMsg,
							WPARAM			wParam,
							LPARAM			lParam);
	void				OnCreate();				// handler for WM_CREATE
	// display the current settings
	void				DisplayModelAndSerialNumber();
	BOOL				GetOurObject(
							IDispatch	**	ppdisp);
	// get the model and the serial number
	void				GetModelAndSerialNumber(
							LPTSTR		*	szModel,
							LPTSTR		*	szSerialNumber);
	// am initialized flag
	BOOL				GetAmInitialized();
	// display the current position
	void				DisplayCurrentPosition();
	// get the wavelength range
	void				GetWavelengthRange(
							double		*	minWave,
							double		*	maxWave);
	void				SetTrackbarRange();
	BOOL				OnCommand(
							WORD			wmID,
							WORD			wmEvent);
	BOOL				OnNotify(
							LPNMHDR			pnmh);
	// Horizontal scroll bar
	BOOL				OnHScroll(
							HWND			hwndScroll,
							WPARAM			wParam);
	// connect and disconnect the sinks
	void				ConnectSinks();
	void				DisconnectSinks();
	// custom sink events
	void				OnError(
							LPCTSTR			szError);
	void				OnStatusMessage(
							LPCTSTR			szStatus,
							BOOL			fAmBusy);
	// property change notifications
	BOOL				OnPropRequestEdit(
							DISPID			dispid);
	void				OnPropChanged(
							DISPID			dispid);
	// edit box subclass
	void				SubclassEditBox(
							UINT			uEditBox);
	void				OnReturnClicked(
							UINT			uEditBox);
	// status control subclass
	void				SubclassStatusBar(
							UINT			uStatus);
	void				OnPaintStatusBar(
							UINT			uStatus);
	// attempt initialization
	void				InitMono();
private:
	CMyObject		*	m_pMyObject;
	HWND				m_hwnd;
	LPTSTR				m_szLastError;			// last error string
	BOOL				m_fHaveError;			// have error flag
	DWORD				m_dwCookie;				// sink connection cookie
	DWORD				m_dwPropNotifyCookie;	// property notify cookie
	LPTSTR				m_szCurrentStatus;
	BOOL				m_fAmBusy;				// busy flag

	// custom sink
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CMyWindow * pMyWindow);
		~CImpISink();
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
	private:
		CMyWindow			*	m_pMyWindow;
		ULONG					m_cRefs;
	};
	// property notification
	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CMyWindow * pMyWindow);
		~CImpIPropNotify();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotify sink methods
		STDMETHODIMP			OnChanged( 
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit( 
									DISPID			dispID);
	private:
		CMyWindow			*	m_pMyWindow;
		ULONG					m_cRefs;
	};
	friend CImpISink;
	friend CImpIPropNotify;

// window procedure
friend LRESULT CALLBACK	WindowProcSciUsbMono(HWND, UINT, WPARAM, LPARAM);
// edit box subclass
friend LRESULT CALLBACK SubclassProcEditBox(HWND, UINT, WPARAM, LPARAM);
// status window subclass
friend LRESULT CALLBACK SubclassProcStatusWindow(HWND, UINT, WPARAM, LPARAM);
};
