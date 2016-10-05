#pragma once

class CMyPropPage : public IPropertyPage
{
public:
	CMyPropPage(void);
	~CMyPropPage(void);
	// IUnknown methods
	STDMETHODIMP			QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();
	// IPropertyPage methods
	STDMETHODIMP			SetPageSite( 
								IPropertyPageSite *pPageSite);
	STDMETHODIMP			Activate( 
								HWND			hWndParent ,  //Parent window handle
								LPCRECT			prc ,      //Pointer to RECT structure
								BOOL			bModal);
	STDMETHODIMP			Deactivate(void);
	STDMETHODIMP			GetPageInfo( 
								PROPPAGEINFO *	pPageInfo);
	STDMETHODIMP			SetObjects( 
								ULONG			cObjects ,  //Number of IUnknown pointers in the ppUnk array
								IUnknown	**	ppUnk);
	STDMETHODIMP			Show(
								UINT			nCmdShow);
	STDMETHODIMP			Move( 
								LPCRECT			prc);
	STDMETHODIMP			IsPageDirty(void);
	STDMETHODIMP			Apply(void);
	STDMETHODIMP			Help( 
								LPCOLESTR		pszHelpDir);
	STDMETHODIMP			TranslateAccelerator( 
								LPMSG			pMsg);
protected:
	// map dialog units to pixels
	void					DialogUnitsToPixels(
								HDC				hdc,
								LPRECT			prc);
	// set the dirty flag
	void					SetDirty(
								BOOL			fDirty);
	// dialog procedure
	BOOL					DlgProc(
								UINT			uMsg,
								WPARAM			wParam,
								LPARAM			lParam);
	BOOL					OnInitDialog();		// handler for WM_INITDIALOG
	BOOL					OnCommand(			// command handler
								WORD			wmID,
								WORD			wmEvent);
	BOOL					OnNotify(
								LPNMHDR			pnmh);
	void					OnEditBoxReturnClicked(
								UINT			nEditBox);
	// subclass an edit box
	void					SubclassEditBox(
								UINT			uControlID);
	// display values
	void					DisplayModel();
	void					DisplaySerialNumber();
//	void					DisplayIdleCurrent();
//	void					DisplayRunCurrent();
//	void					DisplayHighSpeed();
	void					DisplayStepsPerRev();
	void					DisplayCurrentPosition();
	BOOL					GetMonoInfo(
								long			Index,
								VARIANT		*	MonoInfo);
	void					SetMonoInfo(
								long			Index,
								VARIANT		*	MonoInfo);
	void					DisplayNumberOfGratings();
	long					GetNumberOfGratings();
	void					DisplayCurrentGrating();
	long					GetCurrentGrating();
	void					DisplayAutoGrating();
	void					OnClickedAutoGrating();
	// push buttons
	void					OnSetPosition();
	void					OnAdjustZeroPosition();
	// apply values
//	void					ApplyIdleCurrent();
//	void					ApplyRunCurrent();
//	void					ApplyHighSpeed();
	void					ApplyStepsPerRev();
	// display am initialized
	void					DisplayAmInitialized();
	void					OnClickedAmInitialized();
	// property change notification
	void					OnPropChanged(
								DISPID			dispid);
	BOOL					CanAutoSelect();
	// display grating info
	void					DisplayGratingInfo();
	// get grating info
	void					GetGratingInfo(
								long			grating,
								long			Index,
								VARIANT		*	GratingInfo);
	// get object name
	BOOL					GetObjectName(
								IDispatch	*	pdisp,
								LPTSTR		*	szName);
private:
	ULONG					m_cRefs;			// object reference count
	HWND					m_hwndDlg;			// property page window handle
	IPropertyPageSite	*	m_pPropertyPageSite;	// property page site object
	IDispatch			*	m_pdisp;			// object which we are communicating
	BOOL					m_fDirty;			// dirty flag
	BOOL					m_fAmInitializing;	// flag handling OnInitDialog
	// property notify sink
	DWORD					m_dwPropNotifyCookie;
	LPTSTR					m_szName;				// name for this object

// property page procedure
friend LRESULT CALLBACK MyPropertyPageProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
// edit box subclass
friend LRESULT CALLBACK MyEditBoxSubclass(HWND, UINT, WPARAM, LPARAM);

// structure sent in with an edit box
struct MY_SUBCLASS_STRUCT
{
	CMyPropPage			*	pMyPropPage;
	WNDPROC					wpOrig;
};

	class CImpIPropNotify : public IPropertyNotifySink
	{
	public:
		CImpIPropNotify(CMyPropPage * pMyPropPage);
		~CImpIPropNotify();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IPropertyNotifySink methods
		STDMETHODIMP			OnChanged( 
									DISPID			dispID);
		STDMETHODIMP			OnRequestEdit( 
									DISPID			dispID);
	private:
		CMyPropPage			*	m_pMyPropPage;
		ULONG					m_cRefs;
	};
	friend CImpIPropNotify;

};