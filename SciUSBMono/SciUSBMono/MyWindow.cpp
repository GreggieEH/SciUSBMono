#include "stdafx.h"
#include "MyWindow.h"
#include "MyObject.h"
#include "dispids.h"
#include "MyGuids.h"

CMyWindow::CMyWindow(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_hwnd(NULL),
	m_szLastError(NULL),			// last error string
	m_fHaveError(FALSE),			// have error flag
	m_dwCookie(0),					// sink connection cookie
	m_dwPropNotifyCookie(0),		// property notify cookie
	m_szCurrentStatus(NULL),
	m_fAmBusy(FALSE)				// busy flag
{
//	Utils_DoCopyString(&m_szCurrentStatus, TEXT("Not Init"));
	SHStrDup(L"Not Init", &m_szCurrentStatus);
}

CMyWindow::~CMyWindow(void)
{
	// make sure that the sinks are disconnected
	this->DisconnectSinks();
	if (NULL != this->m_szLastError)
	{
		CoTaskMemFree((LPVOID)this->m_szLastError);
		this->m_szLastError = NULL;
	}
//	Utils_DoCopyString(&m_szLastError, NULL);
//	Utils_DoCopyString(&m_szCurrentStatus, NULL);
	if (NULL != this->m_szCurrentStatus)
	{
		CoTaskMemFree((LPVOID)this->m_szCurrentStatus);
		this->m_szCurrentStatus = NULL;
	}
}

BOOL CMyWindow::MyTranslateAccelerator(
							LPMSG			lpmsg)
{
	return FALSE;
}

HWND CMyWindow::GetMyWindow()
{
	return this->m_hwnd;
}

void CMyWindow::SetHostNames(
							LPCTSTR			szContainerApp,  
							LPCTSTR			szContainerObj)
{
}

HWND CMyWindow::CreateInPlaceWindow(
							HINSTANCE		hInst,
							HWND			hwndParent,
							LPCRECT			prcPos)
{
	// register the window
	if (!this->doRegWindowClass()) return NULL;
	if (NULL != this->m_hwnd) return this->m_hwnd;
	HWND				hwnd;
	hwnd = CreateWindowEx(
		0,
		MY_WINDOW_CLASS,
		TEXT(""),
		WS_VISIBLE | WS_CHILD,
		prcPos->left, prcPos->top,
		prcPos->right - prcPos->left,
		prcPos->bottom - prcPos->top,
		hwndParent,
		(HMENU) 1100,
		GetOurInstance(),
		(LPVOID) this);
	return hwnd;
}

void CMyWindow::SetObjectRects(
							LPCRECT			prcPos, 
							LPCRECT			prcClip)
{
	if (NULL != this->m_hwnd)
	{
		MoveWindow(this->m_hwnd, prcPos->left, prcPos->top,
			prcPos->right - prcPos->left,
			prcPos->bottom - prcPos->top,
			TRUE);
	}
}

// get the default window extent
void CMyWindow::GetDefaultExtent(
							long		*	pX,
							long		*	pY)
{
	HDC					hdc		= GetDC(NULL);			// device context for screen
	RECT				rc;

	SetRectEmpty(&rc);
	rc.right	= DEF_X_EXTENT;
	rc.bottom	= DEF_Y_EXTENT;
	this->DialogUnitsToPixels(hdc, &rc);
	*pX			= rc.right;
	*pY			= rc.bottom;
	ReleaseDC(NULL, hdc);
}

// create a font
HFONT CMyWindow::myGetFont()
{
	return NULL;
}

// map dialog units to pixels
void CMyWindow::DialogUnitsToPixels(
							HDC				hdc,
							LPRECT			prc)
{
	TEXTMETRIC				txtMetric;
	long					baseunitX;
	long					baseunitY;
	RECT					rcCopy;

	// get the text size
	GetTextMetrics(hdc, &txtMetric);
	baseunitX	= txtMetric.tmAveCharWidth;
	baseunitY	= txtMetric.tmHeight;
	// copy the input rectangle
	CopyRect(&rcCopy, prc);
	// determine the output dimensions
	prc->left	= MulDiv(rcCopy.left, baseunitX, 4);
	prc->top	= MulDiv(rcCopy.top, baseunitY, 8);
	prc->right	= MulDiv(rcCopy.right, baseunitX, 4);
	prc->bottom	= MulDiv(rcCopy.bottom, baseunitY, 8);
}

// create a child window
HWND CMyWindow::CreateChildWindow(
							LPCTSTR			szWindowClass,
							UINT			nID,
							LPCRECT			prcDialogUnits,
							LPCTSTR			szName,
							DWORD			dwStyle,
							DWORD			dwExtendedStyle)
{
	HDC					hdc		= GetDC(this->m_hwnd);
	RECT				rcClient;			// client rectangle in pixels
	HWND				hwndChild;			// child window

	// convert the dialog units to pixels
	CopyRect(&rcClient, prcDialogUnits);
	this->DialogUnitsToPixels(hdc, &rcClient);
	ReleaseDC(this->m_hwnd, hdc);
	// create the child window
	hwndChild	= CreateWindowEx(
		dwExtendedStyle, 
		szWindowClass,
		szName, 
		dwStyle,
		rcClient.left, 
		rcClient.top, 
		rcClient.right - rcClient.left,
		rcClient.bottom - rcClient.top, 
		this->m_hwnd,
		(HMENU) nID,
		GetOurInstance(),
		NULL);
	ShowWindow(hwndChild, SW_SHOW);
	return hwndChild;
}


BOOL CMyWindow::doRegWindowClass()				// make sure that the window class is registered
{
	WNDCLASSEX			wcx;
	BOOL				fRegistered		= FALSE;
	HINSTANCE			hInst			= GetOurInstance();

	wcx.cbSize		= sizeof(WNDCLASSEX);
	if (!GetClassInfoEx(hInst, MY_WINDOW_CLASS, &wcx))
	{
		wcx.cbSize		= sizeof(wcx);          // size of structure 
		wcx.style		= CS_HREDRAW | CS_VREDRAW;                    // redraw if size changes 
		wcx.lpfnWndProc = WindowProcSciUsbMono;     // points to window procedure 
		wcx.cbClsExtra	= 0;                // no extra class memory 
		wcx.cbWndExtra	= 0;                // no extra window memory 
		wcx.hInstance	= hInst;			// handle to instance 
		wcx.hIcon		= LoadIcon(NULL, IDI_APPLICATION);              // predefined app. icon 
		wcx.hCursor		= LoadCursor(NULL, IDC_ARROW);                    // predefined arrow 
		wcx.hbrBackground = (HBRUSH) GetStockObject(LTGRAY_BRUSH);                  // white background brush 
		wcx.lpszMenuName =  NULL;    // name of menu resource 
		wcx.lpszClassName = MY_WINDOW_CLASS;  // name of window class 
		wcx.hIconSm		= NULL;
		// Register the window class. 
		return RegisterClassEx(&wcx); 
	}
	else return TRUE;
}


// window procedure
LRESULT CALLBACK WindowProcSciUsbMono(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyWindow			*	pWindow		= NULL;

	if (WM_CREATE == uMsg)
	{
		LPCREATESTRUCT	pCreateStruct	= (LPCREATESTRUCT) lParam;
		pWindow			= (CMyWindow*) pCreateStruct->lpCreateParams;
		SetWindowLongPtr(hwnd, GWLP_USERDATA, (LONG_PTR) pWindow);
		pWindow->m_hwnd	= hwnd;
	}
	else
	{
		pWindow			= (CMyWindow*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	}
	if (NULL != pWindow)
		return pWindow->WinProc(uMsg, wParam, lParam);
	else
		return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CMyWindow::WinProc(
							UINT			uMsg,
							WPARAM			wParam,
							LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_CREATE:
		this->OnCreate();
		break;
	case WM_COMMAND:
		if (this->OnCommand(LOWORD(wParam), HIWORD(wParam)))
			return 0;
		break;
	case WM_NOTIFY:
		if (this->OnNotify((LPNMHDR) lParam))
			return 0;
		break;
	case WM_LBUTTONDOWN:
		// attempt initialization
		this->InitMono();
		break;
	case WM_HSCROLL:
		if (this->OnHScroll((HWND) lParam, wParam)) return 0;
		break;
	case WM_DESTROY:
		// disconnect the sinks
		this->DisconnectSinks();
		// default handling
		break;
	default:
		break;
	}
	return DefWindowProc(this->m_hwnd, uMsg, wParam, lParam);
}

void CMyWindow::OnCreate()
{
	RECT					rcChild;
	HWND					hwndEdit;
	INITCOMMONCONTROLSEX	icc;

	icc.dwSize		= sizeof(INITCOMMONCONTROLSEX);
	icc.dwICC		= ICC_BAR_CLASSES | ICC_UPDOWN_CLASS;
	InitCommonControlsEx(&icc);
	// create the child controls
	// title label - mono model and serial number
	rcChild.left	= 7;
	rcChild.top		= 4;
	rcChild.right	= 102;
	rcChild.bottom	= 15;
	this->CreateChildWindow(TEXT("Static"), IDC_WINDOWTITLE, &rcChild, TEXT(""),
		WS_CHILD | SS_LEFT, WS_EX_CLIENTEDGE);
	// position edit bottom
	rcChild.left	= 7;
	rcChild.top		= 18;
	rcChild.right	= 57;
	rcChild.bottom	= 28;
	this->CreateChildWindow(TEXT("Edit"), IDC_EDITPOSITION, &rcChild, TEXT(""),
		WS_CHILD, 0);
	// updown control for the position
	rcChild.left	= 57;
	rcChild.top		= 18;
	rcChild.right	= 67;
	rcChild.bottom	= 28;
	this->CreateChildWindow(TEXT("msctls_updown32"), IDC_UPDPOSITION, &rcChild, TEXT(""),
		WS_CHILD, 0);
	// label for the grating
	rcChild.left	= 72;
	rcChild.top		= 18;
	rcChild.right	= 92;
	rcChild.bottom	= 28;
	this->CreateChildWindow(TEXT("Static"), IDC_GRATING, &rcChild, TEXT(""),
		WS_CHILD | SS_NOTIFY, WS_EX_CLIENTEDGE);
	// updown control for the grating
	rcChild.left	= 92;
	rcChild.top		= 18;
	rcChild.right	= 102;
	rcChild.bottom	= 28;
	this->CreateChildWindow(TEXT("msctls_updown32"), IDC_UPDGRATING, &rcChild, TEXT(""),
		WS_CHILD, 0);
	// trackbar
	rcChild.left	= 7;
	rcChild.top		= 31;
	rcChild.right	= 102;
	rcChild.bottom	= 41;
	this->CreateChildWindow(TEXT("msctls_trackbar32"), IDC_TRKPOSITION, &rcChild, TEXT(""), 
		WS_CHILD | TBS_HORZ | TBS_TOOLTIPS, 0);
	// status
	rcChild.left	= 7;
	rcChild.top		= 44;
	rcChild.right	= 102;
	rcChild.bottom	= 54;
	this->CreateChildWindow(TEXT("Static"), IDC_STATUS, &rcChild, TEXT(""),
		WS_CHILD | SS_LEFT | SS_NOTIFY, WS_EX_CLIENTEDGE);
	// subclass the position edit box
	this->SubclassEditBox(IDC_EDITPOSITION);
	// subclass the status label
	this->SubclassStatusBar(IDC_STATUS);
	// display the current settings
	this->DisplayModelAndSerialNumber();
	// connect the sinks
	this->ConnectSinks();
}

// Horizontal scroll bar
BOOL CMyWindow::OnHScroll(
							HWND			hwndScroll,
							WPARAM			wParam)
{
	DWORD				dwPos;
	IDispatch		*	pdisp;
	if (GetDlgItem(this->m_hwnd, IDC_TRKPOSITION) == hwndScroll)
	{
		if (TB_ENDTRACK == LOWORD(wParam))
		{
			// get the scroll position
            dwPos = SendMessage(hwndScroll, TBM_GETPOS, 0, 0); 
			if (this->GetOurObject(&pdisp))
			{
				Utils_SetDoubleProperty(pdisp, DISPID_position, 1.0 * dwPos);
				pdisp->Release();
			}
			return TRUE;
		}
	}
	return FALSE;
}

// display the current settings
void CMyWindow::DisplayModelAndSerialNumber()
{
	LPTSTR			szModel			= NULL;
	LPTSTR			szSerialNumber	= NULL;
	TCHAR			szStatus[MAX_PATH];

	this->GetModelAndSerialNumber(&szModel, &szSerialNumber);
	if (NULL != szModel && NULL != szSerialNumber)
	{
		StringCchPrintf(szStatus, MAX_PATH, TEXT("%s      %s"), szModel, &szSerialNumber);
		SetDlgItemText(this->m_hwnd, IDC_WINDOWTITLE, szStatus);
	}
	CoTaskMemFree((LPVOID) szModel);
	CoTaskMemFree((LPVOID) szSerialNumber);
}

BOOL CMyWindow::GetOurObject(
							IDispatch	**	ppdisp)
{
	HRESULT				hr;
	hr = this->m_pMyObject->QueryInterface(IID_IDispatch, (LPVOID*) ppdisp);
	return SUCCEEDED(hr);
}

// get the model and the serial number
void CMyWindow::GetModelAndSerialNumber(
							LPTSTR		*	szModel,
							LPTSTR		*	szSerialNumber)
{
	IDispatch		*	pdisp;
	HRESULT				hr;
	VARIANTARG			varg;
	VARIANT				varResult;

	*szModel		= NULL;
	*szSerialNumber	= NULL;
	if (this->GetOurObject(&pdisp))
	{
		InitVariantFromInt32(0, &varg);
		hr = Utils_InvokePropertyGet(pdisp, DISPID_MonoInfo, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToStringAlloc(varResult, szModel);
			VariantClear(&varResult);
		}
		InitVariantFromInt32(1, &varg);
		hr = Utils_InvokePropertyGet(pdisp, DISPID_MonoInfo, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			VariantToStringAlloc(varResult, szSerialNumber);
			VariantClear(&varResult);
		}
		pdisp->Release();
	}
}

// am initialized flag
BOOL CMyWindow::GetAmInitialized()
{
	BOOL			fAmInitialized		= FALSE;
	IDispatch	*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		fAmInitialized	= Utils_GetBoolProperty(pdisp, DISPID_AmOpen);
		pdisp->Release();
	}
	return fAmInitialized;
}

// display the current position
void CMyWindow::DisplayCurrentPosition()
{
	IDispatch		*	pdisp;
	TCHAR				szString[MAX_PATH];
	double				position;

	if (this->GetAmInitialized())
	{
		if (this->GetOurObject(&pdisp))
		{
			position	= Utils_GetDoubleProperty(pdisp, DISPID_position);
			_stprintf_s(szString, MAX_PATH, TEXT("%7.2f"), position);
			SetDlgItemText(this->m_hwnd, IDC_EDITPOSITION, szString);
			// grating number
			SetDlgItemInt(this->m_hwnd, IDC_GRATING, 
				Utils_GetLongProperty(pdisp, DISPID_currentGrating), FALSE);
			SendMessage(GetDlgItem(this->m_hwnd, IDC_TRKPOSITION), TBM_SETPOS,
				TRUE, (LPARAM) (long) floor(position + 0.5));
			pdisp->Release();
		}
	}
	else
	{
		// disable buttons
		EnableWindow(GetDlgItem(this->m_hwnd, IDC_EDITPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwnd, IDC_UPDPOSITION), FALSE);
		EnableWindow(GetDlgItem(this->m_hwnd, IDC_GRATING), FALSE);
		EnableWindow(GetDlgItem(this->m_hwnd, IDC_UPDGRATING), FALSE);
		EnableWindow(GetDlgItem(this->m_hwnd, IDC_TRKPOSITION), FALSE);
	}
}

// get the wavelength range
void CMyWindow::GetWavelengthRange(
							double		*	minWave,
							double		*	maxWave)
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	short int			currentGrating;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	short int			infoIndex;
	*minWave		= 0.0;
	*maxWave		= 0.0;
	if (!this->GetAmInitialized()) return;
	if (this->GetOurObject(&pdisp))
	{
		currentGrating		= (short int) Utils_GetLongProperty(pdisp, DISPID_currentGrating);
		VariantInit(&avarg[1]);
		avarg[1].vt			= VT_BYREF | VT_I2;
		avarg[1].piVal		= &currentGrating;
		// min wavelength
		infoIndex			= 3;
		VariantInit(&avarg[0]);
		avarg[0].vt			= VT_BYREF | VT_I2;
		avarg[0].piVal		= &infoIndex;
		hr = Utils_InvokePropertyGet(pdisp, DISPID_GratingInfo, avarg, 2, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, minWave);
		infoIndex			= 2;
		hr = Utils_InvokePropertyGet(pdisp, DISPID_GratingInfo, avarg, 2, &varResult);
		if (SUCCEEDED(hr)) VariantToDouble(varResult, maxWave);
		pdisp->Release();
	}
}

void CMyWindow::SetTrackbarRange()
{
	double			minWave	= 0.0;
	double			maxWave	= 0.0;
	HWND			hwndTrackbar		= GetDlgItem(this->m_hwnd, IDC_TRKPOSITION);
	long			lMin;
	long			lMax;

	this->GetWavelengthRange(&minWave, &maxWave);
	lMin	= (long) floor(minWave + 0.5);
	lMax	= (long) floor(maxWave + 0.5);
	if (lMax <= lMin)
	{
		lMin	= 0;
		lMax	= 100;
	}
	SendMessage(hwndTrackbar, TBM_SETRANGE, TRUE, MAKELONG(lMin, lMax));
}

BOOL CMyWindow::OnCommand(
							WORD			wmID,
							WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_WINDOWTITLE:
		if (STN_CLICKED == wmEvent)
		{
			// attempt initialization
			this->InitMono();
			return TRUE;
		}
		break;
	case IDC_STATUS:
		if (STN_CLICKED == wmEvent)
		{
			// attempt initialization
			this->InitMono();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

BOOL CMyWindow::OnNotify(
							LPNMHDR			pnmh)
{
	IDispatch			*	pdisp;

	if (IDC_UPDPOSITION == pnmh->idFrom		&&
		UDN_DELTAPOS == pnmh->code)
	{
		LPNMUPDOWN			pnmud		= (LPNMUPDOWN) pnmh;
		double				position	= 0.0;
		// get the current position
		if (this->GetOurObject(&pdisp))
		{
			position	= Utils_GetDoubleProperty(pdisp, DISPID_position);
			Utils_SetDoubleProperty(pdisp, DISPID_position, position - (1.0 * pnmud->iDelta));
			pdisp->Release();
		}
		return TRUE;
	}
	return FALSE;
}

// connect and disconnect the sinks
void CMyWindow::ConnectSinks()
{
	HRESULT				hr;
	IDispatch		*	pdisp;
	CImpISink		*	pSink;
	CImpIPropNotify	*	pPropNotify;
	IUnknown		*	pUnkSink;

	if (this->GetOurObject(&pdisp))
	{
		pSink		= new CImpISink(this);
		hr = pSink->QueryInterface(IID__SciUsbMono, (LPVOID*) &pUnkSink);
		if (SUCCEEDED(hr))
		{
			Utils_ConnectToConnectionPoint(pdisp, pUnkSink, IID__SciUsbMono, TRUE,
				&(this->m_dwCookie));
			pUnkSink->Release();
		}
		pPropNotify	= new CImpIPropNotify(this);
		hr = pPropNotify->QueryInterface(IID_IPropertyNotifySink, (LPVOID*) &pUnkSink);
		if (SUCCEEDED(hr))
		{
			Utils_ConnectToConnectionPoint(pdisp, pUnkSink, IID_IPropertyNotifySink,
				TRUE, &(this->m_dwPropNotifyCookie));
			pUnkSink->Release();
		}
		pdisp->Release();
	}
}

void CMyWindow::DisconnectSinks()
{
	IDispatch		*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		Utils_ConnectToConnectionPoint(pdisp, NULL, IID__SciUsbMono, FALSE, &(this->m_dwCookie));
		Utils_ConnectToConnectionPoint(pdisp, NULL, IID_IPropertyNotifySink,
			FALSE, &(this->m_dwPropNotifyCookie));
		pdisp->Release();
	}
}
// custom sink events
void CMyWindow::OnError(
							LPCTSTR			szError)
{
	this->m_fHaveError	= TRUE;
//	Utils_DoCopyString(&(this->m_szLastError), szError);
	SHStrDup(szError, &this->m_szLastError);
	// repaint the status window
	InvalidateRect(GetDlgItem(this->m_hwnd, IDC_STATUS), NULL, TRUE);
}

void CMyWindow::OnStatusMessage(
							LPCTSTR			szStatus,
							BOOL			fAmBusy)
{
	this->m_fHaveError	= FALSE;
	this->m_fAmBusy		= fAmBusy;
//	Utils_DoCopyString(&(this->m_szCurrentStatus), szStatus);
	SHStrDup(szStatus, &this->m_szCurrentStatus);
	// repaint the status window
	InvalidateRect(GetDlgItem(this->m_hwnd, IDC_STATUS), NULL, TRUE);
}


// property change notifications
BOOL CMyWindow::OnPropRequestEdit(
							DISPID			dispid)
{
	return TRUE;
}

void CMyWindow::OnPropChanged(
							DISPID			dispid)
{
	if (DISPID_ConfigFile == dispid)
	{
		this->DisplayModelAndSerialNumber();			// display the model and serial number
	}
	else if (DISPID_AmOpen == dispid || DISPID_position == dispid)
	{
		this->DisplayCurrentPosition();
	}
	else if (DISPID_currentGrating == dispid)
	{
		// set the trackbar range
		this->SetTrackbarRange();
		// display the current position
		this->DisplayCurrentPosition();
	}
}

CMyWindow::CImpISink::CImpISink(CMyWindow * pMyWindow) :
	m_pMyWindow(pMyWindow),
	m_cRefs(0)
{
}

CMyWindow::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CMyWindow::CImpISink::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || IID__SciUsbMono == riid)
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

STDMETHODIMP_(ULONG) CMyWindow::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyWindow::CImpISink::Release()
{
	ULONG			cRefs;
	cRefs = this->m_cRefs--;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CMyWindow::CImpISink::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CMyWindow::CImpISink::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyWindow::CImpISink::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyWindow::CImpISink::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr) 
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				szString		= NULL;

	VariantInit(&varg);
	if (DISPID_Error == dispIdMember)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
			this->m_pMyWindow->OnError(varg.bstrVal);
//			CoTaskMemFree((LPVOID) szString);
			VariantClear(&varg);
		}
	}
	else if (DISPID_StatusMessage == dispIdMember)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szString);
			SHStrDup(varg.bstrVal, &szString);
			VariantClear(&varg);
			hr = DispGetParam(pDispParams, 1, VT_BOOL, &varg, &uArgErr);
			if (SUCCEEDED(hr))
			{
				this->m_pMyWindow->OnStatusMessage(szString, VARIANT_TRUE == varg.boolVal);
			}
			CoTaskMemFree((LPVOID) szString);
		}
	}
	return S_OK;
}

CMyWindow::CImpIPropNotify::CImpIPropNotify(CMyWindow * pMyWindow) :
	m_pMyWindow(pMyWindow),
	m_cRefs(0)
{
}

CMyWindow::CImpIPropNotify::~CImpIPropNotify()
{
}

// IUnknown methods
STDMETHODIMP CMyWindow::CImpIPropNotify::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyNotifySink == riid)
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

STDMETHODIMP_(ULONG) CMyWindow::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyWindow::CImpIPropNotify::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IPropertyNotify sink methods
STDMETHODIMP CMyWindow::CImpIPropNotify::OnChanged( 
									DISPID			dispID)
{
	this->m_pMyWindow->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CMyWindow::CImpIPropNotify::OnRequestEdit( 
									DISPID			dispID)
{
	return this->m_pMyWindow->OnPropRequestEdit(dispID) ? S_OK : S_FALSE;
}

// edit box subclass
void CMyWindow::SubclassEditBox(
							UINT			uEditBox)
{
	SUBCLASS_DATA	*	pEditBox		= new SUBCLASS_DATA;
	HWND				hwndEdit		= GetDlgItem(this->m_hwnd, uEditBox);
	pEditBox->nControlID	= uEditBox;
	pEditBox->pMyWindow		= this;
	pEditBox->wpOrig		= (WNDPROC) SetWindowLongPtr(hwndEdit, GWLP_WNDPROC,
		(LONG_PTR) SubclassProcEditBox);
	SetWindowLongPtr(hwndEdit, GWLP_USERDATA, (LONG_PTR) pEditBox);
}

void CMyWindow::OnReturnClicked(
							UINT			uEditBox)
{
	TCHAR				szPosition[MAX_PATH];
	float				fVal;
	IDispatch		*	pdisp;
	BOOL				fSuccess		= FALSE;
	if (IDC_EDITPOSITION == uEditBox)
	{
		// read the position from the window
		if (GetDlgItemText(this->m_hwnd, uEditBox, szPosition, MAX_PATH) > 0)
		{
			// set the new position
			if (1 == _stscanf_s(szPosition, TEXT("%f"), &fVal))
			{
				if (this->GetOurObject(&pdisp))
				{
					Utils_SetDoubleProperty(pdisp, DISPID_position, (double) fVal);
					pdisp->Release();
					fSuccess = TRUE;
				}
			}
		}
		if (!fSuccess)
		{
			// display the current position on failure
			this->DisplayCurrentPosition();
		}
	}
}

// edit box subclass
LRESULT CALLBACK SubclassProcEditBox(HWND hwndEdit, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SUBCLASS_DATA	*	pEditBox	= (SUBCLASS_DATA *) GetWindowLongPtr(hwndEdit, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_GETDLGCODE:
		return DLGC_WANTALLKEYS;
	case WM_CHAR:
		if (VK_RETURN == wParam)
		{
			pEditBox->pMyWindow->OnReturnClicked(pEditBox->nControlID);
			return 0;
		}
		break;
	case WM_DESTROY:
		{
			WNDPROC			wpOrig	= pEditBox->wpOrig;
			SetWindowLongPtr(hwndEdit, GWLP_USERDATA, 0);		// clear the user data
			SetWindowLongPtr(hwndEdit, GWLP_WNDPROC, (LONG_PTR) wpOrig);	// reset the window procedure
			delete pEditBox;
			return CallWindowProc(wpOrig, hwndEdit, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return CallWindowProc(pEditBox->wpOrig, hwndEdit, uMsg, wParam, lParam);
}

// status control subclass
void CMyWindow::SubclassStatusBar(
							UINT			uStatus)
{
	SUBCLASS_DATA	*	pStatus			= new SUBCLASS_DATA;
	HWND				hwndStatus		= GetDlgItem(this->m_hwnd, uStatus);
	pStatus->nControlID	= uStatus;
	pStatus->pMyWindow	= this;
	pStatus->wpOrig		= (WNDPROC) SetWindowLongPtr(hwndStatus, GWLP_WNDPROC,
		(LONG_PTR) SubclassProcStatusWindow);
	SetWindowLongPtr(hwndStatus, GWLP_USERDATA, (LONG_PTR) pStatus);
}

void CMyWindow::OnPaintStatusBar(
							UINT			uStatus)
{
	HDC					hdc;
	HWND				hwnd		= GetDlgItem(this->m_hwnd, uStatus);
	PAINTSTRUCT			ps;
	RECT				rc;
	HBRUSH				hbr;			// brush
	HBRUSH				holdBrush;		// old brush
	HPEN				hpen;			// pen
	HPEN				hpenOld;		// old pen
	SIZE				MSpace;			// M space
	int					x, y;			// x and y starting points
	int					ta;				// text alignment
	COLORREF			colorBackground;
	COLORREF			colorForeground;
	TCHAR				szCaption[MAX_PATH];
	size_t				slen;
	int					iBkMode;		// background mode

	hdc		= BeginPaint(hwnd, &ps);
	GetClientRect(hwnd, &rc);
	// determine the colors
	if (this->m_fHaveError)
	{
		colorBackground = RGB(255, 0, 0);
		colorForeground = RGB(0, 0, 0);
		StringCchPrintf(szCaption, MAX_PATH, this->m_szLastError);
	}
	else if (this->m_fAmBusy)
	{
		colorBackground = RGB(0, 192, 0);
		colorForeground = RGB(0, 0, 0);
		StringCchPrintf(szCaption, MAX_PATH, this->m_szCurrentStatus);
	}
	else
	{
		colorBackground = RGB(255, 255, 255);
		colorForeground = RGB(0, 0, 0);
		StringCchPrintf(szCaption, MAX_PATH, this->m_szCurrentStatus);
	}
	// fill in the background
//	hbr = CreateSolidBrush(pMyData->GetBackColor());
	hbr = CreateSolidBrush(colorBackground);
	FillRect(hdc, &rc, hbr);
	DeleteObject((HGDIOBJ) hbr);
	// set the text alignment
	ta = SetTextAlign(hdc, TA_LEFT | TA_BOTTOM);
	// get an M space
	GetTextExtentPoint32(hdc, TEXT("M"), 1, &MSpace);
	// set x and y space
	x = rc.left + MSpace.cx;
	y = (rc.bottom + rc.top + MSpace.cy) >> 1;
	// create the pen
	hpen = CreatePen(PS_SOLID, 1, colorForeground);
	hpenOld = (HPEN) SelectObject(hdc, (HGDIOBJ) hpen);
	// background mode
	iBkMode		= SetBkMode(hdc, TRANSPARENT);
	// output the caption
	StringCchLength(szCaption, MAX_PATH, &slen);
	ExtTextOut(hdc, x, y, ETO_CLIPPED | ETO_OPAQUE, NULL, szCaption, slen, NULL);
	SetBkMode(hdc, iBkMode);
	// reset the text alignment
	SetTextAlign(hdc, ta);
	// cleanup
	SelectObject(hdc, (HGDIOBJ) hpenOld);
	DeleteObject((HGDIOBJ) hpen);
	// end painting
	EndPaint(hwnd, &ps);
}

// status window subclass
LRESULT CALLBACK SubclassProcStatusWindow(HWND hwndStatus, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	SUBCLASS_DATA	*	pStatus	= (SUBCLASS_DATA *) GetWindowLongPtr(hwndStatus, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_PAINT:
		pStatus->pMyWindow->OnPaintStatusBar(pStatus->nControlID);
		return 0;
	case WM_DESTROY:
		{
			WNDPROC			wpOrig	= pStatus->wpOrig;
			SetWindowLongPtr(hwndStatus, GWLP_USERDATA, 0);		// clear the user data
			SetWindowLongPtr(hwndStatus, GWLP_WNDPROC, (LONG_PTR) wpOrig);	// reset the window procedure
			delete pStatus;
			return CallWindowProc(wpOrig, hwndStatus, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return CallWindowProc(pStatus->wpOrig, hwndStatus, uMsg, wParam, lParam);
}

// attempt initialization
void CMyWindow::InitMono()
{
	IDispatch		*	pdisp;
	if (this->GetOurObject(&pdisp))
	{
		if (!Utils_GetBoolProperty(pdisp, DISPID_AmOpen))
		{
			Utils_SetBoolProperty(pdisp, DISPID_AmOpen, TRUE);
		}
		pdisp->Release();
	}
}
