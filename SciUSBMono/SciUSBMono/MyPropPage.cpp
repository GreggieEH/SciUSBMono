#include "stdafx.h"
#include "MyPropPage.h"
#include "dispids.h"
#include "MyGuids.h"
#include "DlgArcusSetup.h"
#include "NamedObjects.h"

CMyPropPage::CMyPropPage(void) :
	m_cRefs(0),					// object reference count
	m_hwndDlg(NULL),			// property page window handle
	m_pPropertyPageSite(NULL),	// property page site object
	m_pdisp(NULL),				// object which we are communicating
	m_fDirty(FALSE),			// dirty flag
	m_fAmInitializing(FALSE),	// flag handling OnInitDialog
	// property notify sink
	m_dwPropNotifyCookie(0),
	m_szName(NULL)				// name for this object
{
	CNamedObjects	*	pNamedObjects	= GetNamedObjects();
	if (NULL != pNamedObjects)
		pNamedObjects->GetNonSelectedName(&m_szName);
}

CMyPropPage::~CMyPropPage(void)
{
	Utils_RELEASE_INTERFACE(this->m_pPropertyPageSite);
	if (NULL != this->m_pdisp)
	{
		Utils_SetBoolProperty(this->m_pdisp, DISPID_PageSelected, FALSE);
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, IID_IPropertyNotifySink,
			FALSE, &(this->m_dwPropNotifyCookie));
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
}

// IUnknown methods
STDMETHODIMP CMyPropPage::QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IPropertyPage == riid)
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

STDMETHODIMP_(ULONG) CMyPropPage::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyPropPage::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IPropertyPage methods
STDMETHODIMP CMyPropPage::SetPageSite( 
								IPropertyPageSite *pPageSite)
{
	if (NULL == this->m_pPropertyPageSite)
	{
		if (NULL == pPageSite) return E_INVALIDARG;
		this->m_pPropertyPageSite	= pPageSite;
		this->m_pPropertyPageSite->AddRef();
	}
	else
	{
		// non null value
		if (NULL != pPageSite) return E_UNEXPECTED;
		Utils_RELEASE_INTERFACE(this->m_pPropertyPageSite);
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::Activate( 
								HWND			hWndParent ,  //Parent window handle
								LPCRECT			prc ,      //Pointer to RECT structure
								BOOL			bModal)
{
	HWND			hwndDlg;
	hwndDlg = CreateDialogParam(GetOurInstance(),
		MAKEINTRESOURCE(IDD_PROPPAGESciUsbMono), 
		hWndParent,
		(DLGPROC) MyPropertyPageProc,
		(LPARAM) this);
	Utils_SetLongProperty(this->m_pdisp, DISPID_SetupWindow, (long) hwndDlg);
	return NULL != hwndDlg ? S_OK : E_FAIL;
}

STDMETHODIMP CMyPropPage::Deactivate(void)
{
	if (NULL != this->m_hwndDlg)
	{
		DestroyWindow(this->m_hwndDlg);
		this->m_hwndDlg = NULL;
		Utils_SetLongProperty(this->m_pdisp, DISPID_SetupWindow, 0);
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::GetPageInfo( 
								PROPPAGEINFO *	pPageInfo)
{
	RECT				rc;							// client rectangle
	HDC					hdc;

	pPageInfo->cb		= sizeof(PROPPAGEINFO);
	// set the title
	pPageInfo->pszTitle	= NULL;
	if (NULL != this->m_szName)
	{
		SHStrDup(this->m_szName, &(pPageInfo->pszTitle));
	}
	else
	{
		// default title
		SHStrDup(TEXT("Sciencetech Mono"), &(pPageInfo->pszTitle));
	}
	pPageInfo->pszHelpFile	= NULL;
	pPageInfo->pszDocString = NULL;
	pPageInfo->dwHelpContext	= 0;
	if (NULL != this->m_hwndDlg)
	{
		GetWindowRect(this->m_hwndDlg, &rc);
	}
	else
	{
		hdc			= GetDC(NULL);
		rc.left		= 0;
		rc.top		= 0;
		rc.right	= 250;
		rc.bottom	= 110;
		this->DialogUnitsToPixels(hdc, &rc);
		ReleaseDC(NULL, hdc);
	}
	pPageInfo->size.cx = rc.right - rc.left;
	pPageInfo->size.cy = rc.bottom - rc.top;
	return S_OK;
}

// map dialog units to pixels
void CMyPropPage::DialogUnitsToPixels(
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
	baseunitY	= (long) floor((txtMetric.tmHeight * 1.2) + 0.5);
	// copy the input rectangle
	CopyRect(&rcCopy, prc);
	// determine the output dimensions
	prc->left	= MulDiv(rcCopy.left, baseunitX, 4);
	prc->top	= MulDiv(rcCopy.top, baseunitY, 8);
	prc->right	= MulDiv(rcCopy.right, baseunitX, 4);
	prc->bottom	= MulDiv(rcCopy.bottom, baseunitY, 8);
}


STDMETHODIMP CMyPropPage::SetObjects( 
								ULONG			cObjects ,  //Number of IUnknown pointers in the ppUnk array
								IUnknown	**	ppUnk)
{
	HRESULT				hr;
	ULONG				i;
//	IDispatch		*	pdisp;
	BOOL				fDone;
	LPTSTR				szName;
	// property notification sink
	CImpIPropNotify	*	pSink;
	IUnknown		*	pUnkSink;

	// disconnect the sinks
//	this->DisconnectCustomSink();
//	this->DisconnectPropertyNotification();
	if (NULL != this->m_pdisp)
	{
		Utils_ConnectToConnectionPoint(this->m_pdisp, NULL, IID_IPropertyNotifySink,
			FALSE, &(this->m_dwPropNotifyCookie));
		Utils_SetBoolProperty(this->m_pdisp, DISPID_PageSelected, FALSE);
		Utils_RELEASE_INTERFACE(this->m_pdisp);
	}
	if (cObjects > 0)
	{
		// find the first non selected page
		fDone	= FALSE;
		i		= 0;
		while (i < cObjects && !fDone)
		{
			if (NULL != this->m_szName)
			{
				IDispatch	*	pdisp;
				hr = (ppUnk[i])->QueryInterface(IID_ISciUsbMono, (LPVOID*) &pdisp);
				if (SUCCEEDED(hr))
				{
					// check the object name
					szName	= NULL;
					if (this->GetObjectName(pdisp, &szName))
					{
						if (0 == lstrcmpi(szName, this->m_szName))
						{
							// have found the desired page
							this->m_pdisp	= pdisp;
							this->m_pdisp->AddRef();
							fDone = TRUE;
						}
						CoTaskMemFree((LPVOID) szName);
					}
					pdisp->Release();
				}
			}
			else
			{
				// have found the desired object
				hr = (ppUnk[i])->QueryInterface(IID_ISciUsbMono, (LPVOID*) &(this->m_pdisp));
				fDone = SUCCEEDED(hr);
			}
			if (fDone)
			{
				// set page selected flag
				Utils_SetBoolProperty(this->m_pdisp, DISPID_PageSelected, TRUE);
				// connect the property notification
				pSink		= new CImpIPropNotify(this);
				hr = pSink->QueryInterface(IID_IPropertyNotifySink, (LPVOID*) &pUnkSink);
				if (SUCCEEDED(hr))
				{
					Utils_ConnectToConnectionPoint(this->m_pdisp, pUnkSink, IID_IPropertyNotifySink,
						TRUE, &(this->m_dwPropNotifyCookie));
					pUnkSink->Release();
				}
				else
					delete pSink;
			}
			if (!fDone) i++;
		}
		if (!fDone) return E_NOINTERFACE;
	}
	return S_OK;
}

STDMETHODIMP CMyPropPage::Show(
								UINT			nCmdShow)
{
	if (NULL != this->m_hwndDlg)
	{
		ShowWindow(this->m_hwndDlg, nCmdShow);
		return S_OK;
	}
	else return E_UNEXPECTED;
}

STDMETHODIMP CMyPropPage::Move( 
								LPCRECT			prc)
{
	if (NULL != this->m_hwndDlg)
	{
		MoveWindow(this->m_hwndDlg,
			prc->left, prc->top,
			prc->right - prc->left,
			prc->bottom - prc->top,
			TRUE);
		return S_OK;
	}
	else return E_UNEXPECTED;
}

STDMETHODIMP CMyPropPage::IsPageDirty(void)
{
	return this->m_fDirty ? S_OK : S_FALSE;
}

STDMETHODIMP CMyPropPage::Apply(void)
{
//	this->ApplyIdleCurrent();
//	this->ApplyRunCurrent();
//	this->ApplyHighSpeed();
	this->ApplyStepsPerRev();
	this->m_fDirty	= FALSE;
	return S_OK;
}

STDMETHODIMP CMyPropPage::Help( 
								LPCOLESTR		pszHelpDir)
{
	return E_NOTIMPL;
}

STDMETHODIMP CMyPropPage::TranslateAccelerator( 
								LPMSG			pMsg)
{
	return E_NOTIMPL;
}

// set the dirty flag
void CMyPropPage::SetDirty(
								BOOL			fDirty)
{
	// check the am initializing flag
	if (this->m_fAmInitializing) return;
	this->m_fDirty = fDirty;
	if (this->m_fDirty && NULL != this->m_pPropertyPageSite)
	{
		this->m_pPropertyPageSite->OnStatusChange(PROPPAGESTATUS_DIRTY);
	}
}

// property page procedure
LRESULT CALLBACK MyPropertyPageProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage			*	pMyPropPage		= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pMyPropPage		= (CMyPropPage *) lParam;
		pMyPropPage->m_hwndDlg	= hwndDlg;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, (LONG_PTR) lParam);
	}
	else
	{
		pMyPropPage		= (CMyPropPage *) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pMyPropPage)
		return pMyPropPage->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

// dialog procedure
BOOL CMyPropPage::DlgProc(
								UINT			uMsg,
								WPARAM			wParam,
								LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_COMMAND:
		return this->OnCommand(LOWORD(wParam), HIWORD(wParam));
	case WM_NOTIFY:
		return this->OnNotify((LPNMHDR) lParam);
	default:
		break;
	}
	return FALSE;
}

BOOL CMyPropPage::OnInitDialog()		// handler for WM_INITDIALOG
{
	// set am initializing flag
	this->m_fAmInitializing	= TRUE;
	// display the current values
	this->DisplayModel();
	this->DisplaySerialNumber();
//	this->DisplayIdleCurrent();
//	this->DisplayRunCurrent();
//	this->DisplayHighSpeed();
	this->DisplayStepsPerRev();
	this->DisplayCurrentPosition();
	this->DisplayNumberOfGratings();
	this->DisplayCurrentGrating();
	this->DisplayAmInitialized();
	// enable or disable the autograting check box
	EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_CHKAUTOGRATING),
		this->CanAutoSelect());
	this->DisplayAutoGrating();
	EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDCURRENTGRATING),
		this->GetNumberOfGratings() > 1);
	// check if change zero offset is allowed
	BOOL		fAllowZeroOffsetChange	= Utils_GetBoolProperty(this->m_pdisp, DISPID_AllowChangeZeroOffset);
	EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_EDITZEROADJUST), fAllowZeroOffsetChange);
	EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_ADJUSTZEROPOSITION), fAllowZeroOffsetChange);
	// display the grating info
	DisplayGratingInfo();
	// subclass the edit boxes
	this->SubclassEditBox(IDC_IDLECURRENT);
	this->SubclassEditBox(IDC_RUNCURRENT);
	this->SubclassEditBox(IDC_EDITHIGHSPEED);
	this->SubclassEditBox(IDC_STEPSPERREV);
	//this->SubclassEditBox(IDC_EDITPOSITION);
	this->SubclassEditBox(IDC_EDITZEROADJUST);
	this->SubclassEditBox(IDC_EDITCURRENTPOSITION);
	this->DisplayReInitOnScanStart();
	// clear am initializing flag
	this->m_fAmInitializing	= FALSE;
	return TRUE;
}

BOOL CMyPropPage::OnCommand(			// command handler
								WORD			wmID,
								WORD			wmEvent)
{
	switch (wmID)
	{
	case IDC_IDLECURRENT:
		if (EN_CHANGE == wmEvent)
		{
			this->SetDirty(TRUE);
			return TRUE;
		}
		break;
	case IDC_RUNCURRENT:
		if (EN_CHANGE == wmEvent)
		{
			this->SetDirty(TRUE);
			return TRUE;
		}
		break;
	case IDC_EDITHIGHSPEED:
		if (EN_CHANGE == wmEvent)
		{
			this->SetDirty(TRUE);
			return TRUE;
		}
		break;
	case IDC_STEPSPERREV:
		if (EN_CHANGE == wmEvent)
		{
			this->SetDirty(TRUE);
			return TRUE;
		}
		break;
	case IDC_SETPOSITION:
		this->OnSetPosition();
		return TRUE;
	case IDC_ADJUSTZEROPOSITION:
		this->OnAdjustZeroPosition();
		return TRUE;
	case IDC_CHKAMINITIALIZED:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedAmInitialized();
			return TRUE;
		}
		break;
	case IDC_CHKAUTOGRATING:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedAutoGrating();
			return TRUE;
		}
		break;
	case IDC_GRATING1:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetLongProperty(this->m_pdisp, DISPID_currentGrating, 1);
			this->DisplayCurrentGrating();
			this->DisplayCurrentPosition();
			return TRUE;
		}
		break;
	case IDC_GRATING2:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetLongProperty(this->m_pdisp, DISPID_currentGrating, 2);
			this->DisplayCurrentGrating();
			this->DisplayCurrentPosition();
			return TRUE;
		}
		break;
	case IDC_GRATING3:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetLongProperty(this->m_pdisp, DISPID_currentGrating, 3);
			this->DisplayCurrentGrating();
			this->DisplayCurrentPosition();
			return TRUE;
		}
		break;
	case IDC_GRATING4:
		if (BN_CLICKED == wmEvent)
		{
			Utils_SetLongProperty(this->m_pdisp, DISPID_currentGrating, 4);
			this->DisplayCurrentGrating();
			this->DisplayCurrentPosition();
			return TRUE;
		}
		break;
	case IDC_ARCUSSETUP:
		{
			CDlgArcusSetup			dlg(this->m_pdisp);
			dlg.doOpenDialog(GetParent(this->m_hwndDlg), GetOurInstance());
		}
		return TRUE;
	case IDC_CHKREINIT:
		if (BN_CLICKED == wmEvent)
		{
			this->OnClickedInitOnScanStart();
			return TRUE;
		}
		break;

	default:
		break;
	}
	return FALSE;
}

BOOL CMyPropPage::OnNotify(
								LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code)
	{
		LPNMUPDOWN			pnmud	= (LPNMUPDOWN) pnmh;
		if (IDC_UPDCURRENTGRATING == pnmh->idFrom)
		{
			long		grating		= this->GetCurrentGrating();
			grating -= pnmud->iDelta;
			if (grating >= 1 && grating <= this->GetNumberOfGratings())
			{
				Utils_SetLongProperty(this->m_pdisp, DISPID_currentGrating, grating);
			}
			this->DisplayCurrentGrating();
			this->DisplayCurrentPosition();
		}
	}
	return FALSE;
}

void CMyPropPage::OnEditBoxReturnClicked(
								UINT			nEditBox)
{
	switch (nEditBox)
	{
	case IDC_IDLECURRENT:
		SetFocus(GetDlgItem(this->m_hwndDlg, IDC_RUNCURRENT));
		break;
	case IDC_RUNCURRENT:
		SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITHIGHSPEED));
		break;
	case IDC_EDITHIGHSPEED:
		SetFocus(GetDlgItem(this->m_hwndDlg, IDC_STEPSPERREV));
		break;
	case IDC_STEPSPERREV:
		break;
	case IDC_EDITCURRENTPOSITION:
		this->OnSetPosition();
		break;
	case IDC_EDITZEROADJUST:
		this->OnAdjustZeroPosition();
		break;
	default:
		break;
	}
}

// subclass an edit box
void CMyPropPage::SubclassEditBox(
								UINT			uControlID)
{
	MY_SUBCLASS_STRUCT	*	pSubclass	= new MY_SUBCLASS_STRUCT;
	HWND					hwndEdit	= GetDlgItem(this->m_hwndDlg, uControlID);

	pSubclass->pMyPropPage	= this;
	pSubclass->wpOrig		= (WNDPROC) SetWindowLongPtr(hwndEdit, GWLP_WNDPROC, (LONG) MyEditBoxSubclass);
	SetWindowLongPtr(hwndEdit, GWLP_USERDATA, (LONG) pSubclass);
}

// edit box subclass
LRESULT CALLBACK MyEditBoxSubclass(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CMyPropPage::MY_SUBCLASS_STRUCT*	pSubclass = (CMyPropPage::MY_SUBCLASS_STRUCT*) GetWindowLongPtr(hwnd, GWLP_USERDATA);
	switch (uMsg)
	{
	case WM_GETDLGCODE:
		return DLGC_WANTALLKEYS;			// want the return key
	case WM_CHAR:
		if (0 != (0x8000 & GetKeyState(VK_RETURN)))
		{
			pSubclass->pMyPropPage->OnEditBoxReturnClicked(GetDlgCtrlID(hwnd));
			return 0;
		}
		break;
	case WM_DESTROY:
		// remove the subclass
		{
			WNDPROC			wpOrig	= pSubclass->wpOrig;
			delete pSubclass;
			SetWindowLongPtr(hwnd, GWLP_USERDATA, 0);
			SetWindowLongPtr(hwnd, GWLP_WNDPROC, (LONG) wpOrig);
			return CallWindowProc(wpOrig, hwnd, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return CallWindowProc(pSubclass->wpOrig, hwnd, uMsg, wParam, lParam);
}

// display values
void CMyPropPage::DisplayModel()
{
	LPTSTR				szModel		= NULL;
	VARIANT				Value;
	size_t				slen;
	SetDlgItemText(this->m_hwndDlg, IDC_MODEL, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_MODEL, &Value))
	{
		VariantToStringAlloc(Value, &szModel);
		if (NULL != szModel)
		{
			StringCchLength(szModel, MAX_PATH, &slen);
			if (slen > 0)
			{
				SetDlgItemText(this->m_hwndDlg, IDC_MODEL, szModel);
			}
			CoTaskMemFree((LPVOID) szModel);
		}
		VariantClear(&Value);
	}
}

void CMyPropPage::DisplaySerialNumber()
{
	LPTSTR				szSerialNumber		= NULL;
	VARIANT				Value;
	size_t				slen;
	SetDlgItemText(this->m_hwndDlg, IDC_SERIALNUMBER, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_SERIALNUMBER, &Value))
	{
		VariantToStringAlloc(Value, &szSerialNumber);
		if (NULL != szSerialNumber)
		{
			StringCchLength(szSerialNumber, MAX_PATH, &slen);
			if (slen > 0)
			{
				SetDlgItemText(this->m_hwndDlg, IDC_SERIALNUMBER, szSerialNumber);
			}
			CoTaskMemFree((LPVOID) szSerialNumber);
		}
		VariantClear(&Value);
	}
}

/*
void CMyPropPage::DisplayIdleCurrent()
{
	VARIANT			Value;
	long			lval;
	SetDlgItemText(this->m_hwndDlg, IDC_IDLECURRENT, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_IDLECURRENT, &Value))
	{
		VariantToInt32(Value, &lval);
		SetDlgItemInt(this->m_hwndDlg, IDC_IDLECURRENT, lval, FALSE);
	}
}

void CMyPropPage::DisplayRunCurrent()
{
	VARIANT			Value;
	long			lval;
	SetDlgItemText(this->m_hwndDlg, IDC_RUNCURRENT, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_RUNCURRENT, &Value))
	{
		VariantToInt32(Value, &lval);
		SetDlgItemInt(this->m_hwndDlg, IDC_RUNCURRENT, lval, FALSE);
	}
}

void CMyPropPage::DisplayHighSpeed()
{
	VARIANT			Value;
	long			lval;
	SetDlgItemText(this->m_hwndDlg, IDC_EDITHIGHSPEED, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_HIGHSPEED, &Value))
	{
		VariantToInt32(Value, &lval);
		SetDlgItemInt(this->m_hwndDlg, IDC_EDITHIGHSPEED, lval, FALSE);
	}
}
*/

void CMyPropPage::DisplayStepsPerRev()
{
	VARIANT			Value;
	double			dval;
	TCHAR			szString[MAX_PATH];
	SetDlgItemText(this->m_hwndDlg, IDC_STEPSPERREV, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_STEPSPERREV, &Value))
	{
		VariantToDouble(Value, &dval);
		_stprintf_s(szString, MAX_PATH, TEXT("%5.2f"), dval);
		SetDlgItemText(this->m_hwndDlg, IDC_STEPSPERREV, szString);
	}
}

/*
void CMyPropPage::ApplyIdleCurrent()
{
	VARIANT			Value;
	long			lval;
	BOOL			fSuccess;

	lval = GetDlgItemInt(this->m_hwndDlg, IDC_IDLECURRENT, &fSuccess, FALSE);
	if (fSuccess)
	{
		InitVariantFromInt32(lval, &Value);
		this->SetMonoInfo(MONO_INFO_IDLECURRENT, &Value);
	}
	this->DisplayIdleCurrent();
}

void CMyPropPage::ApplyRunCurrent()
{
	VARIANT			Value;
	long			lval;
	BOOL			fSuccess;

	lval = GetDlgItemInt(this->m_hwndDlg, IDC_RUNCURRENT, &fSuccess, FALSE);
	if (fSuccess)
	{
		InitVariantFromInt32(lval, &Value);
		this->SetMonoInfo(MONO_INFO_RUNCURRENT, &Value);
	}
	this->DisplayRunCurrent();
}

void CMyPropPage::ApplyHighSpeed()
{
	VARIANT			Value;
	long			lval;
	BOOL			fSuccess;

	lval = GetDlgItemInt(this->m_hwndDlg, IDC_EDITHIGHSPEED, &fSuccess, FALSE);
	if (fSuccess)
	{
		InitVariantFromInt32(lval, &Value);
		this->SetMonoInfo(MONO_INFO_HIGHSPEED, &Value);
	}
	this->DisplayRunCurrent();
}
*/

void CMyPropPage::ApplyStepsPerRev()
{
	TCHAR			szString[MAX_PATH];
	float			fval;
	VARIANT			Value;
	if (GetDlgItemText(this->m_hwndDlg, IDC_STEPSPERREV, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
		{
			InitVariantFromDouble(fval, &Value);
			this->SetMonoInfo(MONO_INFO_STEPSPERREV, &Value);
		}
	}
	this->DisplayStepsPerRev();
}

BOOL CMyPropPage::GetMonoInfo(
								long			Index,
								VARIANT		*	MonoInfo)
{
	short int		sval	= (short) Index;
	VARIANTARG		varg;
	HRESULT			hr;
	VariantInit(&varg);
	varg.vt		= VT_BYREF | VT_I2;
	varg.piVal	= &sval;
	hr = Utils_InvokePropertyGet(this->m_pdisp, DISPID_MonoInfo, &varg, 1, MonoInfo);
	return SUCCEEDED(hr);
}

void CMyPropPage::SetMonoInfo(
								long			Index,
								VARIANT		*	MonoInfo)
{
	short int		sval	= (short) Index;
	VARIANTARG		avarg[2];
	VariantInit(&avarg[1]);
	avarg[1].vt		= VT_BYREF | VT_I2;
	avarg[1].piVal	= &sval;
	VariantInit(&avarg[0]);
	VariantCopy(&avarg[0], MonoInfo);
	Utils_InvokePropertyPut(this->m_pdisp, DISPID_MonoInfo, avarg, 2);
	VariantClear(&avarg[0]);
}

// get grating info
void CMyPropPage::GetGratingInfo(
								long			grating,
								long			Index,
								VARIANT		*	GratingInfo)
{
	short int		sval		= (short) Index;
	short int		sGrating	= (short) grating;
	VARIANTARG		avarg[2];
	HRESULT			hr;
	VariantInit(&avarg[1]);
	avarg[1].vt		= VT_BYREF | VT_I2;
	avarg[1].piVal	= &sGrating;
	VariantInit(&avarg[0]);
	avarg[0].vt		= VT_BYREF | VT_I2;
	avarg[0].piVal	= &sval;
	hr = Utils_InvokePropertyGet(this->m_pdisp, DISPID_GratingInfo, avarg, 2, GratingInfo);
}

void CMyPropPage::DisplayCurrentPosition()
{
	TCHAR				szString[MAX_PATH];

	_stprintf_s(szString, MAX_PATH, TEXT("%6.2f"),
		Utils_GetDoubleProperty(this->m_pdisp, DISPID_position));
	SetDlgItemText(this->m_hwndDlg, IDC_EDITCURRENTPOSITION, szString);
}

// push buttons
void CMyPropPage::OnSetPosition()
{
	TCHAR				szString[MAX_PATH];
	float				fval;
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITPOSITION, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
		{
			Utils_SetDoubleProperty(this->m_pdisp, DISPID_position, (double) fval);
		}
	}
//	this->DisplayCurrentPosition();
}

void CMyPropPage::OnAdjustZeroPosition()
{
	HRESULT				hr;
	TCHAR				szString[MAX_PATH];
	float				fval;
	double				newZero;
	long				newZeroOffset	= 0;
	short int			grating		= (short) this->GetCurrentGrating();
	VARIANT_BOOL		fToNM;
	VARIANTARG			avarg[4];
	VARIANT				varResult;
	BOOL				fSuccess	= FALSE;
	VARIANTARG			pvarg[2];
	if (GetDlgItemText(this->m_hwndDlg, IDC_EDITZEROADJUST, szString, MAX_PATH) > 0)
	{
		if (1 == _stscanf_s(szString, TEXT("%f"), &fval))
		{
			newZero	= fval;
			fToNM	= VARIANT_FALSE;
			VariantInit(&avarg[3]);
			avarg[3].vt			= VT_BYREF | VT_BOOL;
			avarg[3].pboolVal	= &fToNM;
			VariantInit(&avarg[2]);
			avarg[2].vt			= VT_BYREF | VT_I2;
			avarg[2].piVal		= &grating;
			VariantInit(&avarg[1]);
			avarg[1].vt			= VT_BYREF | VT_R8;
			avarg[1].pdblVal	= &newZero;
			VariantInit(&avarg[0]);
			avarg[0].vt			= VT_BYREF | VT_I4;
			avarg[0].plVal		= &newZeroOffset;
			hr = Utils_InvokeMethod(this->m_pdisp, DISPID_ConvertStepsToNM, avarg, 4, &varResult);
			if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
		}
	}
	if (fSuccess)
	{
		InitVariantFromInt32(grating, &pvarg[1]);
		InitVariantFromInt32(newZeroOffset, &pvarg[0]);
		Utils_InvokePropertyPut(this->m_pdisp, DISPID_GratingZeroOffset, pvarg, 2); 	
	}
}

long CMyPropPage::GetNumberOfGratings()
{
	return Utils_GetLongProperty(this->m_pdisp, DISPID_NumberOfGratings);
}

void CMyPropPage::DisplayNumberOfGratings()
{
	long			numGratings = this->GetNumberOfGratings();
	HWND			hwnd;
	SetDlgItemInt(this->m_hwndDlg, IDC_NUMBEROFGRATINGS, numGratings, FALSE);
	if (numGratings < 4)
	{
		hwnd = GetDlgItem(this->m_hwndDlg, IDC_GRATING4);
		EnableWindow(hwnd, FALSE);
		ShowWindow(hwnd, SW_HIDE);
	}
}

void CMyPropPage::DisplayCurrentGrating()
{
//	SetDlgItemInt(this->m_hwndDlg, IDC_CURRENTGRATING,
//		this->GetCurrentGrating(), FALSE);
	long			currentGrating		= this->GetCurrentGrating();
	SendMessage(GetDlgItem(this->m_hwndDlg, IDC_GRATING1), BM_SETCHECK,
		1 == currentGrating ? BST_CHECKED : BST_UNCHECKED, 0);
	SendMessage(GetDlgItem(this->m_hwndDlg, IDC_GRATING2), BM_SETCHECK,
		2 == currentGrating ? BST_CHECKED : BST_UNCHECKED, 0);
	SendMessage(GetDlgItem(this->m_hwndDlg, IDC_GRATING3), BM_SETCHECK,
		3 == currentGrating ? BST_CHECKED : BST_UNCHECKED, 0);
	SendMessage(GetDlgItem(this->m_hwndDlg, IDC_GRATING4), BM_SETCHECK,
		4 == currentGrating ? BST_CHECKED : BST_UNCHECKED, 0);
}

long CMyPropPage::GetCurrentGrating()
{
	return Utils_GetLongProperty(this->m_pdisp, DISPID_currentGrating);
}

// display am initialized
void CMyPropPage::DisplayAmInitialized()
{
	SendMessage(
		GetDlgItem(this->m_hwndDlg, IDC_CHKAMINITIALIZED),
		BM_SETCHECK,
		Utils_GetBoolProperty(this->m_pdisp, DISPID_AmOpen) ? BST_CHECKED : BST_UNCHECKED, 0);
}

// toggle on/off am open
void CMyPropPage::OnClickedAmInitialized()
{
	BOOL			fAmOpen		= Utils_GetBoolProperty(this->m_pdisp, DISPID_AmOpen);
	Utils_SetBoolProperty(this->m_pdisp, DISPID_AmOpen, !fAmOpen);
}

void CMyPropPage::DisplayAutoGrating()
{
	SendMessage(
		GetDlgItem(this->m_hwndDlg, IDC_CHKAUTOGRATING),
		BM_SETCHECK,
		Utils_GetBoolProperty(this->m_pdisp, DISPID_autoGrating) ? BST_CHECKED : BST_UNCHECKED,
		0);
}

void CMyPropPage::OnClickedAutoGrating()
{
	BOOL			fAutoGrating		= Utils_GetBoolProperty(this->m_pdisp, DISPID_autoGrating);
	Utils_SetBoolProperty(this->m_pdisp, DISPID_autoGrating, !fAutoGrating);
}

// property change notification
void CMyPropPage::OnPropChanged(
								DISPID			dispid)
{
	switch (dispid)
	{
	case DISPID_AmOpen:
		this->DisplayAmInitialized();
		break;
	case DISPID_position:
		this->DisplayCurrentPosition();
		break;
	case DISPID_currentGrating:
		this->DisplayCurrentGrating();
		break;
	case DISPID_autoGrating:
		this->DisplayAutoGrating();
		break;
	case DISPID_ReInitOnScanStart:
		this->DisplayReInitOnScanStart();
		break;
	default:
		break;
	}
}

// is auto grating available?
BOOL CMyPropPage::CanAutoSelect()
{
	HRESULT			hr;
	VARIANT			varResult;
	BOOL			fAutoSelect		= FALSE;
	hr = Utils_InvokeMethod(this->m_pdisp, DISPID_CanAutoSelect, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fAutoSelect);
	return fAutoSelect;
}

CMyPropPage::CImpIPropNotify::CImpIPropNotify(CMyPropPage * pMyPropPage) :
	m_pMyPropPage(pMyPropPage),
	m_cRefs(0)
{
}

CMyPropPage::CImpIPropNotify::~CImpIPropNotify()
{
}

// IUnknown methods
STDMETHODIMP CMyPropPage::CImpIPropNotify::QueryInterface(
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

STDMETHODIMP_(ULONG) CMyPropPage::CImpIPropNotify::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMyPropPage::CImpIPropNotify::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IPropertyNotifySink methods
STDMETHODIMP CMyPropPage::CImpIPropNotify::OnChanged( 
									DISPID			dispID)
{
	this->m_pMyPropPage->OnPropChanged(dispID);
	return S_OK;
}

STDMETHODIMP CMyPropPage::CImpIPropNotify::OnRequestEdit( 
									DISPID			dispID)
{
	return S_OK;				// always allow the request
}

// display grating info
void CMyPropPage::DisplayGratingInfo()
{
	long			numGratings		= this->GetNumberOfGratings();
	long			i;
	UINT			nID;
	TCHAR			szGratingInfo[MAX_PATH];
	VARIANT			GratingInfo;
	long			pitch;
	LPTSTR			blaze		= NULL;
	double			minWave;
	double			maxWave;

	for (i=0; i<numGratings; i++)
	{
		switch (i)
		{
		case 0:
			nID = IDC_GRATING1;
			break;
		case 1:
			nID = IDC_GRATING2;
			break;
		case 2:
			nID = IDC_GRATING3;
			break;
		case 3:
			nID = IDC_GRATING4;
			break;
		default:
			break;
		}
		EnableWindow(GetDlgItem(this->m_hwndDlg, nID), TRUE);
		ShowWindow(GetDlgItem(this->m_hwndDlg, nID), SW_SHOW);
		// display the grating info
		this->GetGratingInfo(i+1, 0, &GratingInfo);		// pitch
		VariantToInt32(GratingInfo, &pitch);
		this->GetGratingInfo(i+1, 2, &GratingInfo);		// max wavelength
		VariantToDouble(GratingInfo, &maxWave);
		this->GetGratingInfo(i+1, 3, &GratingInfo);
		VariantToDouble(GratingInfo, &minWave);
		this->GetGratingInfo(i+1, 1, &GratingInfo);
		VariantToStringAlloc(GratingInfo, &blaze);
		VariantClear(&GratingInfo);
		if (NULL != blaze)
		{
			StrTrim(blaze, L" ");
			StringCchPrintf(szGratingInfo, MAX_PATH,
				TEXT("Blaze %s Pitch %1d Wave %1d - %1d"),
				blaze, pitch, (long) floor(minWave), (long) floor(maxWave));
			CoTaskMemFree((LPVOID) blaze);
		}
		else
		{
			StringCchPrintf(szGratingInfo, MAX_PATH,
				TEXT("Pitch %1d Wave range %1d - %1d"),
				blaze, (long) floor(minWave), (long) floor(maxWave));
		}
		SetDlgItemText(this->m_hwndDlg, nID, szGratingInfo);
	}
	while (i < 4)
	{
		switch (i)
		{
		case 0:
			nID = IDC_GRATING1;
			break;
		case 1:
			nID = IDC_GRATING2;
			break;
		case 2:
			nID = IDC_GRATING3;
			break;
		case 3:
			nID = IDC_GRATING4;
			break;
		default:
			break;
		}
		EnableWindow(GetDlgItem(this->m_hwndDlg, nID), FALSE);
		ShowWindow(GetDlgItem(this->m_hwndDlg, nID), SW_HIDE);
		i++;
	}
}

// get object name
BOOL CMyPropPage::GetObjectName(
								IDispatch	*	pdisp,
								LPTSTR		*	szName)
{
	*szName		= NULL;
	Utils_GetStringProperty(pdisp, DISPID_DisplayName, szName);
	return NULL != *szName;
/*
	HRESULT				hr;
	IOleObject		*	pOleObject;
	IOleClientSite	*	pClientSite;		// client site
	IOleControlSite	*	pControlSite;
	IDispatch		*	pdispExtended;		// extended control
	BOOL				fSuccess		= FALSE;
	*szName		= NULL;
	hr = pdisp->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
	if (SUCCEEDED(hr))
	{
		hr = pOleObject->GetClientSite(&pClientSite);
		pOleObject->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pClientSite->QueryInterface(IID_IOleControlSite, (LPVOID*) &pControlSite);
		pClientSite->Release();
	}
	if (SUCCEEDED(hr))
	{
		hr = pControlSite->GetExtendedControl(&pdispExtended);
		pControlSite->Release();
	}
	if (SUCCEEDED(hr))
	{
		Utils_GetStringProperty(pdispExtended, DISPID_Name, szName);
		fSuccess = NULL != *szName;
		pdispExtended->Release();
	}
	return fSuccess;
*/
}


void CMyPropPage::DisplayReInitOnScanStart()
{
	BOOL		fReInitOnScanStart = Utils_GetBoolProperty(this->m_pdisp, DISPID_ReInitOnScanStart);
	Button_SetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKREINIT), fReInitOnScanStart ? BST_CHECKED : BST_UNCHECKED);
}

void CMyPropPage::OnClickedInitOnScanStart()
{
	BOOL		fReInitOnScanStart = BST_CHECKED == Button_GetCheck(GetDlgItem(this->m_hwndDlg, IDC_CHKREINIT));
	Utils_SetBoolProperty(this->m_pdisp, DISPID_ReInitOnScanStart, !fReInitOnScanStart);
}
