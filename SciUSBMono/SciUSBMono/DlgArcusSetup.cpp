#include "stdafx.h"
#include "DlgArcusSetup.h"
#include "dispids.h"

CDlgArcusSetup::CDlgArcusSetup(IDispatch * pdisp) :
	m_pdisp(NULL),
	m_hwndDlg(NULL),
	m_fAllowClose(FALSE)
{
	if (NULL != pdisp)
	{
		this->m_pdisp = pdisp;
		this->m_pdisp->AddRef();
	}
}

CDlgArcusSetup::~CDlgArcusSetup(void)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

BOOL CDlgArcusSetup::doOpenDialog(
						HWND			hwndParent,
						HINSTANCE		hInst)
{
	return IDOK == DialogBoxParam(
		hInst,
		MAKEINTRESOURCE(IDD_DIALOGArcusSetup),
		hwndParent,
		(DLGPROC) DlgProcArcusSetup,
		(LPARAM) this);
}

LRESULT CALLBACK DlgProcArcusSetup(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgArcusSetup	*	pDlg	= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgArcusSetup*) lParam;
		pDlg->m_hwndDlg = hwndDlg;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
	}
	else
	{
		pDlg = (CDlgArcusSetup*) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DialogProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

BOOL CDlgArcusSetup::DialogProc(
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
	case DM_GETDEFID:
		return this->OnGetDefID();
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgArcusSetup::OnInitDialog()
{
	this->DisplayRunCurrent();
	this->DisplayHighSpeed();
	this->DisplayIdleCurrent();
	return TRUE;
}

BOOL CDlgArcusSetup::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
		if (!this->m_fAllowClose)
		{
			this->m_fAllowClose	= TRUE;
			return TRUE;
		}
		// fall through
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	case IDC_IDLECURRENT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->ApplyIdleCurrent();
			return TRUE;
		}
		break;
	case IDC_RUNCURRENT:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->ApplyRunCurrent();
			return TRUE;
		}
		break;
	case IDC_EDITHIGHSPEED:
		if (EN_KILLFOCUS == wmEvent)
		{
			this->ApplyHighSpeed();
			return TRUE;
		}
		break;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgArcusSetup::OnGetDefID()
{
	SHORT			shKeyState		= GetKeyState(VK_RETURN);
	HWND			hwndFocus;
	UINT			nID;
	if (0 != (0x8000 & shKeyState))
	{
		hwndFocus		= GetFocus();
		nID				= GetDlgCtrlID(hwndFocus);
		switch (nID)
		{
		case IDC_IDLECURRENT:
			// dont allow closing
			this->m_fAllowClose	= FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_RUNCURRENT));
			return TRUE;
		case IDC_RUNCURRENT:
			this->m_fAllowClose	= FALSE;
			SetFocus(GetDlgItem(this->m_hwndDlg, IDC_EDITHIGHSPEED));
			return TRUE;
		case IDC_EDITHIGHSPEED:
			// allow closing with this one
			this->ApplyHighSpeed();
			return TRUE;
		default:
			break;
		}
	}
	return FALSE;
}


void CDlgArcusSetup::DisplayIdleCurrent()
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

void CDlgArcusSetup::DisplayRunCurrent()
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

void CDlgArcusSetup::DisplayHighSpeed()
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


void CDlgArcusSetup::ApplyIdleCurrent()
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

void CDlgArcusSetup::ApplyRunCurrent()
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

void CDlgArcusSetup::ApplyHighSpeed()
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

BOOL CDlgArcusSetup::GetMonoInfo(
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

void CDlgArcusSetup::SetMonoInfo(
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

