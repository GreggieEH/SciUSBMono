#include "stdafx.h"
#include "DlgDisplayConfigValues.h"
#include "dispids.h"
#include "DlgArcusSetup.h"

CDlgDisplayConfigValues::CDlgDisplayConfigValues(void) :
	m_pdisp(NULL),				// our object
	m_hwndDlg(NULL),			// dialog handle
	m_DisplayGrating(1)			// grating for which config file info is displayed
{
}

CDlgDisplayConfigValues::~CDlgDisplayConfigValues(void)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
}

void CDlgDisplayConfigValues::SetOurObject(
						IDispatch	*	pdisp)
{
	Utils_RELEASE_INTERFACE(this->m_pdisp);
	if (NULL != pdisp)
	{
		this->m_pdisp	= pdisp;
		this->m_pdisp->AddRef();
	}
}

void CDlgDisplayConfigValues::DoOpenDialog(
						HWND			hwndParent)
{
	DialogBoxParam(
		GetOurInstance(),
		MAKEINTRESOURCE(IDD_DIALOGDisplayConfigValues),
		hwndParent,
		(DLGPROC) DlgProcDisplayConfigValues,
		(LPARAM) this);
}

LRESULT CALLBACK DlgProcDisplayConfigValues(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgDisplayConfigValues	*	pDlg	= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgDisplayConfigValues*) lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg	= hwndDlg;
	}
	else
	{
		pDlg = (CDlgDisplayConfigValues*) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}


BOOL CDlgDisplayConfigValues::DlgProc(
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

BOOL CDlgDisplayConfigValues::OnInitDialog()
{
	INITCOMMONCONTROLSEX			icc;
	icc.dwSize		= sizeof(INITCOMMONCONTROLSEX);
	icc.dwICC		= ICC_UPDOWN_CLASS | ICC_BAR_CLASSES;	// need tool tips and up down classes
	InitCommonControlsEx(&icc);
	Utils_CenterWindow(this->m_hwndDlg);
	this->DisplayModel();
	this->DisplaySerialNumber();
	this->DisplayInputAngle();
	this->DisplayOutputAngle();
	this->DisplayAutoGrating();
	this->DisplayDriveType();
	this->DisplayGearTeeth();
	this->DisplayStepsPerRev();
	this->DisplayApplyBacklash();
	this->DisplayNumberOfGratings();
	// display the first grating
	this->m_DisplayGrating	= 1;
	this->DisplayGratingInfo();
	return TRUE;
}

BOOL CDlgDisplayConfigValues::OnCommand(
						WORD			wmID,
						WORD			wmEvent)
{
	switch (wmID)
	{
	case IDOK:
	case IDCANCEL:
		EndDialog(this->m_hwndDlg, wmID);
		return TRUE;
	case IDC_ARCUSMOTORPARAMS:
		{
			CDlgArcusSetup			dlg(this->m_pdisp);
			dlg.doOpenDialog(GetParent(this->m_hwndDlg), GetOurInstance());
		}
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgDisplayConfigValues::OnNotify(
						LPNMHDR			pnmh)
{
	if (UDN_DELTAPOS == pnmh->code)
	{
		LPNMUPDOWN	pnmud		= (LPNMUPDOWN) pnmh;
		if (IDC_UPDGRATINGNUMBER == pnmh->idFrom)
		{
			long		displayGrating		= this->m_DisplayGrating - pnmud->iDelta;
			long		numGratings			= Utils_GetLongProperty(this->m_pdisp, DISPID_NumberOfGratings);
			if (displayGrating < 1) displayGrating = 1;
			if (displayGrating > numGratings) displayGrating = numGratings;
			this->m_DisplayGrating = displayGrating;
			this->DisplayGratingInfo();
			return TRUE;
		}
	}
	return FALSE;
}


// display values
void CDlgDisplayConfigValues::DisplayModel()
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

void CDlgDisplayConfigValues::DisplaySerialNumber()
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

void CDlgDisplayConfigValues::DisplayInputAngle()
{
	VARIANT				Value;
	double				dval;
	TCHAR				szString[MAX_PATH];
	SetDlgItemText(this->m_hwndDlg, IDC_INPUTANGLE, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_INPUTANGLE, &Value))
	{
		VariantToDouble(Value, &dval);
		_stprintf_s(szString, MAX_PATH, TEXT("%6.3f"), float(dval));
		SetDlgItemText(this->m_hwndDlg, IDC_INPUTANGLE, szString);
	}
}

void CDlgDisplayConfigValues::DisplayOutputAngle()
{
	VARIANT				Value;
	double				dval;
	TCHAR				szString[MAX_PATH];
	SetDlgItemText(this->m_hwndDlg, IDC_OUTPUTANGLE, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_OUTPUTANGLE, &Value))
	{
		VariantToDouble(Value, &dval);
		_stprintf_s(szString, MAX_PATH, TEXT("%6.3f"), float(dval));
		SetDlgItemText(this->m_hwndDlg, IDC_OUTPUTANGLE, szString);
	}
}

void CDlgDisplayConfigValues::DisplayAutoGrating()
{
	HRESULT				hr;
	VARIANTARG			varResult;
	BOOL				fCanSelect	= FALSE;
	hr = Utils_InvokeMethod(this->m_pdisp, DISPID_CanAutoSelect, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fCanSelect);
	SetDlgItemText(this->m_hwndDlg, IDC_AUTOGRATING, 
		fCanSelect ? TEXT("Avail") : TEXT("NOT Avail"));
}

void CDlgDisplayConfigValues::DisplayDriveType()
{
	VARIANT				Value;
	LPTSTR				szString		= NULL;
	SetDlgItemText(this->m_hwndDlg, IDC_DRIVETYPE, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_DRIVETYPE, &Value))
	{
		VariantToStringAlloc(Value, &szString);
		SetDlgItemText(this->m_hwndDlg, IDC_DRIVETYPE, szString);
		CoTaskMemFree((LPVOID) szString);
		VariantClear(&Value);
	}
}

void CDlgDisplayConfigValues::DisplayGearTeeth()
{
	VARIANT				Value;
	long				gearTeeth		= 0;
	if (this->GetMonoInfo(MONO_INFO_GEARTEETH, &Value))
	{
		VariantToInt32(Value, &gearTeeth);
	}
	SetDlgItemInt(this->m_hwndDlg, IDC_GEARTEETH, gearTeeth, FALSE);
}

void CDlgDisplayConfigValues::DisplayStepsPerRev()
{
	VARIANT				Value;
	double				dval;
	TCHAR				szString[MAX_PATH];
	SetDlgItemText(this->m_hwndDlg, IDC_STEPSPERREV, TEXT(""));
	if (this->GetMonoInfo(MONO_INFO_STEPSPERREV, &Value))
	{
		VariantToDouble(Value, &dval);
		_stprintf_s(szString, MAX_PATH, TEXT("%6.2f"), dval);
		SetDlgItemText(this->m_hwndDlg, IDC_STEPSPERREV, szString);
	}
}

void CDlgDisplayConfigValues::DisplayApplyBacklash()
{
	SetDlgItemText(this->m_hwndDlg, IDC_APPLYBACKLASH,
		Utils_GetBoolProperty(this->m_pdisp, DISPID_ApplyBacklashCorrection) ? 
			TEXT("Yes") : TEXT("No"));
}

void CDlgDisplayConfigValues::DisplayNumberOfGratings()
{
	long			nGratings = Utils_GetLongProperty(this->m_pdisp, DISPID_NumberOfGratings);
	SetDlgItemInt(this->m_hwndDlg, IDC_NUMBEROFGRATINGS, nGratings, FALSE);
	EnableWindow(GetDlgItem(this->m_hwndDlg, IDC_UPDGRATINGNUMBER), nGratings > 0);
}

BOOL CDlgDisplayConfigValues::GetMonoInfo(
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

BOOL CDlgDisplayConfigValues::GetGratingInfo(
						long			gratingID,
						long			Index,
						VARIANT		*	GratingInfo)
{
	HRESULT			hr;
	short int		gnum	= (short) gratingID;
	short int		sval	= (short) Index;
	VARIANTARG		avarg[2];
	VariantInit(&avarg[1]);
	avarg[1].vt		= VT_BYREF | VT_I2;
	avarg[1].piVal	= &gnum;
	VariantInit(&avarg[0]);
	avarg[0].vt		= VT_BYREF | VT_I2;
	avarg[0].piVal	= &sval;
	hr = Utils_InvokePropertyGet(this->m_pdisp, DISPID_GratingInfo, avarg, 2, GratingInfo);
	return SUCCEEDED(hr);
}

double CDlgDisplayConfigValues::GetDoubleGratingInfo(
						long			gratingID,
						long			Index)
{
	VARIANT			Value;
	double			dval	= 0.0;
	if (this->GetGratingInfo(gratingID, Index, &Value))
	{
		VariantToDouble(Value, &dval);
	}
	return dval;
}


void CDlgDisplayConfigValues::DisplayGratingInfo()
{
	TCHAR			szString[MAX_PATH];
	SetDlgItemInt(this->m_hwndDlg, IDC_GRATINGNUMBER, this->m_DisplayGrating, FALSE);
	this->DisplayGratingPitch();
	this->DisplayGratingBlaze();
	this->DisplayGratingMinWave();
	this->DisplayGratingMaxWave();
	this->DisplayGratingZeroPos();
	this->DisplayGratingPhaseError();
	this->DisplayGratingStepsPerNM();
	// display the quadratic fit
	_stprintf_s(szString, MAX_PATH, TEXT("%8.3g"), 
		this->GetDoubleGratingInfo(this->m_DisplayGrating, GRATING_INFO_OFFSETTERM));
	SetDlgItemText(this->m_hwndDlg, IDC_OFFSET, szString);
	_stprintf_s(szString, MAX_PATH, TEXT("%8.3g"),
		this->GetDoubleGratingInfo(this->m_DisplayGrating, GRATING_INFO_LINEARTERM));
	SetDlgItemText(this->m_hwndDlg, IDC_LINEAR, szString);
	_stprintf_s(szString, MAX_PATH, TEXT("%8.3g"),
		this->GetDoubleGratingInfo(this->m_DisplayGrating, GRATING_INFO_QUADTERM));
	SetDlgItemText(this->m_hwndDlg, IDC_QUADTERM, szString);
}


void CDlgDisplayConfigValues::DisplayGratingPitch()
{
	VARIANT			Value;
	long			pitch		= 0;
	if (this->GetGratingInfo(this->m_DisplayGrating, 0, &Value))
		VariantToInt32(Value, &pitch);
	SetDlgItemInt(this->m_hwndDlg, IDC_PITCH, pitch, FALSE);
}

void CDlgDisplayConfigValues::DisplayGratingBlaze()
{
	VARIANT			Value;
	LPTSTR			szBlaze		= NULL;
	SetDlgItemText(this->m_hwndDlg, IDC_BLAZE, TEXT(""));
	if (this->GetGratingInfo(this->m_DisplayGrating, 1, &Value))
	{
		VariantToStringAlloc(Value, &szBlaze);
		SetDlgItemText(this->m_hwndDlg, IDC_BLAZE, szBlaze);
		CoTaskMemFree((LPVOID) szBlaze);
		VariantClear(&Value);
	}
}

void CDlgDisplayConfigValues::DisplayGratingMinWave()
{
	VARIANT			Value;
	double			minWave		= 0.0;
	if (this->GetGratingInfo(this->m_DisplayGrating, 3, &Value))
	{
		VariantToDouble(Value, &minWave);
	}
	SetDlgItemInt(this->m_hwndDlg, IDC_MINWAVE, (long) floor(minWave+ 0.5), TRUE);
}

void CDlgDisplayConfigValues::DisplayGratingMaxWave()
{
	VARIANT			Value;
	double			maxWave		= 0.0;
	if (this->GetGratingInfo(this->m_DisplayGrating, 2, &Value))
	{
		VariantToDouble(Value, &maxWave);
	}
	SetDlgItemInt(this->m_hwndDlg, IDC_MAXWAVE, (long) floor(maxWave+ 0.5), TRUE);
}

void CDlgDisplayConfigValues::DisplayGratingZeroPos()
{
	VARIANTARG		varg;
	VARIANT			Value;
	HRESULT			hr;
	long			zeroOffset		= 0;
	InitVariantFromInt32(this->m_DisplayGrating, &varg);
	hr = Utils_InvokePropertyGet(this->m_pdisp, DISPID_GratingZeroOffset, &varg, 1, &Value);
	if (SUCCEEDED(hr)) VariantToInt32(Value, &zeroOffset);
	SetDlgItemInt(this->m_hwndDlg, IDC_ZEROPOS, zeroOffset, TRUE);
}

void CDlgDisplayConfigValues::DisplayGratingPhaseError()
{
	VARIANT			Value;
	double			phaseError		= 0.0;
	TCHAR			szString[MAX_PATH];
	if (this->GetGratingInfo(this->m_DisplayGrating, 6, &Value))
	{
		VariantToDouble(Value, &phaseError);
	}
	_stprintf_s(szString, MAX_PATH, TEXT("%6.2f"), phaseError);
	SetDlgItemText(this->m_hwndDlg, IDC_PHASEERROR, szString);
}

void CDlgDisplayConfigValues::DisplayGratingStepsPerNM()
{
	VARIANT			Value;
	double			stepsPerNM		= 0.0;
	TCHAR			szString[MAX_PATH];
	if (this->GetGratingInfo(this->m_DisplayGrating, 9, &Value))
	{
		VariantToDouble(Value, &stepsPerNM);
	}
	_stprintf_s(szString, MAX_PATH, TEXT("%8.2f"), stepsPerNM);
	SetDlgItemText(this->m_hwndDlg, IDC_STEPSPERNM, szString);
}
