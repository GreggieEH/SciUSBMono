#include "stdafx.h"
#include "DlgRapidScan.h"
#include "MySciUsbMono.h"

#define				TMR_RAPIDSCANSTEPPED		0x0001

CDlgRapidScan::CDlgRapidScan(CMySciUsbMono * pMySciUsbMono) :
	m_pMySciUsbMono(pMySciUsbMono),
	m_grating(0),
	m_startWavelength(0.0),
	m_endWavelength(0.0),
	m_NMPerSecond(0.0),
	m_hwndDlg(NULL),
	m_fWaitingForStartGrating(FALSE),
	m_fWaitingForStartWavelength(FALSE)
{
}

CDlgRapidScan::~CDlgRapidScan(void)
{
	if (NULL != this->m_hwndDlg)
	{
		DestroyWindow(this->m_hwndDlg);
		this->m_hwndDlg	= NULL;
	}
}

BOOL CDlgRapidScan::StartRapidScan(
						long			grating,
						double			startWavelength,
						double			endWavelength,
						double			NMPerSecond)
{
	this->m_grating			= grating;
	this->m_startWavelength	= startWavelength;
	this->m_endWavelength	= endWavelength;
	this->m_NMPerSecond		= NMPerSecond;

	// clear flags
	this->m_fWaitingForStartGrating			= FALSE;
	this->m_fWaitingForStartWavelength		= FALSE;
	// create the dialog
	CreateDialogParam(GetOurInstance(), MAKEINTRESOURCE(IDD_DIALOGDUMMY),
				NULL, (DLGPROC) DlgProcRapidScan, 0);
	return this->GetRapidScanRunning();
}

void CDlgRapidScan::OnChangedGrating(
						long			grating)
{
	if (this->GetRapidScanRunning() && this->m_fWaitingForStartGrating)
	{
		double			currentWavelength;
		double			waveDiff;				// wavelength difference

		// check if the grating was properly set
		if (grating != this->m_grating)
		{
			this->m_fWaitingForStartGrating	= TRUE;
			this->m_pMySciUsbMono->SetcurrentGrating(this->m_grating);
		}
		else
		{
			// clear flag
			this->m_fWaitingForStartGrating	= FALSE;
			currentWavelength	= this->m_pMySciUsbMono->Getposition();
			waveDiff			= fabs(currentWavelength - this->m_startWavelength);
			if (waveDiff > 1.0)
			{	
				this->m_pMySciUsbMono->Setposition(this->m_startWavelength);
				this->m_fWaitingForStartGrating	= TRUE;
			}
			else
			{
				// clear flag
				this->m_fWaitingForStartWavelength	= FALSE;
				this->m_pMySciUsbMono->_StartRapidScan(this->m_grating,
					this->m_startWavelength, this->m_endWavelength, this->m_NMPerSecond);
			}
		}
	}
}

void CDlgRapidScan::OnChangedWavelength(
						double			wavelength)
{
	if (this->GetRapidScanRunning() && this->m_fWaitingForStartWavelength)
	{
		double waveDiff			= fabs(wavelength - this->m_startWavelength);
		if (waveDiff > 1.0)
		{	
			this->m_pMySciUsbMono->Setposition(this->m_startWavelength);
			this->m_fWaitingForStartGrating	= TRUE;
		}
		else
		{
			// clear flag
			this->m_fWaitingForStartWavelength	= FALSE;
			this->m_pMySciUsbMono->_StartRapidScan(this->m_grating,
				this->m_startWavelength, this->m_endWavelength, this->m_NMPerSecond);
		}
	}
}

void CDlgRapidScan::StartRapidScan()
{
	if (this->m_pMySciUsbMono->_StartRapidScan(this->m_grating,
		this->m_startWavelength, this->m_endWavelength, this->m_NMPerSecond))
	{
		SetTimer(this->m_hwndDlg, TMR_RAPIDSCANSTEPPED, 100, NULL);
	}
}

BOOL CDlgRapidScan::GetRapidScanRunning()
{
	return NULL != this->m_hwndDlg && IsWindow(this->m_hwndDlg);
}

void CDlgRapidScan::StopRapidScan()
{
	if (NULL != this->m_hwndDlg)
	{
		DestroyWindow(this->m_hwndDlg);
	}
}

BOOL CDlgRapidScan::DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return this->OnInitDialog();
	case WM_TIMER:
		if (TMR_RAPIDSCANSTEPPED == wParam)
		{
			this->OnRapidScanStepped();
			return TRUE;
		}
		break;
	case WM_DESTROY:
		this->m_hwndDlg	= NULL;
		this->m_pMySciUsbMono->OnRapidScanClosed();
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

BOOL CDlgRapidScan::OnInitDialog()
{
	double			currentWavelength;
	double			waveDiff;				// wavelength difference

	// check if need to set the grating
	if (this->m_pMySciUsbMono->GetcurrentGrating() != this->m_grating)
	{
		this->m_fWaitingForStartGrating	= TRUE;
		this->m_pMySciUsbMono->SetcurrentGrating(this->m_grating);
	}
	else
	{
		currentWavelength	= this->m_pMySciUsbMono->Getposition();
		waveDiff			= fabs(currentWavelength - this->m_startWavelength);
		if (waveDiff > 1.0)
		{	
			this->m_pMySciUsbMono->Setposition(this->m_startWavelength);
			this->m_fWaitingForStartGrating	= TRUE;
		}
		else
		{
			this->m_pMySciUsbMono->_StartRapidScan(this->m_grating,
				this->m_startWavelength, this->m_endWavelength, this->m_NMPerSecond);
		}
	}
	return TRUE;
}

LRESULT CALLBACK DlgProcRapidScan(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	CDlgRapidScan*	pDlg	= NULL;
	if (WM_INITDIALOG == uMsg)
	{
		pDlg = (CDlgRapidScan*) lParam;
		SetWindowLongPtr(hwndDlg, GWLP_USERDATA, lParam);
		pDlg->m_hwndDlg	= hwndDlg;
	}
	else
	{
		pDlg = (CDlgRapidScan*) GetWindowLongPtr(hwndDlg, GWLP_USERDATA);
	}
	if (NULL != pDlg)
		return pDlg->DlgProc(uMsg, wParam, lParam);
	else
		return FALSE;
}

void CDlgRapidScan::OnRapidScanStepped()
{
	this->m_pMySciUsbMono->OnRapidScanStepped();
}
