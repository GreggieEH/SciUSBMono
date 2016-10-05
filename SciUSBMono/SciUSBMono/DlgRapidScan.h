#pragma once

class CMySciUsbMono;

class CDlgRapidScan
{
public:
	CDlgRapidScan(CMySciUsbMono * pMySciUsbMono);
	~CDlgRapidScan(void);
	BOOL			StartRapidScan(
						long			grating,
						double			startWavelength,
						double			endWavelength,
						double			NMPerSecond);
	void			OnChangedGrating(
						long			grating);
	void			OnChangedWavelength(
						double			wavelength);
	BOOL			GetRapidScanRunning();
	void			StopRapidScan();
protected:
	BOOL			DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	void			StartRapidScan();
	void			OnRapidScanStepped();
private:
	CMySciUsbMono * m_pMySciUsbMono;
	long			m_grating;
	double			m_startWavelength;
	double			m_endWavelength;
	double			m_NMPerSecond;
	HWND			m_hwndDlg;
	BOOL			m_fWaitingForStartGrating;
	BOOL			m_fWaitingForStartWavelength;

friend LRESULT CALLBACK	DlgProcRapidScan(HWND, UINT, WPARAM, LPARAM);
};
