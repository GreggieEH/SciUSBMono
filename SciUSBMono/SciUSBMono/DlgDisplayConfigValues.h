#pragma once

class CDlgDisplayConfigValues
{
public:
	CDlgDisplayConfigValues(void);
	~CDlgDisplayConfigValues(void);
	void			SetOurObject(
						IDispatch	*	pdisp);
	void			DoOpenDialog(
						HWND			hwndParent);
protected:
	BOOL			DlgProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	BOOL			OnNotify(
						LPNMHDR			pnmh);
	void			DisplayModel();
	void			DisplaySerialNumber();
	void			DisplayInputAngle();
	void			DisplayOutputAngle();
	void			DisplayAutoGrating();
	void			DisplayDriveType();
	void			DisplayGearTeeth();
	void			DisplayStepsPerRev();
	void			DisplayApplyBacklash();
	void			DisplayNumberOfGratings();
	BOOL			GetMonoInfo(
						long			Index,
						VARIANT		*	MonoInfo);
	BOOL			GetGratingInfo(
						long			gratingID,
						long			Index,
						VARIANT		*	GratingInfo);
	void			DisplayGratingInfo();
	void			DisplayGratingPitch();
	void			DisplayGratingBlaze();
	void			DisplayGratingMinWave();
	void			DisplayGratingMaxWave();
	void			DisplayGratingZeroPos();
	void			DisplayGratingPhaseError();
	void			DisplayGratingStepsPerNM();
	double			GetDoubleGratingInfo(
						long			gratingID,
						long			Index);
private:
	IDispatch	*	m_pdisp;				// our object
	HWND			m_hwndDlg;				// dialog handle
	long			m_DisplayGrating;		// grating for which config file info is displayed

friend LRESULT CALLBACK	DlgProcDisplayConfigValues(HWND, UINT, WPARAM, LPARAM);
};
