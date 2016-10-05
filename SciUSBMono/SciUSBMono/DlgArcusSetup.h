#pragma once

class CDlgArcusSetup
{
public:
	CDlgArcusSetup(IDispatch * pdisp);
	~CDlgArcusSetup(void);
	BOOL			doOpenDialog(
						HWND			hwndParent,
						HINSTANCE		hInst);
protected:
	BOOL			DialogProc(
						UINT			uMsg,
						WPARAM			wParam,
						LPARAM			lParam);
	BOOL			OnInitDialog();
	BOOL			OnCommand(
						WORD			wmID,
						WORD			wmEvent);
	BOOL			OnGetDefID();
	void			DisplayIdleCurrent();
	void			DisplayRunCurrent();
	void			DisplayHighSpeed();
	void			ApplyIdleCurrent();
	void			ApplyRunCurrent();
	void			ApplyHighSpeed();
	BOOL			GetMonoInfo(
						long			Index,
						VARIANT		*	MonoInfo);
	void			SetMonoInfo(
						long			Index,
						VARIANT		*	MonoInfo);
private:
	IDispatch	*	m_pdisp;
	HWND			m_hwndDlg;
	BOOL			m_fAllowClose;

friend LRESULT CALLBACK	DlgProcArcusSetup(HWND, UINT, WPARAM, LPARAM);
};
