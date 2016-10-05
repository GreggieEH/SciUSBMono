#pragma once

class CMySciUsbMono;

class CSciArcus
{
public:
	CSciArcus(CMySciUsbMono * pMySciUsbMono, LPCTSTR szSerialNumber);
	~CSciArcus(void);
	BOOL					doInit();						// initialize the SciArcus object
	void					SetDllDirectory(
								LPCTSTR			szDllDirectory);
	void					SetToolsDirectory(
								LPCTSTR			szToolsDirectory);
	BOOL					GetDllLoaded();
	void					SetDllLoaded();
	BOOL					GetDeviceConnected();
	void					SetDeviceConnected(
								BOOL			fDeviceConnected);
	BOOL					SendReceive(
								LPCTSTR			Send,
								LPTSTR		*	Receive);
	long					GetValue(
								long			Index);
	void					SetValue(
								long			Index,
								long			NewValue);
	BOOL					StoreSettings();
	// find the device index
	BOOL					FindDeviceIndex();
	// motor homing
	BOOL					MotorHomed();
//	void					SetMotorHomed();
	void					Home9030();
	void					Home9055();
	void					Home9040();
	void					Home9010();
	BOOL					GetAmHoming();
	// go to position
	void					MoveToPosition(
								long			NewPosition);
	BOOL					GetMotorMoving();
	BOOL					GetCurrentPosition(
								long		*	position);
	// motor enabled
	BOOL					GetMotorEnabled();
	void					SetMotorEnabled(
								BOOL			motorEnabled);
	// motor IDLE current
	long					GetMotorIdleCurrent();
	void					SetMotorIdleCurrent(
								long			idleCurrent);
	// motor Run Current
	long					GetMotorRunCurrent();
	void					SetMotorRunCurrent(
								long			runCurrent);
	// high speed
	long					GetHighSpeed();
	void					SetHighSpeed(
								long			highSpeed);
	// motor status
	long					GetMotorStatus();
	// check for move error
	void					CheckMoveError(
								LPCTSTR			szReply,
								BOOL		*	fMoveError);
	void					ClearMoveError();
	// set minimum position
	void					SetMinimumPosition(
								long			minimumPosition);
	void					SetMaximumPosition(
								long			maximumPosition);
protected:
	// sink events
	void					OnError(
								LPCTSTR			Error);
	void					OnHaveConnectedDevice(
								LPCTSTR			SerialNumber);
	void					OnHaveConnectedDevice();
//	// set the subclass
//	void					SetSubclass();
//	// remove the subclass
//	void					RemoveSubclass();
	// check if a move was completed
	BOOL					CheckMoveCompleted(
								BOOL		*	fNegativeLimit,
								BOOL		*	fPositiveLimit);
	// move difference
	long					GetMoveDifference(
								long			NewPosition);
	// check absolute position
	BOOL					CheckAbsolutePosition(
								long			position,
								LPTSTR			szError,
								int				nBufferSize);
	// send a command, eat return
	BOOL					MySendCommand(
								LPCTSTR			szCommand);
	// wait until move completed
	BOOL					WaitUntilMoveCompleted(
								BOOL		*	fPositiveLimit = NULL);
private:
	CMySciUsbMono		*	m_pMySciUsbMono;
	IDispatch			*	m_pdispSciArcus;
	LPTSTR					m_szSerialNumber;			// our serial number
	long					m_deviceIndex;				// our device index
	BOOL					m_fAmHomed;					// motor homed flag
	BOOL					m_fAmHoming;				// motor homing flag
	BOOL					m_fAmMoving;				// motor position changing
//	// idle and run currents
//	long					m_idleCurrent;
//	long					m_runCurrent;
	// sink handling
	IID						m_iidSink;					// sink interface id
	DWORD					m_dwCookie;					// connection cookie
	// subclassing the main window
	HWND					m_hwndSubclass;				// subclass window
	// absolute minimum position
	BOOL					m_fMinimumSet;
	long					m_Minimum;
	// absolute maximum position
	BOOL					m_fMaximumSet;
	long					m_Maximum;
	// dialog window
	HWND					m_hwndDlg;

// dialog procedure
friend LRESULT CALLBACK		DlgProcSciArcus(HWND, UINT, WPARAM, LPARAM);

	// sink implementation
	class CImpISink : public IDispatch
	{
	public:
		CImpISink(CSciArcus * pSciArcus);
		~CImpISink();
		// IUnknown methods
		STDMETHODIMP		QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv);
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();
		// IDispatch methods
		STDMETHODIMP		GetTypeInfoCount( 
								PUINT			pctinfo);
		STDMETHODIMP		GetTypeInfo( 
								UINT			iTInfo,         
								LCID			lcid,                   
								ITypeInfo	**	ppTInfo);
		STDMETHODIMP		GetIDsOfNames( 
								REFIID			riid,                  
								OLECHAR		**  rgszNames,  
								UINT			cNames,          
								LCID			lcid,                   
								DISPID		*	rgDispId);
		STDMETHODIMP		Invoke( 
								DISPID			dispIdMember,      
								REFIID			riid,              
								LCID			lcid,                
								WORD			wFlags,              
								DISPPARAMS	*	pDispParams,  
								VARIANT		*	pVarResult,  
								EXCEPINFO	*	pExcepInfo,  
								PUINT			puArgErr); 
	private:
		CSciArcus		*	m_pSciArcus;
		ULONG				m_cRefs;
		DISPID				m_dispidError;
		DISPID				m_dispidHaveConnectedDevice;
	};
	friend CImpISink;
};

struct MY_SUBCLASS
{
	CSciArcus		*	m_pSciArcus;
	WNDPROC				m_wpOrig;
};