#pragma once

class CMyObject;
class CSciArcus;
class CIniFile;
class CQuadratic;
class CDlgRapidScan;
class CStartupValues;

class CMySciUsbMono
{
public:
	CMySciUsbMono(CMyObject * pMyObject);
	~CMySciUsbMono(void);
	HRESULT				InitNew();
	long				GetSaveSize();
	HRESULT				Save(
							IStream		*	pStm, 
							BOOL			fClearDirty);
	HRESULT				Load(	
							IStream		*	pStm);
	BOOL				GetDirty();
	HRESULT				Load(
							IPropertyBag*	pPropBag, 
							IErrorLog	*	pErrorLog);
	HRESULT				Save(
							IPropertyBag*	pPropBag,
							BOOL			fClearDirty,
							BOOL			fSaveAllProperties);
	BOOL				PaintToRect(
							HDC				hdc, 
							LPCRECT			prc);
	// properties and methods
	double				Getposition();
	void				Setposition(
							double			position);
	long				GetcurrentGrating();
	BOOL				SetcurrentGrating(
							long			currentGrating);
	BOOL				GetAmOpen();
	void				SetAmOpen(
							BOOL			fAmOpen);
	long				GetNumberOfGratings();
	void				SetNumberOfGratings(
							long			numGratings);
	void				GetConfigFile(
							LPTSTR		*	szConfigFile);
	void				SetConfigFile(
							LPCTSTR			szConfigFile);
	BOOL				GetautoGrating();
	void				SetautoGrating(
							BOOL			autoGrating);
	void				GetINIFile(
							LPTSTR		*	szINIFile);
	void				SetINIFile(
							LPCTSTR			szINIFile);
	BOOL				GetAmBusy();
	void				GetDisplayName(
							LPTSTR		*	szDisplayName);
	void				SetDisplayName(
							LPCTSTR			szDisplayName);
	BOOL				GetWaitForComplete();
	void				SetWaitForComplete(
							BOOL			waitForComplete);
	void				GetMonoInfo(
							short int		Index,
							VARIANT		*	MonoInfo);
	BOOL				SetMonoInfo(
							short int		Index,
							VARIANT		*	newValue);
	void				GetGratingInfo(
							short int		gratingNumber,
							short int		Index,
							VARIANT		*	GratingInfo);
	BOOL				SetGratingInfo(
							short int		gratingNumber,
							short int		Index,
							VARIANT		*	newValue);
	double				GetmoveTime(
							double			newPosition);
	void				Setup();
	BOOL				SetGratingParams();
	BOOL				IsValidPosition(
							double			position);
	BOOL				ConvertStepsToNM(
							BOOL			fToNM,
							short int		gratingID,
							double		*	position,
							long		*	Steps);
	void				Abort();
	void				WriteConfig(
							LPCTSTR			Config);
	void				WriteINI(
							LPCTSTR			INIFile);
	double				GetCounter();
	BOOL				setInitialPositions(
							long			currentGrating, 
							BOOL			autoGrating, 
							double			CurrentPos);
	BOOL				SaveGratingZeroPosition(
							short int		Grating, 
							long			ZeroPos);
	void				doInit();
	BOOL				CanAutoSelect();
	// run current
	long				GetRunCurrent();
	void				SetRunCurrent(
							long			runCurrent);
	// idle current
	long				GetIdleCurrent();
	void				SetIdleCurrent(
							long			idleCurrent);
	long				GetGratingZeroOffset(
							long			Grating);
	void				SetGratingZeroOffset(
							long			Grating,
							long			zeroOffset);
	double				GetGratingStepsPerNM(
							long			Grating);
	void				SetGratingStepsPerNM(
							long			Grating,
							double			stepsPerNM);
	// high speed
	long				GetHighSpeed();
	void				SetHighSpeed(
							long			highSpeed);
	// SciArcus
	HWND				GetControlWindow();
	void				OnMoveCompleted();
	void				OnDeviceHomed();
	void				OnDeviceConnected();
	// set the busy flag
	void				SetAmBusy(
							BOOL			fAmBusy);
	// backlash correction
	BOOL				GetApplyBacklashCorrection();
	void				SetApplyBacklashCorrection(
							BOOL			fApplyBacklashCorrection);
	long				GetBacklashSteps();
	void				SetBacklashSteps(
							long			backlashSteps);
	// acceleration
	long				GetAcceleration();
	void				SetAcceleration(
							long			acceleration);
	// find grating corresponding to a given wavelength
	BOOL				FindGrating(
							double			wavelength,
							long		*	grating);
	CMyObject		*	GetMyObject();
	// setup window
	HWND				GetSetupWindow();
	void				SetSetupWindow(
							HWND			hwndSetup);
	// rapid scanning
	BOOL				GetRapidScanRunning();
	void				SetRapidScanRunning(
							BOOL			fRapidScanRunning);
	BOOL				StartRapidScan(
							long			grating,
							double			startWavelength,
							double			endWavelength,
							double			NMPerSecond);
	BOOL				_StartRapidScan(
							long			grating,
							double			startWavelength,
							double			endWavelength,
							double			NMPerSecond);
	void				OnRapidScanStepped();
	void				OnRapidScanClosed();
	// Grating dispersion calculation
	double				GetGratingDispersion(
							long			gratingID);
protected:
	struct GRATING_INFO;
	GRATING_INFO	*	GetGratingInfo(
							long			gratingIndex);
	CQuadratic		*	GetQuadratic(
							long			gratingIndex);
	BOOL				ConvertSteps(
							long			gratingID,
							long			steps,
							double		*	position);
	BOOL				ConvertPosition(
							long			gratingID,
							double			position,
							long		*	steps);
	// read mono info
	BOOL				ReadMonoInfo(
							CIniFile	*	pIniFile);
	BOOL				ReadGratingInfo(
							CIniFile	*	pIniFile,
							long			grating);
	BOOL				ReadMonoInfo(
							IDispatch	*	pdispConfigFile);
	BOOL				ReadGratingInfo(
							IDispatch	*	pdispConfigFile,
							long			grating);
	// get the steps per radian for a Direct Drive
	double				GetStepsPerRadian();
	// determine the period for a Direct Drive grating
	double				CalculatePeriod(
							long			grating);
	// obtain a parent window somehow
	HWND				GetParentWindow();
	// check the input number of steps
	BOOL				CheckSteps(
							long			steps);
	// configuration file object
	BOOL				GetConfigFile(
							IDispatch	**	ppdisp);
	BOOL				LoadConfigFile(
							IDispatch	*	pdisp,
							LPCTSTR			szConfigFile);
	long				GetSectionCount(
							IDispatch	*	pdispConfigFile);
	BOOL				GetSection(
							IDispatch	*	pdispConfigFile,
							long			index,
							IDispatch	**	ppdispSection);
	BOOL				GetSectionName(
							IDispatch	*	pdispSection,
							LPTSTR			szSection,
							UINT			nBufferSize);
	BOOL				GetSectionName(
							IDispatch	*	pdispConfigFile,
							long			Index,
							LPTSTR			szSection,
							UINT			nBufferSize);
	BOOL				FindNamedSection(
							IDispatch	*	pdispConfigFile,
							LPCTSTR			szSection,
							IDispatch	**	ppdispSection);
	long				GetParameterCount(
							IDispatch	*	pdispSection);
	BOOL				GetParameterName(
							IDispatch	*	pdispSection,
							long			index,
							LPTSTR			szParameter,
							UINT			nBufferSize);
	BOOL				GetParameter(
							IDispatch	*	pdispSection,
							long			Index,
							IDispatch	**	ppdispParameter);
	BOOL				GetParameterName(
							IDispatch	*	pdispParameter,
							LPTSTR			szParameter,
							UINT			nBufferSize);
	BOOL				FindNamedParameter(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							IDispatch	**	ppdispParameter);
	BOOL				GetParameterValue(
							IDispatch	*	pdispParameter,
							VARIANT		*	Value);
	BOOL				GetLongParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							long		*	retVal);
	BOOL				GetDoubleParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							double		*	retVal);
	BOOL				GetStringParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							LPTSTR			szString,
							UINT			nBufferSize);
	BOOL				GetParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSection,
							LPCTSTR			szParameter,
							VARIANT		*	Value);
	BOOL				GetStringParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							LPCTSTR			szDefault,
							LPTSTR			szOutput,
							UINT			nBufferSize);
	BOOL				GetLongParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							long			defValue,
							long		*	pvalue);
	BOOL				GetDoubleParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							double			defValue,
							double		*	pvalue);
	void				DetermineTempOffset(
							long			gratingID);		
	void				MoveToSteps(
							long			steps);
protected:
	// grating info structure
	struct GRATING_INFO
	{
		long			pitch;
		TCHAR			szBlaze[40];
		double			wavelengthRange[2];
		long			zeroPosition;
		long			location;
		double			phaseError;
		double			stepsPerNM;
		double			OffsetFactor;
		double			LinearFactor;
		double			QuadraticFactor;
		double			DualPassOffset;
		long			GetZeroPos() {return zeroPosition + location - tempZeroOffset;}
		double			ZeroPositionOffset;
		long			tempZeroOffset;
	};
private:
	CMyObject		*	m_pMyObject;
	CSciArcus		*	m_pSciArcus;
	// mono info
	LPTSTR				m_szModel;
	LPTSTR				m_szSerialNumber;
	LPTSTR				m_szDisplayName;
	LPTSTR				m_szDriveType;
	double				m_inputAngle;
	double				m_outputAngle;
	double				m_focalLength;
	double				m_minWave;
	double				m_maxWave;
	long				m_defaultPitch;
	long				m_idleCurrent;
	long				m_runCurrent;
	double				m_stepsPerRev;			// steps per revolution
	long				m_gearTeeth;			// gear teeth
	long				m_highSpeed;			// high speed
	long				m_acceleration;			// acceleration
	long				m_maxSteps;				// maximum steps for monochromator
	long				m_minSteps;				// minimum steps for monochromator
	// grating info
	long				m_numberOfGratings;
	GRATING_INFO	*	m_paGratingInfo;
	CQuadratic		**	m_paQuadratic;
	// config file
	LPTSTR				m_szConfigFile;
	// wait for completion flag
	BOOL				m_fWaitForComplete;
	// current grating
	long				m_currentGrating;
	// waiting for grating change
	BOOL				m_fWaitingForGratingChange;
	// waiting for wavelength change
	BOOL				m_fWaitingForPositionChange;
	// dummy dialog
	HWND				m_hwndDummy;
	// flag apply backlash correction
	BOOL				m_fApplyBacklashCorrection;
	long				m_backlashSteps;				// number of steps to go back
	// am initialized flag
	BOOL				m_fAmInitialized;
	// auto grating flag
	BOOL				m_fCanAutoSelect;
	BOOL				m_fAutoGrating;
	// setup window
	HWND				m_hwndSetup;
	// dirty flag
	BOOL				m_fDirty;
	// rapid scan dialog
	CDlgRapidScan	*	m_pDlgRapidScan;
	BOOL				m_fRapidScanSuccess;
	// start up values
	CStartupValues	*	m_pStartupValues;
	// motor ID
	TCHAR				m_szMotorID[MAX_PATH];

	// host object
	class CHostObject : public IDispatch
	{
	public:
		CHostObject(CMySciUsbMono * pMySciUsbMono);
		~CHostObject();
		// IUnknown methods
		STDMETHODIMP			QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv);
		STDMETHODIMP_(ULONG)	AddRef();
		STDMETHODIMP_(ULONG)	Release();
		// IDispatch methods
		STDMETHODIMP			GetTypeInfoCount( 
									PUINT			pctinfo);
		STDMETHODIMP			GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo);
		STDMETHODIMP			GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId);
		STDMETHODIMP			Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr); 
	protected:
		HRESULT					SendCommand(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult);
		HRESULT					doDelay(
									DISPPARAMS	*	pDispParams);
		void					doDelay(
									long			delayTimeInMS);
	private:
		CMySciUsbMono		*	m_pMySciUsbMono;
		ITypeInfo			*	m_pTypeInfo;
		ULONG					m_cRefs;
	};
	friend CHostObject;
};
