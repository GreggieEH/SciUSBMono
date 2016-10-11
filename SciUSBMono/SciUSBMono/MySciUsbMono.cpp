#include "stdafx.h"
#include "MySciUsbMono.h"
#include "MyObject.h"
#include "SciArcus.h"
#include "dispids.h"
#include "IniFile.h"
#include "Quadratic.h"
#include "MyGuids.h"
#include "dispids.h"
#include "DlgRapidScan.h"
#include "StartupValues.h"

CMySciUsbMono::CMySciUsbMono(CMyObject * pMyObject) :
	m_pMyObject(pMyObject),
	m_pSciArcus(NULL),
	// mono info
	m_szModel(NULL),
	m_szSerialNumber(NULL),
	m_szDisplayName(NULL),
	m_szDriveType(NULL),
	m_inputAngle(0.0),
	m_outputAngle(0.0),
	m_focalLength(0.0),
	m_minWave(0.0),
	m_maxWave(0.0),
	m_defaultPitch(0),
	m_idleCurrent(600),
	m_runCurrent(600),
	m_stepsPerRev(0.0),					// steps per revolution
	m_gearTeeth(200),					// gear teeth
	m_highSpeed(5000),					// high speed
	m_acceleration(10),					// acceleration
	m_maxSteps(0),						// maximum steps for monochromator
	m_minSteps(0),						// minimum steps for monochromator
	// grating info
	m_numberOfGratings(0),
	m_paGratingInfo(NULL),
	m_paQuadratic(NULL),
	// config file
	m_szConfigFile(NULL),
	// wait for completion flag
	m_fWaitForComplete(TRUE),
	// current grating
	m_currentGrating(1),
	// waiting for grating change
	m_fWaitingForGratingChange(FALSE),
	// waiting for wavelength change
	m_fWaitingForPositionChange(FALSE),
	// dummy dialog
	m_hwndDummy(NULL),
	// flag apply backlash correction
	m_fApplyBacklashCorrection(TRUE),
	m_backlashSteps(0),				// number of steps to go back
	// am initialized flag
	m_fAmInitialized(FALSE),
	// auto grating flag
	m_fCanAutoSelect(TRUE),
	m_fAutoGrating(TRUE),
	// setup window
	m_hwndSetup(NULL),
	// dirty flag
	m_fDirty(FALSE),
	// rapid scan dialog
	m_pDlgRapidScan(NULL),
	m_fRapidScanSuccess(FALSE),
	// start up values
	m_pStartupValues(NULL),
	m_ReInitOnScanStart(TRUE)			// flag reinitialization on scan start
{
	// default model and serial number
//	Utils_DoCopyString(&m_szModel, TEXT("9030"));
	SHStrDup(L"9055", &m_szModel);
//	Utils_DoCopyString(&m_szSerialNumber, TEXT("12345678"));
	SHStrDup(L"12345678", &m_szSerialNumber);
	// default motor ID
	StringCchCopy(m_szMotorID, MAX_PATH, L"jsa00");

}

CMySciUsbMono::~CMySciUsbMono(void)
{
	// remove the dummy dialog
	if (NULL != this->m_hwndDummy)
		DestroyWindow(this->m_hwndDummy);
	Utils_DELETE_POINTER(this->m_pDlgRapidScan);
	Utils_DELETE_POINTER(this->m_pSciArcus);
	if (NULL != this->m_szModel)
	{
		CoTaskMemFree((LPVOID)m_szModel);
		this->m_szModel = NULL;
	}
	if (NULL != this->m_szSerialNumber)
	{
		CoTaskMemFree((LPVOID)m_szSerialNumber);
		this->m_szSerialNumber = NULL;
	}
	if (NULL != this->m_szDriveType)
	{
		CoTaskMemFree((LPVOID)m_szDriveType);
		this->m_szDriveType = NULL;
	}

//	Utils_DoCopyString(&m_szModel, NULL);
//	Utils_DoCopyString(&m_szSerialNumber, NULL);
//	Utils_DoCopyString(&m_szDriveType, NULL);
	if (NULL != this->m_paQuadratic)
	{
		for (long i=0; i<this->m_numberOfGratings; i++)
		{
			Utils_DELETE_POINTER(this->m_paQuadratic[i]);
		}
		delete [] this->m_paQuadratic;
		this->m_paQuadratic	= NULL;
	}
	if (NULL != this->m_paGratingInfo)
		delete [] this->m_paGratingInfo;
	this->m_numberOfGratings	= 0;
//	Utils_DoCopyString(&m_szConfigFile, NULL);
	if (NULL != this->m_szConfigFile)
	{
		CoTaskMemFree((LPVOID)m_szConfigFile);
		this->m_szConfigFile = NULL;
	}

	// start up values
	Utils_DELETE_POINTER(this->m_pStartupValues);
}

// persist to stream current grating, wavelength and auto grating flag

HRESULT CMySciUsbMono::InitNew()
{
	return S_OK;
}

long CMySciUsbMono::GetSaveSize()
{
	return (sizeof(long) + sizeof(double) + sizeof(long)) * 2;
}

HRESULT	CMySciUsbMono::Save(
							IStream		*	pStm, 
							BOOL			fClearDirty)
{
	HRESULT			hr;
	long			grating		= this->GetcurrentGrating();
	double			position	= this->Getposition();
	BOOL			autoGrating	= this->GetautoGrating();

	hr = IStream_Write(pStm, (const void*) &grating, sizeof(long));
	if (SUCCEEDED(hr)) hr = IStream_Write(pStm, (const void*) &position, sizeof(double));
	if (SUCCEEDED(hr)) hr = IStream_Write(pStm, (const void*) &autoGrating, sizeof(BOOL));
	if (SUCCEEDED(hr) && fClearDirty) this->m_fDirty	= FALSE;
	return hr;
}

HRESULT	CMySciUsbMono::Load(	
							IStream		*	pStm)
{
	HRESULT			hr;
	long			grating;
	double			position;
	BOOL			autoGrating;

	hr = IStream_Read(pStm, (void*) &grating, sizeof(long));
	if (SUCCEEDED(hr)) hr = IStream_Read(pStm, (void*) &position, sizeof(double));
	if (SUCCEEDED(hr)) hr = IStream_Read(pStm, (void*) &autoGrating, sizeof(BOOL));
	if (SUCCEEDED(hr))
	{
		// start up values
		Utils_DELETE_POINTER(this->m_pStartupValues);
		this->m_pStartupValues	= new CStartupValues(grating, position, autoGrating);
	}
	return S_OK;
}

BOOL CMySciUsbMono::GetDirty()
{
	return this->m_fDirty;
}

HRESULT CMySciUsbMono::Load(
							IPropertyBag*	pPropBag, 
							IErrorLog	*	pErrorLog)
{
	return S_OK;
}

HRESULT CMySciUsbMono::Save(
							IPropertyBag*	pPropBag,
							BOOL			fClearDirty,
							BOOL			fSaveAllProperties)
{
	return S_OK;
}

BOOL CMySciUsbMono::PaintToRect(
							HDC				hdc, 
							LPCRECT			prc)
{
	HBRUSH			hbrBackground;

	hbrBackground = CreateSolidBrush(RGB(192, 192, 192));
	FillRect(hdc, prc, hbrBackground);
	DeleteObject((HGDIOBJ) hbrBackground);
	return TRUE;
}

// properties and methods
double CMySciUsbMono::Getposition()
{
	double				position		= 0.0;
	long				steps;
//	CQuadratic		*	pQuadratic;
	TCHAR				szMessage[MAX_PATH];

	if (NULL != this->m_pSciArcus && this->m_pSciArcus->GetCurrentPosition(&steps))
	{
StringCchPrintf(szMessage, MAX_PATH, TEXT("In Getposition Step position = %1d"),
				steps);
DoLogString(szMessage);
		this->ConvertStepsToNM(TRUE, (short) this->GetcurrentGrating(), &position, &steps);
_stprintf_s(szMessage, MAX_PATH, TEXT("In Getposition nm = %7.1f"), position);
DoLogString(szMessage);
/*
		// apply the quadratic correction
		pQuadratic = this->GetQuadratic(this->GetcurrentGrating());
		if (NULL != pQuadratic)
			position = pQuadratic->ExpToActual(position);
*/
	}
	return position;
}

void CMySciUsbMono::Setposition(
							double			position)
{
	long				steps;
//	CQuadratic		*	pQuadratic;
	long				grating;
	TCHAR				szError[MAX_PATH];
	BOOL				fChangedGrating		= FALSE;

	if (this->GetAmBusy()) return;
	// check if the desired position is valid
//	if (!this->IsValidPosition(position))
//		return;
	// check if position is valid
	if (!FindGrating(position, &grating))
	{
		_stprintf_s(szError, MAX_PATH, TEXT("Position %8.1f Is Out Of Range"), position);
		this->m_pMyObject->FireError(szError);
		return;
	}
	// allow the client to disallow the move
	if (!this->m_pMyObject->FireQueryAllowChangePosition(position))
		return;
	if (NULL != this->m_pSciArcus)
	{
		// do we need to change the grating?
		if (this->GetcurrentGrating() != grating)
		{
			if (!this->SetcurrentGrating(grating))
				return;
			fChangedGrating	= TRUE;
		}
/*
		// apply the correction factor
		pQuadratic = this->GetQuadratic(this->GetcurrentGrating());
		if (NULL != pQuadratic)
			position = pQuadratic->ActualToExp(position);
*/
		if (fChangedGrating)
		{
			DoLogString(TEXT("In Setposition, found actual position"));
		}

		// convert to steps
		if (this->ConvertStepsToNM(FALSE, (long) this->GetcurrentGrating(), &position, &steps))
		{
			if (fChangedGrating)
			{
				DoLogString(TEXT("In Setposition, found step position"));
			}

			// check the input steps
			if (!this->CheckSteps(steps))
			{
				_stprintf_s(szError, MAX_PATH, TEXT("Position %8.1f Is Out Of Range"), position);
				this->m_pMyObject->FireError(szError);
				return;
			}
			this->MoveToSteps(steps);			// move to the number of steps
//			this->m_pSciArcus->MoveToPosition(steps);
			Utils_OnPropChanged(this->m_pMyObject,  DISPID_position);
			this->m_pMyObject->FireHaveNewPosition(this->Getposition());
			this->m_fDirty	= TRUE;
			if (NULL != this->m_pStartupValues)
			{
				Utils_DELETE_POINTER(this->m_pStartupValues);
			}
		}
	}
}

void CMySciUsbMono::MoveToSteps(
	long			steps)
{
	long		backlashSteps = this->GetBacklashSteps();
	long		currentSteps;
	long		backlashPoint;
	if (this->GetApplyBacklashCorrection() && this->m_pSciArcus->GetCurrentPosition(&currentSteps) && steps < currentSteps)
	{
		// move to the backlash point
		backlashPoint = steps - backlashSteps;
		this->m_pSciArcus->MoveToPosition(backlashPoint);
	}
	this->m_pSciArcus->MoveToPosition(steps);
}


// check the input number of steps
BOOL CMySciUsbMono::CheckSteps(
							long			steps)
{
	BOOL				fPositiveSense;			// flag positive sense

	// if the maximum number of steps is 0, this has no effect
	if (0 == this->m_maxSteps) return TRUE;
	if (this->m_maxSteps < 0)
	{
		fPositiveSense = FALSE;
		// maximum number of steps less than zero
		if (steps < this->m_maxSteps)
			return FALSE;
	}
	else
	{
		fPositiveSense = TRUE;
		// maximum number of steps greater than zero
		if (steps > this->m_maxSteps)
			return FALSE;
	}
	// check that the minimum steps is in range
	if (fPositiveSense)
		return steps >= this->m_minSteps;
	else
		return steps <= this->m_minSteps;
}


long CMySciUsbMono::GetcurrentGrating()
{
	return m_currentGrating;
}

BOOL CMySciUsbMono::SetcurrentGrating(
							long			currentGrating)
{
	TCHAR				szMessage[MAX_PATH];

	DoLogString(TEXT("In changing Grating"));

	if (NULL == this->m_pSciArcus) return FALSE;
	DoLogString(TEXT("Have SciArcus"));
	if (this->GetAmBusy()) return FALSE;
	DoLogString(TEXT("Not Busy"));
	if (currentGrating != this->m_currentGrating)
	{
		GRATING_INFO	*	pGratingInfo	= this->GetGratingInfo(currentGrating);
		if (NULL != pGratingInfo)
		{
DoLogString(TEXT("Have Grating Info"));
			if (!this->m_pMyObject->FireRequestChangeGrating(currentGrating))
			{
				DoLogString(TEXT("Grating change dis allowed"));
				this->m_pMyObject->FireError(TEXT("Grating change dis allowed"));
//				return FALSE;
			}
//			DoLogString(TEXT("In changing Grating"));

			long			currentPosition;
			this->m_pSciArcus->GetCurrentPosition(&currentPosition);
			// grating zero position
			StringCchPrintf(szMessage, MAX_PATH,
				TEXT("In changing Grating, current position = %d, new position = %d"),
				currentPosition, pGratingInfo->GetZeroPos());
			DoLogString(szMessage);
			
			// move to the zero position of the new grating
			this->m_currentGrating = currentGrating;
			while (this->m_pSciArcus->GetMotorMoving())
			{
				DoLogString(TEXT("Waiting for motor to stop moving"));
				MyYield();
			}
			this->MoveToSteps(pGratingInfo->GetZeroPos());
	//		this->m_pSciArcus->MoveToPosition(pGratingInfo->GetZeroPos());

			DoLogString(TEXT("Have changed Grating"));
			// sink notifications
			Utils_OnPropChanged(this->m_pMyObject, DISPID_currentGrating);
			this->m_pMyObject->FireChangedGrating(this->m_currentGrating);
//			this->m_fWaitingForGratingChange	= TRUE;
//			this->SetAmBusy(TRUE);
			this->m_fDirty		= TRUE;
			if (NULL != this->m_pStartupValues)
			{
				this->Setposition(this->m_pStartupValues->GetPosition());
			}
		}
		else
		{
DoLogString(TEXT("Don't have grating info"));
		}
	}
	return TRUE;
}

BOOL CMySciUsbMono::GetAmOpen()
{
	if (NULL != this->m_pSciArcus)
	{
		return this->m_fAmInitialized;
	}
	else
		return FALSE;
}

void CMySciUsbMono::SetAmOpen(
							BOOL			fAmOpen)
{
	if (fAmOpen)
	{
		// make sure that the SciArcus control is set
		if (NULL == this->m_pSciArcus)
		{
			this->m_pSciArcus	= new CSciArcus(this, this->m_szMotorID);
			// set the position range
			if (0 != this->m_maxSteps)
				this->m_pSciArcus->SetMaximumPosition(this->m_maxSteps);
			if (0 != this->m_minSteps)
				this->m_pSciArcus->SetMinimumPosition(this->m_minSteps);

			if (!this->m_pSciArcus->doInit())
			{
				this->m_pMyObject->FireError(TEXT("Failed to Load SciArcus.exe"));
				return;
			}
			// attempt to load the DLL
			this->m_pSciArcus->SetDllLoaded();
			if (!this->m_pSciArcus->GetDllLoaded())
			{
				this->m_pMyObject->FireError(TEXT("Failed to Load Arcus DLL"));
				return;
			}
			// connect the device
			if (!this->m_pSciArcus->FindDeviceIndex())
			{
				this->m_pMyObject->FireError(TEXT("Could not find Arcus Device!"));
				return;
			}
		}
		// home the device
		this->m_pSciArcus->SetDeviceConnected(TRUE);
		this->m_pMyObject->FireBusyStatusChange(TRUE);
		this->m_pMyObject->FireStatusMessage(TEXT("Initializing Mono"), TRUE);
		if (this->m_pSciArcus->MotorHomed())
		{
			this->m_fAmInitialized	= TRUE;
			this->m_pMyObject->FireAmInitPropChanged(TRUE);
			this->OnDeviceHomed();
		}
		this->m_pMyObject->FireStatusMessage(TEXT("Ready"), FALSE);
	}
	else
	{
		Utils_DELETE_POINTER(this->m_pSciArcus);
		this->m_fAmInitialized	= FALSE;
		this->m_pMyObject->FireAmInitPropChanged(FALSE);
	}
	Utils_OnPropChanged(this->m_pMyObject, DISPID_AmOpen);
}

long CMySciUsbMono::GetNumberOfGratings()
{
	return this->m_numberOfGratings;
}

void CMySciUsbMono::SetNumberOfGratings(
							long			numGratings)
{
	long				i;
	if (numGratings != this->m_numberOfGratings)
	{
		if (NULL != this->m_paQuadratic)
		{
			for (i=0; i<this->m_numberOfGratings; i++)
			{
				Utils_DELETE_POINTER(this->m_paQuadratic[i]);
			}
			delete [] this->m_paQuadratic;
			this->m_paQuadratic	= NULL;
		}
		if (NULL != this->m_paGratingInfo)
		{
			delete [] this->m_paGratingInfo;
		}
		this->m_numberOfGratings	= numGratings;
		this->m_paGratingInfo	= new GRATING_INFO [this->m_numberOfGratings];
		this->m_paQuadratic		= new CQuadratic* [this->m_numberOfGratings];
		for (i=0; i<this->m_numberOfGratings; i++)
		{
			this->m_paQuadratic[i]	= NULL;
			ZeroMemory((PVOID)&(this->m_paGratingInfo[i]), sizeof(GRATING_INFO));
		}
	}
}

void CMySciUsbMono::GetConfigFile(
							LPTSTR		*	szConfigFile)
{
	*szConfigFile	= NULL;
//	Utils_DoCopyString(szConfigFile, this->m_szConfigFile);
	SHStrDup(this->m_szConfigFile, szConfigFile);
}

void CMySciUsbMono::SetConfigFile(
							LPCTSTR			szConfigFile)
{
	long				grating;
	TCHAR				szFilePath[MAX_PATH];

	if (PathIsRelative(szConfigFile))
	{
		GetModuleFileName(GetOurInstance(), szFilePath, MAX_PATH);
		PathAppend(szFilePath, szConfigFile);
	}
	else
	{
		StringCchCopy(szFilePath, MAX_PATH, szConfigFile);
	}

//	Utils_DoCopyString(&m_szConfigFile, szFilePath);
	SHStrDup(szFilePath, &m_szConfigFile);
	if (NULL != StrStrI(szConfigFile, TEXT("SciRunOS")))
	{
		// read the config file using ConfigFile dll
		IDispatch		*	pdisp;
		if (this->GetConfigFile(&pdisp))
		{
			if (this->LoadConfigFile(pdisp, this->m_szConfigFile))
			{
				if (!this->ReadMonoInfo(pdisp))
				{
					this->m_pMyObject->FireError(TEXT("Couldn;t read configuration file"));
					return;
				}
				// loop over the gratings
				for (grating = 1; grating <= this->m_numberOfGratings; grating++)
				{
					this->ReadGratingInfo(pdisp, grating);
				}
			}
			pdisp->Release();
		}
	}
	else
	{
		CIniFile			iniFile;
		// read the config file
		if (iniFile.Init(this->m_szConfigFile))
		{
			if (!this->ReadMonoInfo(&iniFile))
			{
				this->m_pMyObject->FireError(TEXT("Couldn;t read configuration file"));
				return;
			}
			// loop over the gratings
			for (grating = 1; grating <= this->m_numberOfGratings; grating++)
			{
				this->ReadGratingInfo(&iniFile, grating);
			}
		}
	}
}

// read mono info
BOOL CMySciUsbMono::ReadMonoInfo(
							CIniFile	*	pIniFile)
{
	TCHAR				szValue[MAX_PATH];
	long				lVal;

	// model
	if (pIniFile->ReadStringValue(TEXT("Monochromator"), TEXT("Model"), TEXT("9030"),
		szValue, MAX_PATH))
	{
		StrTrim(szValue, TEXT(" "));
//		Utils_DoCopyString(&m_szModel, szValue);
		SHStrDup(szValue, &m_szModel);
	}
	// serial number
	if (pIniFile->ReadStringValue(TEXT("Monochromator"), TEXT("Serial"), TEXT("1234"), 
		szValue, MAX_PATH))
	{
		StrTrim(szValue, TEXT(" "));
//		Utils_DoCopyString(&(this->m_szSerialNumber), szValue);
		SHStrDup(szValue, &m_szSerialNumber);
	}
	pIniFile->ReadIntValue(TEXT("Monochromator"), TEXT("DefaultPitch"),
		1200, &(this->m_defaultPitch));
	if (pIniFile->ReadStringValue(TEXT("Monochromator"), TEXT("MotorID"), TEXT("jsa00"), 
		szValue, MAX_PATH))
	{
		StrTrim(szValue, TEXT(" "));
		StringCchCopy(this->m_szMotorID, MAX_PATH, szValue);
	}
/*
	double				m_inputAngle;
	double				m_outputAngle;
	double				m_focalLength;
	double				m_minWave;
	double				m_maxWave;
*/
	pIniFile->ReadDoubleValue(TEXT("Monochromator"), TEXT("InputAngle"), 6.912, 
		&(this->m_inputAngle));
	pIniFile->ReadDoubleValue(TEXT("Monochromator"), TEXT("OutputAngle"), 6.18,
		&(this->m_outputAngle));
	pIniFile->ReadDoubleValue(TEXT("Monochromator"), TEXT("FocalLength"), 550.0,
		&(this->m_focalLength));
	pIniFile->ReadDoubleValue(TEXT("Monochromator"), TEXT("MinWave"), -100.0,
		&(this->m_minWave));
	pIniFile->ReadDoubleValue(TEXT("Monochromator"), TEXT("MaxWave"), 1200.0,
		&(this->m_maxWave));
	if (pIniFile->ReadIntValue(TEXT("Monochromator"), TEXT("AutoGrating"), 1, &lVal))
	{
		this->m_fCanAutoSelect = 0 != lVal;
	}
	// USB mono section
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("IdleCurrent"), 600, &(this->m_idleCurrent));
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("RunCurrent"), 600, &(this->m_runCurrent));
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("GearTeeth"), 200, &(this->m_gearTeeth));
	pIniFile->ReadDoubleValue(TEXT("USB Mono"), TEXT("StepsPerRev"), 80.0, &(this->m_stepsPerRev));
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("NumGratings"), 1, &lVal);
	this->SetNumberOfGratings(lVal);
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("DriveType"), 0, &lVal);
//	Utils_DoCopyString(&(this->m_szDriveType), 0 == lVal ? TEXT("SineDrive") : TEXT("DirectDrive"));
	SHStrDup(0 == lVal ? L"SineDrive" : L"DirectDrive", &m_szDriveType);
	// high speed
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("HighSpeed"), 5000, &(this->m_highSpeed));
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("Acceleration"), 10, &(this->m_acceleration));
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("MaxSteps"), 0, &(this->m_maxSteps));
//	if (0 != this->m_maxSteps)
//		this->m_pSciArcus->SetMaximumPosition(this->m_maxSteps);
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("MinSteps"), 0, &(this->m_minSteps));
//	if (0 != this->m_minSteps)
//		this->m_pSciArcus->SetMinimumPosition(this->m_minSteps);
	pIniFile->ReadIntValue(TEXT("USB Mono"), TEXT("ApplyBacklash"), 0, &lVal);
	this->m_fApplyBacklashCorrection	= 0 != lVal;
	pIniFile->ReadIntValue(L"USB Mono", L"BacklashSteps", 0, &lVal);
	this->SetBacklashSteps(lVal);
	return TRUE;
}

BOOL CMySciUsbMono::ReadMonoInfo(
							IDispatch	*	pdispConfigFile)
{
	TCHAR				szValue[MAX_PATH];
	long				lVal;

	// model
	if (this->GetStringParameterValue(pdispConfigFile,
		TEXT("Monochromator"), TEXT("Model"), TEXT("9030"),
		szValue, MAX_PATH))
	{
		StrTrim(szValue, TEXT(" "));
//		Utils_DoCopyString(&m_szModel, szValue);
		SHStrDup(szValue, &m_szModel);
	}
	// serial number
	if (this->GetStringParameterValue(pdispConfigFile, TEXT("Monochromator"), 
		TEXT("Serial"), TEXT("1234"), 
		szValue, MAX_PATH))
	{
		StrTrim(szValue, TEXT(" "));
//		Utils_DoCopyString(&(this->m_szSerialNumber), szValue);
		SHStrDup(szValue, &m_szSerialNumber);
	}
	GetLongParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("DefaultPitch"),
		1200, &(this->m_defaultPitch));
/*
	double				m_inputAngle;
	double				m_outputAngle;
	double				m_focalLength;
	double				m_minWave;
	double				m_maxWave;
*/
	GetDoubleParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("InputAngle"), 6.912, 
		&(this->m_inputAngle));
	GetDoubleParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("OutputAngle"), 6.18,
		&(this->m_outputAngle));
	GetDoubleParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("FocalLength"), 550.0,
		&(this->m_focalLength));
	GetDoubleParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("MinWave"), -100.0,
		&(this->m_minWave));
	GetDoubleParameterValue(pdispConfigFile, TEXT("Monochromator"), TEXT("MaxWave"), 1200.0,
		&(this->m_maxWave));
	if (GetLongParameterValue(pdispConfigFile, TEXT("Monochromator"), 
			TEXT("AutoGrating"), 0, &lVal))
	{
		this->m_fCanAutoSelect = 0 != lVal;
	}
	// Gratings section
	/**************************
	// specific for the USB Mono:
	**************************/
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("IdleCurrent"), 600, &(this->m_idleCurrent));
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("RunCurrent"), 600, &(this->m_runCurrent));
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("HighSpeed"), 5000, &(this->m_highSpeed));
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("Acceleration"), 10, &(this->m_acceleration));
	/**************************
	// common for the monochromators
	**************************/
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("GearTeeth"), 200, &(this->m_gearTeeth));
	this->GetDoubleParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("StepsPerRev"), 80.0, &(this->m_stepsPerRev));
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("Gratings"), 1, &lVal);
	this->SetNumberOfGratings(lVal);
	// drive type
	this->GetStringParameterValue(pdispConfigFile, TEXT("Gratings"), 
		TEXT("DriveType"), TEXT("Direct"), 
		szValue, MAX_PATH);
//	Utils_DoCopyString(&(this->m_szDriveType), 0 == lstrcmpi(szValue, TEXT("Direct")) ?
//		TEXT("DirectDrive") : TEXT("SineDrive"));
	SHStrDup((0 == lstrcmpi(szValue, L"Direct") ? L"DirectDrive" : L"SineDrive"), &m_szDriveType);


//	this->GetLongParameterValue(pdispConfigFile, TEXT("USB Mono"), TEXT("DriveType"), 0, &lVal);
//	Utils_DoCopyString(&(this->m_szDriveType), 0 == lVal ? TEXT("SineDrive") : TEXT("DirectDrive"));
	// high speed
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("StepsToEnd"), 0, &(this->m_maxSteps));
//	if (0 != this->m_maxSteps)
//		this->m_pSciArcus->SetMaximumPosition(this->m_maxSteps);
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), TEXT("MinSteps"), 0, &(this->m_minSteps));
//	if (0 != this->m_minSteps)
//		this->m_pSciArcus->SetMinimumPosition(this->m_minSteps);
	this->GetLongParameterValue(pdispConfigFile, TEXT("Gratings"), 
		TEXT("ApplyBacklash"), 0, &lVal);
	this->m_fApplyBacklashCorrection	= 0 != lVal;
	this->GetLongParameterValue(pdispConfigFile, L"Gratings", L"BacklashSteps",
		0, &lVal);
	this->SetBacklashSteps(lVal);
	return TRUE;
}

// acceleration
long CMySciUsbMono::GetAcceleration()
{
	return this->m_acceleration;
}

void CMySciUsbMono::SetAcceleration(
							long			acceleration)
{
	this->m_acceleration = acceleration;
}

// high speed
long CMySciUsbMono::GetHighSpeed()
{
	return this->m_highSpeed;
}

void CMySciUsbMono::SetHighSpeed(
							long			highSpeed)
{
	this->m_highSpeed	= highSpeed;
//	this->m_fDirty		= TRUE;
	if (NULL != this->m_pSciArcus)
		this->m_pSciArcus->SetHighSpeed(this->m_highSpeed);
}


BOOL CMySciUsbMono::ReadGratingInfo(
							CIniFile	*	pIniFile,
							long			grating)
{
	TCHAR			szSection[MAX_PATH];
	GRATING_INFO*	pGratingInfo	=	this->GetGratingInfo(grating);
//	CQuadratic	*	pQuadratic;
	if (NULL == pGratingInfo)
		return FALSE;

	StringCchPrintf(szSection, MAX_PATH, TEXT("Grating %1d"), grating);
	pIniFile->ReadIntValue(szSection, TEXT("Pitch"), 1200, &(pGratingInfo->pitch));
	pIniFile->ReadStringValue(szSection, TEXT("Blaze"), TEXT("UV"), 
		pGratingInfo->szBlaze, 40);
	pIniFile->ReadDoubleValue(szSection, TEXT("MinEffWave"), 
		500.0, &(pGratingInfo->wavelengthRange[0]));
	pIniFile->ReadDoubleValue(szSection, TEXT("MaxEffWave"),
		1200.0, &(pGratingInfo->wavelengthRange[1]));
	pIniFile->ReadIntValue(szSection, TEXT("ZeroPos"), 0, &(pGratingInfo->zeroPosition));
	pIniFile->ReadIntValue(szSection, TEXT("Location"), 0, &(pGratingInfo->location));
	pIniFile->ReadDoubleValue(szSection, TEXT("PhaseErr"), 0.0,
		&(pGratingInfo->phaseError));
	pIniFile->ReadDoubleValue(szSection, TEXT("StepsPerNM"), 50.0,
		&(pGratingInfo->stepsPerNM));
	pIniFile->ReadDoubleValue(szSection, TEXT("Correction Offset"), 0.0,
		&(pGratingInfo->OffsetFactor));
	pIniFile->ReadDoubleValue(szSection, TEXT("Correction Linear"), 1.0,
		&(pGratingInfo->LinearFactor));
	pIniFile->ReadDoubleValue(szSection, TEXT("Correction Quadratic"), 0.0,
		&(pGratingInfo->QuadraticFactor));
	// form the quadratic correction
	if (NULL != this->m_paQuadratic)
	{
		if (NULL == this->m_paQuadratic[grating - 1])
			this->m_paQuadratic[grating - 1]	= new CQuadratic();
		this->m_paQuadratic[grating-1]->SetLinearFactor(pGratingInfo->LinearFactor);
		this->m_paQuadratic[grating-1]->SetOffsetFactor(pGratingInfo->OffsetFactor);
		this->m_paQuadratic[grating-1]->SetQuadFactor(pGratingInfo->QuadraticFactor);
		this->m_paQuadratic[grating-1]->SetMaxWavelength(pGratingInfo->wavelengthRange[1]);
		this->m_paQuadratic[grating-1]->SetMinWavelength(-100.0);
		this->m_paQuadratic[grating-1]->MakeCorrections();
	}
	pIniFile->ReadDoubleValue(szSection, TEXT("DualPassOffset"), 0.0,
		&(pGratingInfo->DualPassOffset));

	return TRUE;
}


BOOL CMySciUsbMono::ReadGratingInfo(
							IDispatch	*	pdispConfigFile,
							long			grating)
{
	TCHAR			szSection[MAX_PATH];
	GRATING_INFO*	pGratingInfo	=	this->GetGratingInfo(grating);
//	CQuadratic	*	pQuadratic;
	if (NULL == pGratingInfo)
		return FALSE;

	StringCchPrintf(szSection, MAX_PATH, TEXT("Grating#%1d"), grating);
	this->GetLongParameterValue(pdispConfigFile, szSection, TEXT("Pitch"), 1200, &(pGratingInfo->pitch));
	this->GetStringParameterValue(pdispConfigFile, szSection, TEXT("Blaze"), TEXT("UV"), 
		pGratingInfo->szBlaze, 40);
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("MinEffWave"), 
		500.0, &(pGratingInfo->wavelengthRange[0]));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("MaxEffWave"),
		1200.0, &(pGratingInfo->wavelengthRange[1]));
	this->GetLongParameterValue(pdispConfigFile, szSection, TEXT("ZeroPos"), 0, &(pGratingInfo->zeroPosition));
	this->GetLongParameterValue(pdispConfigFile, szSection, TEXT("Location"), 0, &(pGratingInfo->location));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("PhaseErr"), 0.0,
		&(pGratingInfo->phaseError));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("StepsPerNM"), 50.0,
		&(pGratingInfo->stepsPerNM));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("Correction Offset"), 0.0,
		&(pGratingInfo->OffsetFactor));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("Correction Linear"), 1.0,
		&(pGratingInfo->LinearFactor));
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("Correction Quadratic"), 0.0,
		&(pGratingInfo->QuadraticFactor));
	// form the quadratic correction
	if (NULL != this->m_paQuadratic)
	{
		if (NULL == this->m_paQuadratic[grating - 1])
			this->m_paQuadratic[grating - 1]	= new CQuadratic();
		this->m_paQuadratic[grating-1]->SetLinearFactor(pGratingInfo->LinearFactor);
		this->m_paQuadratic[grating-1]->SetOffsetFactor(pGratingInfo->OffsetFactor);
		this->m_paQuadratic[grating-1]->SetQuadFactor(pGratingInfo->QuadraticFactor);
		this->m_paQuadratic[grating-1]->SetMaxWavelength(pGratingInfo->wavelengthRange[1]);
		this->m_paQuadratic[grating-1]->SetMinWavelength(-100.0);
		this->m_paQuadratic[grating-1]->MakeCorrections();
	}
	this->GetDoubleParameterValue(pdispConfigFile, szSection, TEXT("DualPassOffset"), 0.0,
		&(pGratingInfo->DualPassOffset));

	return TRUE;

}

BOOL CMySciUsbMono::GetautoGrating()
{
	return this->m_fAutoGrating && this->CanAutoSelect();
}

void CMySciUsbMono::SetautoGrating(
							BOOL			autoGrating)
{
	if (!this->CanAutoSelect())
	{
		this->m_fAutoGrating	= FALSE;
		return;
	}
	this->m_fAutoGrating = autoGrating;
	Utils_OnPropChanged(this->m_pMyObject, DISPID_autoGrating);
	this->m_pMyObject->FireAutoGratingPropChanged(this->m_fAutoGrating);
	this->m_fDirty = TRUE;
}

void CMySciUsbMono::GetINIFile(
							LPTSTR		*	szINIFile)
{
	*szINIFile	= NULL;
}

void CMySciUsbMono::SetINIFile(
							LPCTSTR			szINIFile)
{
}

BOOL CMySciUsbMono::GetAmBusy()
{
	if (NULL != this->m_pSciArcus)
	{
		return this->m_pSciArcus->GetMotorMoving();
	}
	return FALSE;
}

void CMySciUsbMono::GetDisplayName(
							LPTSTR		*	szDisplayName)
{
	*szDisplayName	= NULL;
	if (NULL != this->m_szDisplayName)
		SHStrDup(this->m_szDisplayName, szDisplayName);
//		Utils_DoCopyString(szDisplayName, this->m_szDisplayName);
}

void CMySciUsbMono::SetDisplayName(
							LPCTSTR			szDisplayName)
{
//	Utils_DoCopyString(&m_szDisplayName, szDisplayName);
	SHStrDup(szDisplayName, &m_szDisplayName);
	Utils_OnPropChanged(this->m_pMyObject, DISPID_DisplayName);
}

BOOL CMySciUsbMono::GetWaitForComplete()
{
	return this->m_fWaitForComplete;
}

void CMySciUsbMono::SetWaitForComplete(
							BOOL			waitForComplete)
{
	this->m_fWaitForComplete	= waitForComplete;
}

void CMySciUsbMono::GetMonoInfo(
							short int		Index,
							VARIANT		*	MonoInfo)
{
	VariantInit(MonoInfo);
	switch (Index)
	{
	case 0:
		InitVariantFromString(this->m_szModel, MonoInfo);
		// model info
		break;
	case 1:
		// serial number
		InitVariantFromString(this->m_szSerialNumber, MonoInfo);
		break;
	case 2:
		// resolution
		break;
	case 3:
		// default pitch
		InitVariantFromInt32(this->m_defaultPitch, MonoInfo);
		break;
	case 4:
		// max wave
		InitVariantFromDouble(this->m_maxWave, MonoInfo);
		break;
	case 5:
		// min wave
		InitVariantFromDouble(this->m_minWave, MonoInfo);
		break;
	case 6:
		// nm step resolution
		break;
	case 7:
		// input angle
		InitVariantFromDouble(this->m_inputAngle, MonoInfo);
		break;
	case 8:
		// output angle
		InitVariantFromDouble(this->m_outputAngle, MonoInfo);
		break;
	case 9:
		// focal length
		break;
	case 10:
		// init mode
		break;
	case 11:
		// drive motor
		break;
	case 12:
		// NM dec places
		break;
	case 13:
		// num motors
		break;
	case 14:
		// Drive type
		InitVariantFromString(this->m_szDriveType, MonoInfo);
		break;
	case 15:
		// gear teeth
		InitVariantFromInt32(this->m_gearTeeth, MonoInfo);
		break;
	case 16:
		// switch motor
		break;
	case MONO_INFO_IDLECURRENT:
		InitVariantFromInt32(this->m_idleCurrent, MonoInfo);
		break;
	case MONO_INFO_RUNCURRENT:
		InitVariantFromInt32(this->m_runCurrent, MonoInfo);
		break;
	case MONO_INFO_HIGHSPEED:
		InitVariantFromInt32(this->m_highSpeed, MonoInfo);
		break;
	case MONO_INFO_STEPSPERREV:
		InitVariantFromDouble(this->m_stepsPerRev, MonoInfo);
		break;
	case MONO_INFO_NUMGRATINGS:
		InitVariantFromInt32(this->GetNumberOfGratings(), MonoInfo);
		break;
	default:
		break;
	}
/*
    Select Case Index

        Case 0
            GetMonoInfo = CVar(mono.Model)
        Case 1
            GetMonoInfo = CVar(mono.Serial)
        Case 2
            GetMonoInfo = CVar(mono.Resolution)
        Case 3
            GetMonoInfo = CVar(mono.DefaultPitch)
        Case 4
            GetMonoInfo = CVar(mono.MaxWave)
        Case 5
            GetMonoInfo = CVar(mono.minWave)
        Case 6
            GetMonoInfo = CVar(mono.NmStepRes)
        Case 7
            GetMonoInfo = CVar(mono.InputAngle)
        Case 8
            GetMonoInfo = CVar(mono.OutputAngle)
        Case 9
            GetMonoInfo = CVar(mono.FocalLength)
        Case 10                                     ' init mode
            GetMonoInfo = CVar(mono.InitMode)
        Case 11
            GetMonoInfo = CVar(mono.DriveMotor)     ' drive motor
        Case 12                                     ' nm decimal places
            GetMonoInfo = CVar(mono.NmDecPlaces)
        Case 13
            GetMonoInfo = CVar(mono.NumMotors)
        Case 14
            If mono.DriveType = SINE_DRIVE Then
                GetMonoInfo = CVar("Sine Drive")
            Else
                GetMonoInfo = CVar("Direct Drive")
            End If
        Case 15                     ' gear teeth
            GetMonoInfo = CVar(mono.GearTeeth)
        Case 16
            GetMonoInfo = CVar(mono.SwitchMotor)
        Case Else
            GetMonoInfo = CVar(0)
    End Select

*/
}

BOOL CMySciUsbMono::SetMonoInfo(
							short int		Index,
							VARIANT		*	newValue)
{
	HRESULT			hr;
	long			lval;
	double			dval;
	LPTSTR			szString		= NULL;

	switch (Index)
	{
	case MONO_INFO_IDLECURRENT:
		hr = VariantToInt32(*newValue, &lval);
		if (SUCCEEDED(hr))
		{
			this->m_idleCurrent	= lval;
//			this->m_fDirty		= TRUE;
		}
		break;
	case MONO_INFO_RUNCURRENT:
		hr = VariantToInt32(*newValue, &lval);
		if (SUCCEEDED(hr))
		{
			this->m_runCurrent	= lval;
//			this->m_fDirty		= TRUE;
		}
		break;
	case MONO_INFO_HIGHSPEED:
		hr = VariantToInt32(*newValue, &lval);
		if (SUCCEEDED(hr))
		{
			this->m_highSpeed	= lval;
//			this->m_fDirty		= TRUE;
		}
		break;
	case MONO_INFO_STEPSPERREV:
		hr = VariantToDouble(*newValue, &dval);
		if (SUCCEEDED(hr))
		{
			this->m_stepsPerRev = dval;
//			this->m_fDirty		= TRUE;
		}
		break;
	case MONO_INFO_MODEL:
		hr = VariantToStringAlloc(*newValue, &szString);
		if (SUCCEEDED(hr))
		{

		}
		break;
	}
	return FALSE;
}

void CMySciUsbMono::GetGratingInfo(
							short int		gratingNumber,
							short int		Index,
							VARIANT		*	GratingInfo)
{
	GRATING_INFO	*	pGratingInfo		= this->GetGratingInfo(gratingNumber);
	CQuadratic		*	pQuadratic;
	if (NULL == pGratingInfo) return;
	switch (Index)
	{
	case 0:
		// pitch
		InitVariantFromInt32(pGratingInfo->pitch, GratingInfo);
		break;
	case 1:
		InitVariantFromString(pGratingInfo->szBlaze, GratingInfo);
		break;
	case 2:
	case 14:
		// max effective wavelength
		InitVariantFromDouble(pGratingInfo->wavelengthRange[1], GratingInfo);
		break;
	case 3:
	case 13:
		// min effective wavelength
		InitVariantFromDouble(pGratingInfo->wavelengthRange[0], GratingInfo);
		break;
	case 7:
		// grating nm step resolution
		{
			double				stepsPerNM = this->GetGratingStepsPerNM(gratingNumber);
			if (stepsPerNM != 0.0)
				InitVariantFromDouble(1.0/stepsPerNM, GratingInfo);
			else
				InitVariantFromDouble(1.0 / 50.0, GratingInfo);
		}
		break;
	case 6:						// phase error
		InitVariantFromDouble(pGratingInfo->phaseError, GratingInfo);
		break;
	case 9:						// steps per nm
		InitVariantFromDouble(this->GetGratingStepsPerNM(gratingNumber), GratingInfo);
		break;
	case GRATING_INFO_OFFSETTERM:
		pQuadratic = this->GetQuadratic(gratingNumber);
		if (NULL != pQuadratic)
			InitVariantFromDouble(pQuadratic->GetOffsetFactor(), GratingInfo);
		else
			InitVariantFromDouble(0.0, GratingInfo);
		break;
	case GRATING_INFO_LINEARTERM:
		pQuadratic = this->GetQuadratic(gratingNumber);
		if (NULL != pQuadratic)
			InitVariantFromDouble(pQuadratic->GetLinearFactor(), GratingInfo);
		else
			InitVariantFromDouble(1.0, GratingInfo);
		break;
	case GRATING_INFO_QUADTERM:
		pQuadratic = this->GetQuadratic(gratingNumber);
		if (NULL != pQuadratic)
			InitVariantFromDouble(pQuadratic->GetQuadFactor(), GratingInfo);
		else
			InitVariantFromDouble(0.0, GratingInfo);
		break;
	case 4:
		InitVariantFromInt32(pGratingInfo->zeroPosition, GratingInfo);
		break;
	case GRATING_INFO_DUALPASSOFFSET:
		InitVariantFromDouble(pGratingInfo->DualPassOffset, GratingInfo);
		break;
	default:
		break;
	}
/*
Public Enum GratingIndex
    GI_Pitch = 0                ' grating pitch
    GI_Blaze = 1
    GI_MaxEffWave = 2
    GI_MinEffWave = 3
    GI_ZeroPos = 4
    GI_Location = 5
    GI_PhaseErr = 6
    GI_NmStepRes = 7
    GI_Resolution = 8
    GI_StepsPerNM = 9           ' steps per nanometers
    GI_GratingPeriod = 10       ' grating period for direct drive
    GI_StepsPerRadian = 11      ' steps per radian
    GI_Period = 12              ' grating period for direct drive
    GI_MinWave = 13             ' physical minimum wavelength for grating
    GI_MaxWave = 14             ' physical maximum wavelength for grating
End Enum


    Select Case Index
        Case GI_Pitch
            Grating(ID).Pitch = CInt(Value)
        Case GI_Blaze
            Grating(ID).Blaze = CInt(Value)
        Case GI_MaxEffWave
            Grating(ID).MaxEffWave = CDbl(Value)
'            MsgBox "Grating Max wavelength grating = " & ID & " value = " & Value
        Case GI_MinEffWave
            Grating(ID).MinEffWave = CDbl(Value)
 '           MsgBox "Grating Min wavelength grating = " & ID & " value = " & Value
        Case GI_ZeroPos
            Grating(ID).ZeroPos = CLng(Value)
        Case GI_Location
            Grating(ID).Location = CLng(Value)
        Case GI_PhaseErr
            Grating(ID).PhaseErr = CDbl(Value)
        Case GI_NmStepRes
            Grating(ID).NmStepRes = CDbl(Value)
        Case GI_MinWave ' = 13             ' physical minimum wavelength for grating
            Grating(ID).minWave = CDbl(Value)
        Case GI_MaxWave ' = 14             ' physical maximum wavelength for grating
            Grating(ID).minWave = CDbl(Value)
    End Select
*/
}

BOOL CMySciUsbMono::SetGratingInfo(
							short int		gratingNumber,
							short int		Index,
							VARIANT		*	newValue)
{
	GRATING_INFO	*	pGratingInfo		= this->GetGratingInfo(gratingNumber);
	long				lval;
	double				dval;
	HRESULT				hr;
	if (NULL == pGratingInfo) return FALSE;
	switch (Index)
	{
	case 0:
		// pitch
//		InitVariantFromInt32(pGratingInfo->pitch, GratingInfo);
		break;
	case 2:
	case 14:
		// max effective wavelength
//		InitVariantFromDouble(pGratingInfo->wavelengthRange[1], GratingInfo);
		break;
	case 3:
	case 13:
		// min effective wavelength
//		InitVariantFromDouble(pGratingInfo->wavelengthRange[0], GratingInfo);
		break;
	case 7:
/*
		// grating nm step resolution
		{
			double				stepsPerNM = this->GetGratingStepsPerNM(gratingNumber);
			if (stepsPerNM != 0.0)
				InitVariantFromDouble(1.0/stepsPerNM, GratingInfo);
			else
				InitVariantFromDouble(1.0 / 50.0, GratingInfo);
		}
*/
		break;
	case 4:					// zero position
		hr = VariantToInt32(*newValue, &lval);
		if (SUCCEEDED(hr))
		{
			pGratingInfo->zeroPosition	= lval;
			return TRUE;
		}
		break;
	case 5:					// location
		break;
	case 6:				// phase error
		hr = VariantToDouble(*newValue, &dval);
		if (SUCCEEDED(hr))
		{
			pGratingInfo->phaseError	= dval;
			return TRUE;
		}
		break;
	case GRATING_INFO_ZEROPOSOFFSET:
		hr = VariantToDouble(*newValue, &dval);
		if (SUCCEEDED(hr))
		{
			pGratingInfo->ZeroPositionOffset = dval;
			return TRUE;
		}
		break;
	case GRATING_INFO_TEMPZEROOFFSET:
		this->DetermineTempOffset(gratingNumber);
		return TRUE;
	default:
		break;
	}
	return FALSE;
}

double CMySciUsbMono::GetmoveTime(
							double			newPosition)
{
	

	return 0.0;
}

void CMySciUsbMono::Setup()
{
	HRESULT				hr;
	IOleObject		*	pOleObject;
	HWND				hwndParent		= this->GetParentWindow();

	// check if setup window is already open
	if (NULL != this->m_hwndSetup && IsWindow(this->m_hwndSetup))
	{
		ShowWindow(this->m_hwndSetup, SW_SHOW);
		return;
	}
   	hr = this->m_pMyObject->QueryInterface(IID_IOleObject, (LPVOID*) &pOleObject);
	if (SUCCEEDED(hr))
	{
		hr = pOleObject->DoVerb(OLEIVERB_PROPERTIES, NULL, NULL, -1, hwndParent, NULL);
		pOleObject->Release();
	}
	// update the config file if dirty
//	if (this->m_fDirty)
//	{
		this->WriteConfig(this->m_szConfigFile);
//		this->m_fDirty	= FALSE;
//	}
}

BOOL CMySciUsbMono::SetGratingParams()
{
	return TRUE;
}

BOOL CMySciUsbMono::IsValidPosition(
							double			position)
{
	long			grating;
	if (this->FindGrating(position, &grating))
	{
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

BOOL CMySciUsbMono::ConvertStepsToNM(
							BOOL			fToNM,
							short int		gratingID,
							double		*	position,
							long		*	Steps)
{
	if (fToNM)
	{
		// convert steps to NM
		return this->ConvertSteps(gratingID, *Steps, position);
	}
	else
	{
		// convert NM to steps
		return this->ConvertPosition(gratingID, *position, Steps);
	}
	return FALSE;
}

BOOL CMySciUsbMono::ConvertSteps(
							long			gratingID,
							long			steps,
							double		*	position)
{
	GRATING_INFO	*	pGratingInfo		= this->GetGratingInfo(gratingID);
 	CQuadratic		*	pQuadratic			= this->GetQuadratic(gratingID);
	BOOL				fSuccess			= FALSE;

	if (0 == lstrcmpi(this->m_szDriveType, TEXT("SineDrive")))
	{
		if (NULL != pGratingInfo && pGratingInfo->stepsPerNM != 0.0)
		{
			*position = ((steps - pGratingInfo->GetZeroPos()) / pGratingInfo->stepsPerNM);
			fSuccess = TRUE;
		}
	}
	else if (0 == lstrcmpi(this->m_szDriveType, TEXT("DirectDrive")))
	{
/*
        Value = Sin(CDbl(Steps - Grating(gratingID).ZeroPos) / _
            Grating(gratingID).StepsPerRad) * Grating(gratingID).Period
*/
		double			stepsPerRadian		= this->GetStepsPerRadian();
		double			period				= this->CalculatePeriod(gratingID);
		long			gratingZeroPosition	= pGratingInfo->GetZeroPos();
		*position	= sin(((steps - gratingZeroPosition) * 1.0) / stepsPerRadian) * period;
		fSuccess	= TRUE;
	}
	if (fSuccess && NULL != pQuadratic)
	{
		*position = pQuadratic->ExpToActual(*position);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::ConvertPosition(
							long			gratingID,
							double			position,
							long		*	steps)
{
	GRATING_INFO	*	pGratingInfo		= this->GetGratingInfo(gratingID);
 	CQuadratic		*	pQuadratic			= this->GetQuadratic(gratingID);
	if (NULL == pGratingInfo) return FALSE;
	if (NULL != pQuadratic)
		position = pQuadratic->ActualToExp(position);
	if (0 == lstrcmpi(this->m_szDriveType, TEXT("SineDrive")))
	{
		*steps = pGratingInfo->GetZeroPos() + 
			(long) floor((position * pGratingInfo->stepsPerNM) + 0.5);
		return TRUE;
	}
	else if (0 == lstrcmpi(this->m_szDriveType, TEXT("DirectDrive")))
	{
/*
        'get the grating angle (radians)
        Value = ArcSin(Value / Grating(gratingID).Period, fHaveError, errorString)
        If fHaveError Then
            DoRaiseError CStr(errorString)
            NmToSteps = 0
            Exit Function
        End If
        'Check for ArcSin Error
        If Value <> TWOPI Then
            'Compute number of motor steps from end switch
            Steps = Value * Grating(gratingID).StepsPerRad + _
                Grating(gratingID).ZeroPos
        Else
            Steps = 0
        End If
*/
		double				value;
		double				period			= this->CalculatePeriod(gratingID);
		double				stepsPerRadian	= this->GetStepsPerRadian();
		value	= asin(position / period);
		*steps	= (long) floor((value * stepsPerRadian) + 0.5) + pGratingInfo->GetZeroPos();
		return TRUE;
	}
	return FALSE;
}

void CMySciUsbMono::Abort()
{
}

void CMySciUsbMono::WriteConfig(
							LPCTSTR			Config)
{
	CIniFile			iniFile;
	long				grating;
	GRATING_INFO	*	pGratingInfo;		//	=	this->GetGratingInfo(grating);
	TCHAR				szSection[MAX_PATH];		// grating section

	if (iniFile.Init(Config))
	{
		// USB mono section
		iniFile.WriteIntValue(TEXT("USB Mono"), TEXT("IdleCurrent"), this->m_idleCurrent);
		iniFile.WriteIntValue(TEXT("USB Mono"), TEXT("RunCurrent"), this->m_runCurrent);
		iniFile.WriteDoubleValue(TEXT("USB Mono"), TEXT("StepsPerRev"), this->m_stepsPerRev);
		iniFile.WriteIntValue(TEXT("USB Mono"), TEXT("HighSpeed"), this->m_highSpeed);
		// loop over the gratings
		for (grating = 1; grating <= this->m_numberOfGratings; grating++)
		{
			StringCchPrintf(szSection, MAX_PATH, TEXT("Grating %1d"), grating);
			pGratingInfo = this->GetGratingInfo(grating);
			if (NULL != pGratingInfo)
			{
				iniFile.WriteIntValue(szSection, TEXT("ZeroPos"), pGratingInfo->zeroPosition);
				iniFile.WriteDoubleValue(szSection, TEXT("PhaseErr"), pGratingInfo->phaseError);
			}
		}
	}
}

void CMySciUsbMono::WriteINI(
							LPCTSTR			INIFile)
{
}

double CMySciUsbMono::GetCounter()
{
	return this->Getposition();
}

BOOL CMySciUsbMono::setInitialPositions(
							long			currentGrating, 
							BOOL			autoGrating, 
							double			CurrentPos)
{
	return FALSE;
}

BOOL CMySciUsbMono::SaveGratingZeroPosition(
							short int		Grating, 
							long			ZeroPos)
{
	return FALSE;
}

void CMySciUsbMono::doInit()
{
	this->SetAmOpen(TRUE);
}

BOOL CMySciUsbMono::CanAutoSelect()
{
	if (this->m_fCanAutoSelect)
	{
		// can auto select if there is more than 1 grating
		return this->GetNumberOfGratings() > 1;
	}
	else
	{
		return FALSE;
	}
}

// run current
long CMySciUsbMono::GetRunCurrent()
{
	return this->m_runCurrent;
}

void CMySciUsbMono::SetRunCurrent(
							long			runCurrent)
{
	this->m_runCurrent = runCurrent;
	Utils_OnPropChanged(this->m_pMyObject, DISPID_RunCurrent);
}

// idle current
long CMySciUsbMono::GetIdleCurrent()
{
	return this->m_idleCurrent;
}

void CMySciUsbMono::SetIdleCurrent(
							long			idleCurrent)
{
	this->m_idleCurrent = idleCurrent;
	Utils_OnPropChanged(this->m_pMyObject, DISPID_IdleCurrent);
}

CMySciUsbMono::GRATING_INFO* CMySciUsbMono::GetGratingInfo(
							long			gratingIndex)
{
	if (gratingIndex >= 1 && gratingIndex <= this->m_numberOfGratings)
	{
		return &(this->m_paGratingInfo[gratingIndex - 1]);
	}
	else
		return NULL;
}

CQuadratic* CMySciUsbMono::GetQuadratic(
							long			gratingIndex)
{
	if (gratingIndex >= 1 && gratingIndex <= this->m_numberOfGratings && NULL != this->m_paQuadratic)
	{
		return this->m_paQuadratic[gratingIndex - 1];
	}
	else
	{
		return NULL;
	}
}

long CMySciUsbMono::GetGratingZeroOffset(
							long			Grating)
{
	GRATING_INFO	* pGratingInfo	=	this->GetGratingInfo(Grating);
	if (NULL != pGratingInfo)
		return pGratingInfo->GetZeroPos();
	else
		return 0;
}

void CMySciUsbMono::SetGratingZeroOffset(
							long			Grating,
							long			zeroOffset)
{
	GRATING_INFO	* pGratingInfo	=	this->GetGratingInfo(Grating);
	if (NULL != pGratingInfo)
	{
		pGratingInfo->zeroPosition = zeroOffset;
		Utils_OnPropChanged(this->m_pMyObject, DISPID_GratingZeroOffset);
		this->m_fDirty	= TRUE;
	}
}

double CMySciUsbMono::GetGratingStepsPerNM(
							long			Grating)
{
	GRATING_INFO	* pGratingInfo	=	this->GetGratingInfo(Grating);
	if (NULL != pGratingInfo)
		return pGratingInfo->stepsPerNM;
	else
		return 0.0;
}
void CMySciUsbMono::SetGratingStepsPerNM(
							long			Grating,
							double			stepsPerNM)
{
	GRATING_INFO	* pGratingInfo	=	this->GetGratingInfo(Grating);
	if (NULL != pGratingInfo)
	{
		pGratingInfo->stepsPerNM = stepsPerNM;
		Utils_OnPropChanged(this->m_pMyObject, DISPID_GratingStepsPerNM);
	}
}

LRESULT CALLBACK DlgProcDummy(HWND, UINT, WPARAM, LPARAM);

// SciArcus
HWND CMySciUsbMono::GetControlWindow()
{
	HWND				hwndControl		= this->m_pMyObject->GetControlWindow();
	if (NULL == hwndControl)
		hwndControl	= this->m_pMyObject->FireRequestMainWindow();
	if (NULL == hwndControl)
	{
		// use dummy dialog
		if (NULL == this->m_hwndDummy)
		{
			this->m_hwndDummy = CreateDialogParam(GetOurInstance(), 
				MAKEINTRESOURCE(IDD_DIALOGDUMMY),
				NULL, (DLGPROC) DlgProcDummy, 0);
		}
		hwndControl = this->m_hwndDummy;
	}
	return hwndControl;
}

LRESULT CALLBACK DlgProcDummy(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		return TRUE;
	default:
		break;
	}
	return FALSE;
}


void CMySciUsbMono::OnMoveCompleted()
{
	if (this->m_fWaitingForPositionChange)
	{
		this->m_fWaitingForPositionChange	= FALSE;
//		Utils_OnPropChanged(this->m_pMyObject, DISPID_position);
		this->SetAmBusy(FALSE);
	}
	else if (this->m_fWaitingForGratingChange)
	{
		this->m_fWaitingForGratingChange	= FALSE;
		Utils_OnPropChanged(this->m_pMyObject, DISPID_currentGrating);
		this->SetAmBusy(FALSE);
	}
}

void CMySciUsbMono::OnDeviceHomed()
{
//	Utils_OnPropChanged(this->m_pMyObject, DISPID_AmOpen);
	this->m_pMyObject->OnHaveInitialized();
	this->SetAmBusy(FALSE);
	this->m_currentGrating	= 1;
//	this->Setposition(0.0);
	if (NULL != this->m_pStartupValues)
	{
		this->SetautoGrating(this->m_pStartupValues->GetAutoGrating());
		if (this->m_currentGrating != this->m_pStartupValues->GetGrating())
		{
			this->SetcurrentGrating(this->m_pStartupValues->GetGrating());
		}
		else
		{
			this->Setposition(this->m_pStartupValues->GetPosition());
		}
	}
	else
	{
		this->Setposition(50.0);
	}
}

void CMySciUsbMono::OnDeviceConnected()
{
	// attempt to home
//	this->m_pSciArcus->SetMotorHomed();
//	this->m_currentGrating	= 1;
//	this->Setposition(0.0);
}

// set the busy flag
void CMySciUsbMono::SetAmBusy(
							BOOL			fAmBusy)
{
	this->m_pMyObject->FireBusyStatusChange(fAmBusy);
}

// backlash correction
BOOL CMySciUsbMono::GetApplyBacklashCorrection()
{
	return this->m_fApplyBacklashCorrection;
}

void CMySciUsbMono::SetApplyBacklashCorrection(
							BOOL			fApplyBacklashCorrection)
{
	this->m_fApplyBacklashCorrection = fApplyBacklashCorrection;
	Utils_OnPropChanged(this->m_pMyObject, DISPID_ApplyBacklashCorrection);
}

long CMySciUsbMono::GetBacklashSteps()
{
	return this->m_backlashSteps;
}

void CMySciUsbMono::SetBacklashSteps(
	long			backlashSteps)
{
	this->m_backlashSteps = backlashSteps;
}

// get the steps per radian for a Direct Drive
double CMySciUsbMono::GetStepsPerRadian()
{
	/*
               .StepsPerRad = CDbl( _
                    motor.GetStepsPerRev(mono.DriveMotor)) _
                    * CDbl(mono.GearTeeth) / TWOPI
	*/
	double				PI		= atan(1.0) * 4.0;
	return (this->m_stepsPerRev * this->m_gearTeeth) / (2.0 * PI);
}

// determine the period for a Direct Drive grating
double CMySciUsbMono::CalculatePeriod(
							long			grating)
{
	/*
                .Period = 2000000# * _
                    Cos((mono.InputAngle + _
                    mono.OutputAngle - _
                    .PhaseErr) * PI / 180#) _
                    / .Pitch
	*/
	GRATING_INFO*		pGratingInfo		= this->GetGratingInfo(grating);
	if (NULL == pGratingInfo) return 0.0;
	double				PI					= atan(1.0) * 4.0;
	double				pitch				= pGratingInfo->pitch;
	double				phaseErr			= pGratingInfo->phaseError;
	return 2000000.0 * cos((this->m_inputAngle + 
		this->m_outputAngle - phaseErr) * PI / 180.0) / pitch;
}

// forward declaration of callback function
BOOL CALLBACK ETWProc( HWND hwnd, LPARAM lParam );

// obtain a parent window somehow
HWND CMySciUsbMono::GetParentWindow()
{
	HWND			hwndParent		= this->m_pMyObject->FireRequestMainWindow();
	if (NULL == hwndParent)
	{
		// get a window to use
	    EnumThreadWindows(GetCurrentThreadId(), ETWProc, (LPARAM)&hwndParent);
	}
	return hwndParent;
}

BOOL CALLBACK ETWProc( HWND hwnd, LPARAM lParam ) 
{
    DWORD *pdw = (DWORD *)lParam;

    /*
     * If this window has no parent, then it is a toplevel
     * window for the thread.  Remember the last one we find since it
     * is probably the main window.
     */

    if (GetParent(hwnd) == NULL) 
	{
        *pdw = (DWORD)hwnd;
    }

    return TRUE;
}

// find grating corresponding to a given wavelength
BOOL CMySciUsbMono::FindGrating(
							double			wavelength,
							long		*	grating)
{
	long			currentGrating		= this->GetcurrentGrating();
	GRATING_INFO*	pGratingInfo;
	long			i;
	long			numGrating;
	BOOL			fSuccess			= FALSE;

//DoLogString(TEXT("In FindGrating"));

	// first check the min and max effective wavelength for the current grating
	pGratingInfo = this->GetGratingInfo(currentGrating);
	if (NULL != pGratingInfo)
	{
		if (wavelength <= pGratingInfo->wavelengthRange[1] && 
			wavelength >= pGratingInfo->wavelengthRange[0])
		{
			*grating = currentGrating;
			return TRUE;
		}
	}
//DoLogString(TEXT("Grating out of range"));
//return FALSE;
	// check the auto grating flag
	if (this->GetautoGrating())
	{
		// check the other gratings
		numGrating	= this->GetNumberOfGratings();
		i			= 1;
		while (i<= numGrating && !fSuccess)
		{
			pGratingInfo	= this->GetGratingInfo(i);
			if (NULL != pGratingInfo)
			{
				if (wavelength <= pGratingInfo->wavelengthRange[1] && 
					wavelength >= pGratingInfo->wavelengthRange[0])
				{
					*grating = i;
					fSuccess = TRUE;
				}
			}
			if (!fSuccess) i++;
		}
	}
	else
	{
		// not auto grating, allow if wavelength is less then minimum effective wavelength
		// of the current grating
		pGratingInfo = this->GetGratingInfo(currentGrating);
		if (NULL != pGratingInfo)
		{
TCHAR			szMessage[MAX_PATH];
_stprintf_s(szMessage, MAX_PATH, 
			TEXT("Desired wavelength = %6.2f grating min = %6.2f mono min = %6.2f"),
			wavelength, pGratingInfo->wavelengthRange[0], this->m_minWave);
DoLogString(szMessage);

			if (wavelength < pGratingInfo->wavelengthRange[0])
			{
				// check that the mono allows the position
//				double			minWave = this->m_minWave * this->m_defaultPitch / pGratingInfo->pitch;
				if (wavelength >= this->m_minWave)
				{
					fSuccess	= TRUE;
					*grating	= currentGrating;
				}
			}
StringCchPrintf(szMessage, MAX_PATH, TEXT("FindGrating %s"), fSuccess ? TEXT("True") : TEXT("False"));
DoLogString(szMessage);
		}
	}
	return fSuccess;
}

CMyObject* CMySciUsbMono::GetMyObject()
{
	return this->m_pMyObject;
}


CMySciUsbMono::CHostObject::CHostObject(CMySciUsbMono * pMySciUsbMono) :
	m_pMySciUsbMono(pMySciUsbMono),
	m_pTypeInfo(NULL),
	m_cRefs(0)
{
}

CMySciUsbMono::CHostObject::~CHostObject()
{
	Utils_RELEASE_INTERFACE(this->m_pTypeInfo);
}

// IUnknown methods
STDMETHODIMP CMySciUsbMono::CHostObject::QueryInterface(
									REFIID			riid,
									LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid || IID_IHostObject == riid)
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

STDMETHODIMP_(ULONG) CMySciUsbMono::CHostObject::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CMySciUsbMono::CHostObject::Release()
{
	ULONG			cRefs		= --m_cRefs;
	if (0 == cRefs) delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CMySciUsbMono::CHostObject::GetTypeInfoCount( 
									PUINT			pctinfo)
{
	*pctinfo	= 1;
	return S_OK;
}

STDMETHODIMP CMySciUsbMono::CHostObject::GetTypeInfo( 
									UINT			iTInfo,         
									LCID			lcid,                   
									ITypeInfo	**	ppTInfo)
{
	HRESULT				hr;
	ITypeLib		*	pTypeLib;

	*ppTInfo		= NULL;
	if (NULL == this->m_pTypeInfo)
	{
		hr = GetTypeLib(&pTypeLib);
		if (SUCCEEDED(hr))
		{
			hr = pTypeLib->GetTypeInfoOfGuid(IID_IHostObject, &(this->m_pTypeInfo));
			pTypeLib->Release();
		}
	}
	else hr = S_OK;
	if (SUCCEEDED(hr))
	{
		*ppTInfo	= this->m_pTypeInfo;
		this->m_pTypeInfo->AddRef();
	}
	return hr;
}

STDMETHODIMP CMySciUsbMono::CHostObject::GetIDsOfNames( 
									REFIID			riid,                  
									OLECHAR		**  rgszNames,  
									UINT			cNames,          
									LCID			lcid,                   
									DISPID		*	rgDispId)
{
	HRESULT				hr;
	ITypeInfo		*	pTypeInfo;

	hr = this->GetTypeInfo(0, lcid, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = DispGetIDsOfNames(pTypeInfo, rgszNames, cNames, rgDispId);
		pTypeInfo->Release();
	}
	return hr;
}

STDMETHODIMP CMySciUsbMono::CHostObject::Invoke( 
									DISPID			dispIdMember,      
									REFIID			riid,              
									LCID			lcid,                
									WORD			wFlags,              
									DISPPARAMS	*	pDispParams,  
									VARIANT		*	pVarResult,  
									EXCEPINFO	*	pExcepInfo,  
									PUINT			puArgErr)
{
	switch (dispIdMember)
	{
	case DISPID_Host_SendCommand:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->SendCommand(pDispParams, pVarResult);
		break;
	case DISPID_Host_doDelay:
		if (0 != (wFlags & DISPATCH_METHOD))
			return this->doDelay(pDispParams);
		break;
	default:
		break;
	}
	return DISP_E_MEMBERNOTFOUND;
}

HRESULT CMySciUsbMono::CHostObject::SendCommand(
									DISPPARAMS	*	pDispParams,
									VARIANT		*	pVarResult)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	LPTSTR				Send		= NULL;
	LPTSTR				Receive		= NULL;
	if (NULL == this->m_pMySciUsbMono->m_pSciArcus) return E_UNEXPECTED;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	SHStrDup(varg.bstrVal, &Send);
	VariantClear(&varg);
	if (this->m_pMySciUsbMono->m_pSciArcus->SendReceive(Send, &Receive))
	{
		InitVariantFromString(Receive, pVarResult);
		CoTaskMemFree((LPVOID) Receive);
	}
	CoTaskMemFree((LPVOID) Send);
	return S_OK;
}

HRESULT CMySciUsbMono::CHostObject::doDelay(
									DISPPARAMS	*	pDispParams)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	hr = DispGetParam(pDispParams, 0, VT_I4, &varg, &uArgErr);
	if (FAILED(hr)) return hr;
	this->doDelay(varg.lVal);
	return S_OK;
}

void CMySciUsbMono::CHostObject::doDelay(
									long			delayTimeInMS)
{
	double				endTime		= 1.0 * delayTimeInMS;
	MSG					msg;			// message structure
	LARGE_INTEGER		liFreq;			// performance frequency
	LARGE_INTEGER		liStart;			// start count
	LARGE_INTEGER		liCount;
	double				currentTime		= 0.0;				// current time

	QueryPerformanceFrequency(&liFreq);
	QueryPerformanceCounter(&liStart);
	while (currentTime < endTime)
	{
		while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		Sleep(10);
		// get the current time
		QueryPerformanceCounter(&liCount);			// read the count
		currentTime = (liCount.QuadPart - liStart.QuadPart) * 1000.0 / (liFreq.QuadPart);
	}
}

// setup window
HWND CMySciUsbMono::GetSetupWindow()
{
	return this->m_hwndSetup;
}

void CMySciUsbMono::SetSetupWindow(
							HWND			hwndSetup)
{
	this->m_hwndSetup = hwndSetup;
}

// rapid scanning
BOOL CMySciUsbMono::GetRapidScanRunning()
{
	if (NULL != this->m_pDlgRapidScan)
	{
		return this->m_pDlgRapidScan->GetRapidScanRunning();
	}
	else
		return FALSE;
}

void CMySciUsbMono::SetRapidScanRunning(
							BOOL			fRapidScanRunning)
{
	if (!fRapidScanRunning)
	{
		if (this->m_pDlgRapidScan)
		{
			this->m_pDlgRapidScan->StopRapidScan();
		}
	}
}

BOOL CMySciUsbMono::StartRapidScan(
							long			grating,
							double			startWavelength,
							double			endWavelength,
							double			NMPerSecond)
{
	if (NULL == this->m_pDlgRapidScan)
		this->m_pDlgRapidScan		= new CDlgRapidScan(this);
	// clear success flag
	this->m_fRapidScanSuccess	= FALSE;
	return this->m_pDlgRapidScan->StartRapidScan(grating, startWavelength, endWavelength, NMPerSecond);
}

BOOL CMySciUsbMono::_StartRapidScan(
							long			grating,
							double			startWavelength,
							double			endWavelength,
							double			NMPerSecond)
{
	if (0.0 == NMPerSecond) NMPerSecond = 10.0;
	long			startSteps;
	long			endSteps;
	double			numberOfSeconds		= (endWavelength - startWavelength) / NMPerSecond;
	long			highSpeed;			// new high speed value
	long			stepDiff;			// difference in steps

	this->ConvertPosition(grating, startWavelength, &startSteps);
	this->ConvertPosition(grating, endWavelength, &endSteps);
	// determine the high speed value
	if (endSteps > startSteps)
		stepDiff	= endSteps - startSteps;
	else
		stepDiff	= startSteps - endSteps;
	highSpeed = (long) floor((stepDiff / numberOfSeconds) + 0.5);
	this->m_pSciArcus->SetHighSpeed(highSpeed);
	this->m_pSciArcus->MoveToPosition(endSteps);
	return TRUE;
}

void CMySciUsbMono::OnRapidScanStepped()
{
	BOOL				fMotorMoving		= this->m_pSciArcus->GetMotorMoving();
	long				position;
	long				grating				= this->m_currentGrating;
	double				wavelength;

	if (this->m_pSciArcus->GetCurrentPosition(&position))
	{
		this->ConvertSteps(grating, position, &wavelength);
		this->m_pMyObject->FireRapidScanStepped(wavelength);
	}
	if (!fMotorMoving)
	{
		this->m_fRapidScanSuccess	= TRUE;
		this->SetRapidScanRunning(FALSE);
	}
}

void CMySciUsbMono::OnRapidScanClosed()
{
	Utils_DELETE_POINTER(this->m_pDlgRapidScan);
	this->m_pMyObject->FireRapidScanEnded(this->m_fRapidScanSuccess);
	// reset the high speed
	this->SetHighSpeed(this->m_highSpeed);
}

// Grating dispersion calculation
double CMySciUsbMono::GetGratingDispersion(
	long			gratingID)
{
	// get the grating pitch
	VARIANT			Value;
	long			pitch = 1200;
	this->GetGratingInfo(gratingID, 0, &Value);
	VariantToInt32(Value, &pitch);
	if (0 == pitch) pitch = 1200;
	return 4.0 * 1200 / pitch;
}

// configuration file object
BOOL CMySciUsbMono::GetConfigFile(
							IDispatch	**	ppdisp)
{
	HRESULT				hr;
	LPOLESTR			ProgID		= NULL;
	CLSID				clsid;
	*ppdisp		= NULL;
	SHStrDup(L"Sciencetech.ConfigFile.1", &ProgID);
	hr = CLSIDFromProgID(ProgID, &clsid);
	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
		hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (LPVOID*) ppdisp);
	return SUCCEEDED(hr);
}

BOOL CMySciUsbMono::LoadConfigFile(
							IDispatch	*	pdisp,
							LPCTSTR			szConfigFile)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(pdisp, TEXT("FileName"), &dispid);
	Utils_SetStringProperty(pdisp, dispid, szConfigFile);
	Utils_GetMemid(pdisp, TEXT("ReadConfigFile"), &dispid);
	hr = Utils_InvokeMethod(pdisp, dispid, NULL, 0, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

long CMySciUsbMono::GetSectionCount(
							IDispatch	*	pdispConfigFile)
{
	DISPID				dispid;
	Utils_GetMemid(pdispConfigFile, TEXT("NumberOfSections"), &dispid);
	return Utils_GetLongProperty(pdispConfigFile, dispid);
}

BOOL CMySciUsbMono::GetSection(
							IDispatch	*	pdispConfigFile,
							long			index,
							IDispatch	**	ppdispSection)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	*ppdispSection		= NULL;
	Utils_GetMemid(pdispConfigFile, TEXT("Section"), &dispid);
	InitVariantFromInt32(index, &varg);
	hr = Utils_DoInvoke(pdispConfigFile, dispid, DISPATCH_PROPERTYGET, &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult.vt && NULL != varResult.pdispVal)
		{
			*ppdispSection	= varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess	= TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetSectionName(
							IDispatch	*	pdispSection,
							LPTSTR			szSection,
							UINT			nBufferSize)
{
	HRESULT				hr;
	DISPID				dispid;
	LPTSTR				szString		= NULL;
	BOOL				fSuccess		= FALSE;
	szSection[0]	= '\0';
	Utils_GetMemid(pdispSection, TEXT("SectionName"), &dispid);
	hr = Utils_GetStringProperty(pdispSection, dispid, &szString);
	if (NULL != szString)
	{
		StringCchCopy(szSection, nBufferSize, szString);
		fSuccess	= TRUE;
		CoTaskMemFree((LPVOID) szString);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetSectionName(
							IDispatch	*	pdispConfigFile,
							long			Index,
							LPTSTR			szSection,
							UINT			nBufferSize)
{
	IDispatch	*	pdispSection;
	BOOL			fSuccess		= FALSE;
	szSection[0]	= '\0';
	if (this->GetSection(pdispConfigFile, Index, &pdispSection))
	{
		fSuccess = this->GetSectionName(pdispSection, szSection, nBufferSize);
		pdispSection->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::FindNamedSection(
							IDispatch	*	pdispConfigFile,
							LPCTSTR			szSection,
							IDispatch	**	ppdispSection)
{
	long		nSections	= this->GetSectionCount(pdispConfigFile);
	long		i			= 0;
	BOOL		fDone		= FALSE;
	IDispatch*	pdisp;
	TCHAR		szTemp[MAX_PATH];		// ith section name
	*ppdispSection		= NULL;
	while (i < nSections && !fDone)
	{
		// get the ith section
		if (this->GetSection(pdispConfigFile, i, &pdisp))
		{
			// check for the desired section
			if (this->GetSectionName(pdisp, szTemp, MAX_PATH))
			{
				if (0 == lstrcmpi(szTemp, szSection))
				{
					// have found the desired section
					fDone = TRUE;
					*ppdispSection	= pdisp;
					pdisp->AddRef();
				}
			}
			pdisp->Release();
		}
		if (!fDone) i++;
	}
	return fDone;
}

long CMySciUsbMono::GetParameterCount(
							IDispatch	*	pdispSection)
{
	DISPID				dispid;
	Utils_GetMemid(pdispSection, TEXT("NumberOfParameters"), &dispid);
	return Utils_GetLongProperty(pdispSection, dispid);
}

BOOL CMySciUsbMono::GetParameter(
							IDispatch	*	pdispSection,
							long			Index,
							IDispatch	**	ppdispParameter)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	*ppdispParameter	= NULL;
	Utils_GetMemid(pdispSection, TEXT("Parameter"), &dispid);
	InitVariantFromInt32(Index, &varg);
	hr = Utils_DoInvoke(pdispSection, dispid, DISPATCH_PROPERTYGET, &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
		if (VT_DISPATCH == varResult. vt && NULL != varResult.pdispVal)
		{
			*ppdispParameter	= varResult.pdispVal;
			varResult.pdispVal->AddRef();
			fSuccess	= TRUE;
		}
		VariantClear(&varResult);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetParameterName(
							IDispatch	*	pdispParameter,
							LPTSTR			szParameter,
							UINT			nBufferSize)
{
	HRESULT				hr;
	DISPID				dispid;
	LPTSTR				szString		= NULL;
	BOOL				fSuccess		= FALSE;
	szParameter[0]	= '\0';
	Utils_GetMemid(pdispParameter, TEXT("ParameterName"), &dispid);
	hr = Utils_GetStringProperty(pdispParameter, dispid, &szString);
	if (NULL != szString)
	{
		StringCchCopy(szParameter, nBufferSize, szString);
		fSuccess	= TRUE;
		CoTaskMemFree((LPVOID) szString);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetParameterName(
							IDispatch	*	pdispSection,
							long			index,
							LPTSTR			szParameter,
							UINT			nBufferSize)
{
	IDispatch	*	pdispParameter;
	BOOL			fSuccess	= FALSE;
	szParameter[0]	= '\0';
	if (this->GetParameter(pdispSection, index, &pdispParameter))
	{
		fSuccess = this->GetParameterName(pdispParameter, szParameter, nBufferSize);
		pdispParameter->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::FindNamedParameter(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							IDispatch	**	ppdispParameter)
{
	long		nParams			= this->GetParameterCount(pdispSection);
	long		i				= 0;
	BOOL		fDone			= FALSE;
	IDispatch*	pdisp;
	TCHAR		szTemp[MAX_PATH];
	*ppdispParameter	= NULL;
	while (i < nParams && !fDone)
	{
		if (this->GetParameter(pdispSection, i, &pdisp))
		{
			if (this->GetParameterName(pdisp, szTemp, MAX_PATH))
			{
				if (0 == lstrcmpi(szTemp, szParameter))
				{
					fDone = TRUE;
					*ppdispParameter	= pdisp;
					pdisp->AddRef();
				}
			}
			pdisp->Release();
		}
		if (!fDone) i++;
	}
	return fDone;
}

BOOL CMySciUsbMono::GetParameterValue(
							IDispatch	*	pdispParameter,
							VARIANT		*	Value)
{
	HRESULT				hr;
	DISPID				dispid;
	Utils_GetMemid(pdispParameter, TEXT("ParameterValue"), &dispid);
	hr = Utils_InvokePropertyGet(pdispParameter, dispid, NULL, 0, Value);
	return SUCCEEDED(hr);
}

BOOL CMySciUsbMono::GetLongParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							long		*	retVal)
{
	BOOL			fSuccess	= FALSE;
	IDispatch	*	pdispParameter;
	VARIANT			Value;
	*retVal		= 0;
	if (this->FindNamedParameter(pdispSection, szParameter, &pdispParameter))
	{
		if (this->GetParameterValue(pdispParameter, &Value))
		{
			fSuccess	= TRUE;
			VariantToInt32(Value, retVal);
		}
		pdispParameter->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetDoubleParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							double		*	retVal)
{
	BOOL			fSuccess	= FALSE;
	IDispatch	*	pdispParameter;
	VARIANT			Value;
	*retVal	= 0.0;
	if (this->FindNamedParameter(pdispSection, szParameter, &pdispParameter))
	{
		if (this->GetParameterValue(pdispParameter, &Value))
		{
			fSuccess	= TRUE;
			VariantToDouble(Value, retVal);
		}
		pdispParameter->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetStringParameterValue(
							IDispatch	*	pdispSection,
							LPCTSTR			szParameter,
							LPTSTR			szString,
							UINT			nBufferSize)
{
	BOOL			fSuccess	= FALSE;
	IDispatch	*	pdispParameter;
	VARIANT			Value;
	LPTSTR			szTemp		= NULL;
	szString[0]		= '\0';
	if (this->FindNamedParameter(pdispSection, szParameter, &pdispParameter))
	{
		if (this->GetParameterValue(pdispParameter, &Value))
		{
			VariantToStringAlloc(Value, &szTemp);
			if (NULL != szTemp)
			{
				StringCchCopy(szString, nBufferSize, szTemp);
				CoTaskMemFree((LPVOID) szTemp);
				fSuccess	= TRUE;
			}
		}
		pdispParameter->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSection,
							LPCTSTR			szParameter,
							VARIANT		*	Value)
{
	IDispatch		*	pdispSection;
	IDispatch		*	pdispParameter;
	BOOL				fSuccess		= FALSE;
	if (this->FindNamedSection(pdispConfig, szSection, &pdispSection))
	{
		if (this->FindNamedParameter(pdispSection, szParameter, &pdispParameter))
		{
			fSuccess = this->GetParameterValue(pdispParameter, Value);
			pdispParameter->Release();
		}
		pdispSection->Release();
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetStringParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							LPCTSTR			szDefault,
							LPTSTR			szOutput,
							UINT			nBufferSize)
{
	BOOL			fSuccess	= FALSE;
	VARIANT			Value;
	LPTSTR			szTemp		= NULL;
	StringCchCopy(szOutput, nBufferSize, szDefault);
	if (this->GetParameterValue(pdispConfig,szSectionName, szValueName, &Value))
	{
		VariantToStringAlloc(Value, &szTemp);
		if (NULL != szTemp)
		{
			StringCchCopy(szOutput, nBufferSize, szTemp);
			CoTaskMemFree((LPVOID) szTemp);
			fSuccess = TRUE;
		}
		VariantClear(&Value);
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetLongParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							long			defValue,
							long		*	pvalue)
{
	BOOL			fSuccess	= FALSE;
	VARIANT			Value;
	*pvalue		= defValue;
	if (this->GetParameterValue(pdispConfig, szSectionName, szValueName, &Value))
	{
		VariantToInt32(Value, pvalue);
		fSuccess = TRUE;
	}
	return fSuccess;
}

BOOL CMySciUsbMono::GetDoubleParameterValue(
							IDispatch	*	pdispConfig,
							LPCTSTR			szSectionName,
							LPCTSTR			szValueName,
							double			defValue,
							double		*	pvalue)
{
	BOOL			fSuccess	= FALSE;
	VARIANT			Value;
	*pvalue	= defValue;
	if (this->GetParameterValue(pdispConfig, szSectionName, szValueName, &Value))
	{
		VariantToDouble(Value, pvalue);
		fSuccess = TRUE;
	}
	return fSuccess;
}

void CMySciUsbMono::DetermineTempOffset(
							long			gratingID)
{
	GRATING_INFO	* pGratingInfo	=	this->GetGratingInfo(gratingID);
	if (NULL != pGratingInfo)
	{
		pGratingInfo->tempZeroOffset	= 0;
		long			posZero;
		long			posOffset;
		this->ConvertPosition(gratingID, 0.0, &posZero);
		this->ConvertPosition(gratingID, pGratingInfo->ZeroPositionOffset, &posOffset);
		pGratingInfo->tempZeroOffset	= posOffset - posZero;
	}
}

BOOL CMySciUsbMono::GetReInitOnScanStart()
{
	return this->m_ReInitOnScanStart;
}

void CMySciUsbMono::SetReInitOnScanStart(
	BOOL		ReInitOnScanStart)
{
	this->m_ReInitOnScanStart = ReInitOnScanStart;
}

// scan start actions
void CMySciUsbMono::ScanStart()
{
	if (this->GetReInitOnScanStart())
	{
		this->SetAmOpen(TRUE);
	}
}
