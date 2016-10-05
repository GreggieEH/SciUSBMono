#include "stdafx.h"
#include "SciArcus.h"
#include "MySciUsbMono.h"
#include "MyObject.h"

#define				MY_PROP						TEXT("SciArcus")
#define				TMR_CHECKMOVECOMPLETED		0x0103
#define				TIMER_INTERVAL				USER_TIMER_MINIMUM

CSciArcus::CSciArcus(CMySciUsbMono * pMySciUsbMono, LPCTSTR szSerialNumber) :
	m_pMySciUsbMono(pMySciUsbMono),
	m_pdispSciArcus(NULL),
	m_szSerialNumber(NULL),			// our serial number
	m_deviceIndex(-1),				// our device index
	m_fAmHomed(FALSE),				// motor homed flag
	m_fAmHoming(FALSE),				// motor homing flag
	m_fAmMoving(FALSE),				// motor position changing
	// sink handling
	m_iidSink(IID_NULL),			// sink interface id
	m_dwCookie(0),					// connection cookie
	// subclassing the main window
	m_hwndSubclass(NULL),			// subclass window
	// absolute minimum position
	m_fMinimumSet(FALSE),
	m_Minimum(0),
	// absolute maximum position
	m_fMaximumSet(FALSE),
	m_Maximum(0)

{
//	Utils_DoCopyString(&(this->m_szSerialNumber), szSerialNumber);
	SHStrDup(szSerialNumber, &(this->m_szSerialNumber));
}

CSciArcus::~CSciArcus(void)
{
DoLogString(TEXT("Closing CSciArcus"));

//	Utils_DoCopyString(&(this->m_szSerialNumber), NULL);
	CoTaskMemFree((PVOID)this->m_szSerialNumber);
	this->m_szSerialNumber = NULL;
	// make sure that the subclass is closed
//	this->RemoveSubclass();
	if (NULL != this->m_pdispSciArcus)
	{
DoLogString(TEXT("Disconnecting SciArcus"));
TCHAR			szMessage[MAX_PATH];
		Utils_ConnectToConnectionPoint(this->m_pdispSciArcus, NULL, this->m_iidSink, 
			FALSE, &(this->m_dwCookie));
//		Utils_RELEASE_INTERFACE(this->m_pdispSciArcus);
ULONG		cRefs	= this->m_pdispSciArcus->Release();
StringCchPrintf(szMessage, MAX_PATH, TEXT("SciArcus ref count = %1d"), cRefs);
DoLogString(szMessage);
		this->m_pdispSciArcus	= NULL;
	
	}
}

BOOL CSciArcus::doInit()						// initialize the SciArcus object
{
	HRESULT					hr;
	LPOLESTR				ProgID		= NULL;
	CLSID					clsid;
	BOOL					fSuccess	= FALSE;
	IUnknown			*	punk;					// IUnknown of the running process
	CImpISink			*	pSink;					// sink implementation
	IUnknown			*	punkSink;

	// get the SciArcus class ID
//	Utils_AnsiToUnicode(TEXT("Sciencetech.SciArcus.1"), &ProgID);
	hr = CLSIDFromProgID(L"Sciencetech.SciArcus.1", &clsid);
//	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
		// check if already running
		hr = GetActiveObject(clsid, NULL, &punk);
		if (SUCCEEDED(hr))
		{
			// already running, get the dispatch interface
			hr = punk->QueryInterface(IID_IDispatch, (LPVOID*) &(this->m_pdispSciArcus));
			punk->Release();
		}
		else
		{
			// create a new instance
			hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
				(LPVOID*) &(this->m_pdispSciArcus));
		}
		// create and connect the sink
		pSink = new CImpISink(this);
		if (NULL != pSink)
		{
			hr = pSink->QueryInterface(IID_IUnknown, (LPVOID*) &punkSink);
			if (SUCCEEDED(hr))
			{
				hr = Utils_ConnectToConnectionPoint(this->m_pdispSciArcus, punkSink,
					this->m_iidSink, TRUE, &(this->m_dwCookie));
				fSuccess = SUCCEEDED(hr);
				punkSink->Release();
			}
		}
	}
//	if (fSuccess) this->SetSubclass();
	return fSuccess;
}

// find the device index
BOOL CSciArcus::FindDeviceIndex()
{
	HRESULT					hr;
	DISPID					dispid;
	long					numDevices;
	long					i;				// index over the device
	VARIANTARG				varg;
	VARIANT					varResult;
	BOOL					fDone		= FALSE;
	LPTSTR					szSerialNumber;

	if (NULL == this->m_szSerialNumber) return FALSE;
	if (!this->GetDllLoaded()) return FALSE;
	// determine the number of devices
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("NumDevices"), &dispid);
	numDevices	= Utils_GetLongProperty(this->m_pdispSciArcus, dispid);
	if (0 == numDevices) return FALSE;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("SerialNumber"), &dispid);
	// loop over the available devices
	fDone	= FALSE;
	i		= 0;
	while (i < numDevices && !fDone)
	{
		InitVariantFromInt32(i, &varg);
		hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
		if (SUCCEEDED(hr))
		{
			szSerialNumber	= NULL;
			VariantToStringAlloc(varResult, &szSerialNumber);
			if (NULL != szSerialNumber)
			{
				if (0 == lstrcmpi(szSerialNumber, this->m_szSerialNumber))
				{
					fDone = TRUE;
					this->m_deviceIndex	= i;
				}
				CoTaskMemFree((LPVOID) szSerialNumber);
			}
			VariantClear(&varResult);
		}
		// increment the index
		if (!fDone) i++;
	}
	return fDone;
}

// motor enabled
BOOL CSciArcus::GetMotorEnabled()
{
	HRESULT					hr;
	DISPID					dispid;
	VARIANTARG				varg;
	VARIANT					varResult;
	BOOL					fMotorEnabled		= FALSE;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("MotorPowerEnabled"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &varg);
	hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fMotorEnabled);
	return fMotorEnabled;
}

void CSciArcus::SetMotorEnabled(
								BOOL			motorEnabled)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("MotorPowerEnabled"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[1]);
	InitVariantFromBoolean(motorEnabled, &avarg[0]);
	Utils_InvokePropertyPut(this->m_pdispSciArcus, dispid, avarg, 2);
}

void CSciArcus::SetDllDirectory(
								LPCTSTR			szDllDirectory)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("DllDirectory"), &dispid);
	Utils_SetStringProperty(this->m_pdispSciArcus, dispid, szDllDirectory);
}

void CSciArcus::SetToolsDirectory(
								LPCTSTR			szToolsDirectory)
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("ToolsDirectory"), &dispid);
	Utils_SetStringProperty(this->m_pdispSciArcus, dispid, szToolsDirectory);
}

BOOL CSciArcus::GetDllLoaded()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("DllLoaded"), &dispid);
	return Utils_GetBoolProperty(this->m_pdispSciArcus, dispid);
}

void CSciArcus::SetDllLoaded()
{
	DISPID				dispid;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("DllLoaded"), &dispid);
	Utils_SetBoolProperty(this->m_pdispSciArcus, dispid, TRUE);
}

BOOL CSciArcus::GetDeviceConnected()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fAmConnected		= FALSE;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("AmConnected"), &dispid);
	// attempt to connect our device
	InitVariantFromInt32(this->m_deviceIndex, &varg);
	hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fAmConnected);
	return fAmConnected;
}

void CSciArcus::SetDeviceConnected(
								BOOL			fDeviceConnected)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("AmConnected"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[1]);
	InitVariantFromBoolean(fDeviceConnected, &avarg[0]);
	Utils_InvokePropertyPut(this->m_pdispSciArcus, dispid, avarg, 2);
	if (fDeviceConnected && this->GetDeviceConnected())
	{
		this->OnHaveConnectedDevice(this->m_szSerialNumber);
	}
}

BOOL CSciArcus::SendReceive(
								LPCTSTR			Send,
								LPTSTR		*	Receive)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[3];
	BSTR				bstrReceive		= NULL;			// received string
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	*Receive		= NULL;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("SendReceive"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[2]);
	InitVariantFromString(Send, &avarg[1]);
	VariantInit(&avarg[0]);
	avarg[0].vt		= VT_BYREF | VT_BSTR;
	avarg[0].pbstrVal	= &bstrReceive;
	hr = Utils_InvokeMethod(this->m_pdispSciArcus, dispid, avarg, 3, &varResult);
	if (NULL != bstrReceive)
	{
//		Utils_UnicodeToAnsi(bstrReceive, Receive);
		SHStrDup(bstrReceive, Receive);
		SysFreeString(bstrReceive);
		fSuccess	= TRUE;
	}
	return fSuccess;
}

long CSciArcus::GetValue(
								long			Index)
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			avarg[2];
	VARIANT				varResult;
	long				lValue		= 0;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("Value"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[1]);
	InitVariantFromInt32(Index, &avarg[0]);
	hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, avarg, 2, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &lValue);
	return lValue;
}

void CSciArcus::SetValue(
								long			Index,
								long			NewValue)
{
	DISPID				dispid;
	VARIANTARG			avarg[3];
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("Value"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[2]);
	InitVariantFromInt32(Index, &avarg[1]);
	InitVariantFromInt32(NewValue, &avarg[0]);
	Utils_InvokePropertyPut(this->m_pdispSciArcus, dispid, avarg, 3);
}

BOOL CSciArcus::StoreSettings()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess		= FALSE;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("StoreSettings"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &varg);
	hr = Utils_InvokeMethod(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToBoolean(varResult, &fSuccess);
	return fSuccess;
}

// check if a move was completed
BOOL CSciArcus::CheckMoveCompleted(
								BOOL		*	fNegativeLimit,
								BOOL		*	fPositiveLimit)
{
	// check the motor status
	TCHAR				szSend[MAX_PATH];
	BOOL				fMoveCompleted		= FALSE;
	LPTSTR				szReceive			= NULL;
	long				lMotorStatus;

	*fNegativeLimit		= FALSE;
	*fPositiveLimit		= FALSE;
	StringCchPrintf(szSend, MAX_PATH, TEXT("MST"));
	if (this->SendReceive(szSend, &szReceive))
	{
		if (1 == _stscanf_s(szReceive, TEXT("%d"), &lMotorStatus))
		{
			if (0 == (lMotorStatus & 0x0001)	&&
				0 == (lMotorStatus & 0x0002)	&&
				0 == (lMotorStatus & 0x0004))
			{
				fMoveCompleted = TRUE;
			}
			if (0 != (lMotorStatus & 0x0008)	||
				0 != (lMotorStatus & 0x0020))
			{
				*fNegativeLimit	= TRUE;
			}
			if (0 != (lMotorStatus & 0x0010)	||
				0 != (lMotorStatus & 0x0040))
			{
				*fPositiveLimit = TRUE;
			}
		}
		CoTaskMemFree((LPVOID) szReceive);
	}
	if (!fMoveCompleted) return FALSE;
/*
	if (this->m_fAmHoming)
	{
		this->m_fAmHoming		= FALSE;
		this->m_fAmMoving		= FALSE;
		this->m_fAmHomed		= TRUE;
		this->m_pMySciUsbMono->OnDeviceHomed();
	}
	else
	{
		this->m_fAmHoming		= FALSE;
		this->m_fAmMoving		= FALSE;
		// sink notification
		this->m_pMySciUsbMono->OnMoveCompleted();
	}
	this->m_fAmHoming		= FALSE;
	this->m_fAmMoving		= FALSE;
*/
	return fMoveCompleted;
}

// motor homing
BOOL CSciArcus::MotorHomed()
{
	return this->m_fAmHomed;
}

void CSciArcus::Home9010()
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive	= NULL;
	BOOL				fDone;
	BOOL				fNegativeLimit;
	BOOL				fPositiveLimit;
	BOOL				fError;
	long				highSpeed		= this->m_pMySciUsbMono->GetHighSpeed();
	long				lMotorStatus;
	long				position		= 0;

	// make sure that the motor is enabled
	this->SetMotorEnabled(TRUE);
	// absolute move
	this->MySendCommand(TEXT("ABS"));
	this->SetHighSpeed(highSpeed);
	// check if on end switch
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
//	// home correction amount = 0
//	this->MySendCommand(TEXT("HCA=0"));
	// check that we are not on an end switch
	lMotorStatus	= this->GetMotorStatus();
	if ((0 != (lMotorStatus & 16)) || (0 != (lMotorStatus & 32)))
	{
		// disable the limit switches
		this->MySendCommand(TEXT("DL=1"));
		// move + 500 from current position
		this->GetCurrentPosition(&position);
		position -= 200;
		StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
		this->MySendCommand(szSend);
		WaitUntilMoveCompleted();			// wait until move completed
	}
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
	// go to positive limit switch
	if (this->MySendCommand(TEXT("L+")))
	{
		// wait until complete and not on positive limit switch
		fPositiveLimit = TRUE;
		while (fPositiveLimit)
			this->WaitUntilMoveCompleted(&fPositiveLimit);
		lMotorStatus = this->GetMotorStatus();
		fError = 0 != (128 & lMotorStatus);
	}
	if (fError) this->MySendCommand(TEXT("CLR"));
	// move away
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// move + 500 from current position
	this->GetCurrentPosition(&position);
	position += 500;
	StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
	this->MySendCommand(szSend);
	WaitUntilMoveCompleted();			// wait until move completed
	// reduce speed
	SetHighSpeed(highSpeed/10);
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
	// go to negative limit switch
	if (this->MySendCommand(TEXT("L+")))
	{
		// wait until complete and not on positive limit switch
		fPositiveLimit = TRUE;
		while (fPositiveLimit)
			this->WaitUntilMoveCompleted(&fPositiveLimit);
		lMotorStatus = this->GetMotorStatus();
		fError = 0 != (128 & lMotorStatus);
	}
	if (fError) this->MySendCommand(TEXT("CLR"));
	// disable the limit switch
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// move + 500 from current position
	this->GetCurrentPosition(&position);
	position -= 250;
	StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
	this->MySendCommand(szSend);
	WaitUntilMoveCompleted();			// wait until move completed
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));

	// set high speed
	this->SetHighSpeed(highSpeed);
	this->m_fAmHomed	= TRUE;
	this->m_pMySciUsbMono->OnDeviceHomed();
}


void CSciArcus::Home9030()
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive	= NULL;
	BOOL				fDone;
	BOOL				fNegativeLimit;
	BOOL				fPositiveLimit;
	BOOL				fError;
	long				highSpeed		= this->m_pMySciUsbMono->GetHighSpeed();
	long				lMotorStatus;
	long				position		= 0;

	// make sure that the motor is enabled
	this->SetMotorEnabled(TRUE);
	// absolute move
	this->MySendCommand(TEXT("ABS"));
	this->SetHighSpeed(highSpeed);
	// check if on end switch
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
	// check that we are not on an end switch
	lMotorStatus	= this->GetMotorStatus();
	if ((0 != (lMotorStatus & 16)) || (0 != (lMotorStatus & 32)))
	{
		// disable the limit switches
		this->MySendCommand(TEXT("DL=1"));
		// move + 500 from current position
		this->GetCurrentPosition(&position);
		position += 500;
		StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
		this->MySendCommand(szSend);
		WaitUntilMoveCompleted();			// wait until move completed
	}
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
	// go to negative limit switch
	if (this->MySendCommand(TEXT("L-")))
	{
		// wait until complete
		this->WaitUntilMoveCompleted();
		lMotorStatus = this->GetMotorStatus();
		fError = 0 != (128 & lMotorStatus);
	}
	if (fError) this->MySendCommand(TEXT("CLR"));
	// move away
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// move + 500 from current position
	this->GetCurrentPosition(&position);
	position += 500;
	StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
	this->MySendCommand(szSend);
	WaitUntilMoveCompleted();			// wait until move completed
	// reduce speed
	SetHighSpeed(highSpeed/10);
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("DO1=0"));
	// go to negative limit switch
	if (this->MySendCommand(TEXT("L-")))
	{
		// wait until complete
		this->WaitUntilMoveCompleted();
		lMotorStatus = this->GetMotorStatus();
		fError = 0 != (128 & lMotorStatus);
	}
	if (fError) this->MySendCommand(TEXT("CLR"));
	// disable the limit switch
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// move + 500 from current position
	this->GetCurrentPosition(&position);
	position += 250;
	StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
	this->MySendCommand(szSend);
	WaitUntilMoveCompleted();			// wait until move completed
	// turn on the limit switches
	this->MySendCommand(TEXT("DL=0"));

	// set high speed
	this->SetHighSpeed(highSpeed);
	this->m_fAmHomed	= TRUE;
	this->m_pMySciUsbMono->OnDeviceHomed();
}

// send a command, eat return
BOOL CSciArcus::MySendCommand(
								LPCTSTR			szCommand)
{
	LPTSTR				szReceive		= NULL;
	BOOL				fSuccess		= FALSE;
	if (this->SendReceive(szCommand, &szReceive))
	{
		CoTaskMemFree((LPVOID) szReceive);
		szReceive = NULL;
		fSuccess = TRUE;
	}
	return fSuccess;
}

void CSciArcus::Home9055()
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive	= NULL;
	BOOL				fDone;
	long				highSpeed		= this->m_pMySciUsbMono->GetHighSpeed();
	BOOL				fNegativeLimit;
	BOOL				fPositiveLimit;
	long				lMotorStatus;
	BOOL				fError;
	long				acceleration	= this->m_pMySciUsbMono->GetAcceleration();
	long				position		= 0;

	this->m_fAmHoming	= TRUE;
	// make sure that the motor is enabled
	this->SetMotorEnabled(TRUE);
	// absolute move
	this->MySendCommand(TEXT("ABS"));
	// set the acceleration
	StringCchPrintf(szSend, MAX_PATH, TEXT("ACC=%1d"), acceleration);
	this->MySendCommand(szSend);
	// set the speed
	this->SetHighSpeed(highSpeed);
	// turn on the limit switches
	this->MySendCommand(TEXT("DO1=0"));
	// check that we are not on an end switch
	lMotorStatus	= this->GetMotorStatus();
	if ((0 != (lMotorStatus & 16)) || (0 != (lMotorStatus & 32)))
	{
		// disable the limit switches
		this->MySendCommand(TEXT("DL=1"));
		// move + 2000 from current position
		this->GetCurrentPosition(&position);
		position += 2000;
		StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
		this->MySendCommand(szSend);
		WaitUntilMoveCompleted();			// wait until move completed
	}
	// make sure that the limit switches are enabled
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("LCA=0"));
	// go to the negative limit switch at high speed
	if (this->MySendCommand(TEXT("L-")))
	{
		// wait until complete
//		this->WaitUntilMoveCompleted();
//		lMotorStatus = this->GetMotorStatus();
//		fError = 0 != (128 & lMotorStatus);
		fDone = FALSE;
		fError = FALSE;
		while (!fDone && !fError)
		{
			lMotorStatus = this->GetMotorStatus();
			if (0 == lMotorStatus)
			{
				// resend the command if move stopped
				this->MySendCommand(L"L-");
			}
			fDone = 16 == lMotorStatus;
			fError = (0 != (128 & lMotorStatus)) || (0 != (64 & lMotorStatus));
			if (!fDone && !fError)
			{
				MyYield();
			}
		}
	}
/*
		fDone = FALSE;
		fError = FALSE;
		while (!fDone && !fError)
		{
			lMotorStatus = this->GetMotorStatus();
			fDone = 0 != (16 & lMotorStatus);
			fError = 0 != (128 & lMotorStatus);
			if (fDone) fError = 160 == lMotorStatus;
//			fDone = 160 == this->GetMotorStatus();
//			fDone = this->CheckMoveCompleted(&fNegativeLimit, &fPositiveLimit);
			if (!fDone)
			{
				// yield for messages
				MyYield();
			}
		}
	}
*/
	// clear the limit switch error
//	if (fError)
//	{
		this->MySendCommand(TEXT("CLR"));
//	}
	// go to the positive limit switch at low speed
	this->SetHighSpeed(highSpeed / 10);
	StringCchPrintf(szSend, MAX_PATH, TEXT("L+"));
	if (this->MySendCommand(szSend))
	{
		this->WaitUntilMoveCompleted();
		lMotorStatus = this->GetMotorStatus();
		fError = 32 != lMotorStatus;
	}
/*
	if (this->SendReceive(szSend, &szReceive))
	{
		// wait until complete
		fDone = FALSE;
		fError = FALSE;
		while (!fDone)
		{
//			fDone = this->CheckMoveCompleted(&fNegativeLimit, &fPositiveLimit);
			lMotorStatus = this->GetMotorStatus();
			fDone = 0 != (lMotorStatus & 32);
			if (fDone)
				fError = 32 != lMotorStatus;
			if (!fDone)
			{
				// yield for messages
				MyYield();
			}
		}
	}
*/
	// clear the limit switch
//	if (fError)
//	{
		this->MySendCommand(TEXT("CLR"));
//	}
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// turn OFF the limit switches
	this->MySendCommand(TEXT("DO1=1"));
	// set the high speed
	this->SetHighSpeed(highSpeed);
	this->m_fAmHoming	= FALSE;

	this->m_fAmHomed	= TRUE;
	this->m_pMySciUsbMono->OnDeviceHomed();
}

// home the 9040
void CSciArcus::Home9040()
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive	= NULL;
	BOOL				fDone;
	long				highSpeed		= this->m_pMySciUsbMono->GetHighSpeed();
	BOOL				fNegativeLimit;
	BOOL				fPositiveLimit;
	long				lMotorStatus;
	BOOL				fError;
	long				acceleration	= this->m_pMySciUsbMono->GetAcceleration();
	long				position		= 0;

	this->m_fAmHoming	= TRUE;
	// make sure that the motor is enabled
	this->SetMotorEnabled(TRUE);
	// absolute move
	this->MySendCommand(TEXT("ABS"));
	// set the acceleration
	StringCchPrintf(szSend, MAX_PATH, TEXT("ACC=%1d"), acceleration);
	this->MySendCommand(szSend);
	// set the speed
	this->SetHighSpeed(highSpeed);
	// turn on the limit switches
	this->MySendCommand(TEXT("DO1=0"));
	// check that we are not on an end switch
	lMotorStatus	= this->GetMotorStatus();
	if ((0 != (lMotorStatus & 16)) || (0 != (lMotorStatus & 32)))
	{
		// disable the limit switches
		this->MySendCommand(TEXT("DL=1"));
		// move + 2000 from current position
		this->GetCurrentPosition(&position);
		position += 2000;
		StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), position);
		this->MySendCommand(szSend);
		WaitUntilMoveCompleted();			// wait until move completed
	}
	// make sure that the limit switches are enabled
	this->MySendCommand(TEXT("DL=0"));
	this->MySendCommand(TEXT("LCA=0"));
	// go to the negative limit switch at high speed
	if (this->MySendCommand(TEXT("L-")))
	{
		// wait until complete
		this->WaitUntilMoveCompleted();
		lMotorStatus = this->GetMotorStatus();
		fError = 0 != (128 & lMotorStatus);
	}
	// clear the limit switch error
	if (fError) this->MySendCommand(TEXT("CLR"));
	// go to the positive limit switch at low speed
	this->SetHighSpeed(highSpeed / 10);
	StringCchPrintf(szSend, MAX_PATH, TEXT("L+"));
	if (this->MySendCommand(szSend))
	{
		this->WaitUntilMoveCompleted();
		lMotorStatus = this->GetMotorStatus();
		fError = 32 != lMotorStatus;
	}
	// clear the limit switch error
	if (fError) this->MySendCommand(TEXT("CLR"));
	// call this position 0
	this->MySendCommand(TEXT("PX=0"));
	// disable the limit switches
	this->MySendCommand(TEXT("DL=1"));
	// turn OFF the limit switches
	this->MySendCommand(TEXT("DO1=1"));
	// set the high speed
	this->SetHighSpeed(highSpeed);
	this->m_fAmHoming	= FALSE;

	this->m_fAmHomed	= TRUE;
	this->m_pMySciUsbMono->OnDeviceHomed();
}

BOOL CSciArcus::GetAmHoming()
{
	return this->m_fAmHoming;
}

// go to position
void CSciArcus::MoveToPosition(
								long			NewPosition)
{
	if (!this->MotorHomed()) return;
	TCHAR			szError[MAX_PATH];
	long			moveDiff = this->GetMoveDifference(NewPosition);

	if (!this->CheckAbsolutePosition(NewPosition, szError, MAX_PATH))
	{
		this->m_pMySciUsbMono->GetMyObject()->FireError(szError);
		return;
	}
	if (moveDiff > 260000 || moveDiff < -260000)
	{
		long			CurrentPosition;
		long			tPos;

		// move in two parts
		this->GetCurrentPosition(&CurrentPosition);
		tPos = CurrentPosition + (moveDiff / 2);
		this->MoveToPosition(tPos);
		this->MoveToPosition(NewPosition);
		return;
	}
	if (!this->GetMotorMoving() && !this->GetAmHoming())
	{
		TCHAR				szSend[MAX_PATH];
		LPTSTR				szReceive	= NULL;
		BOOL				fDone;
		BOOL				fNegativeLimit;
		BOOL				fPositiveLimit;
		BOOL				fMoveError		= FALSE;
		BOOL				fSentOK			= FALSE;

		// make sure that the motor is enabled
		this->SetMotorEnabled(TRUE);

		StringCchPrintf(szSend, MAX_PATH, TEXT("X%1d"), NewPosition);
		if (this->SendReceive(szSend, &szReceive))
		{
			// check for move error
			this->CheckMoveError(szReceive, &fMoveError);
			fSentOK = !fMoveError;
			CoTaskMemFree((LPVOID) szReceive);
		}
		if (fMoveError)
		{
			// try again after clearing move error
			this->ClearMoveError();
			if (this->SendReceive(szSend, &szReceive))
			{
				// check for move error
				this->CheckMoveError(szReceive, &fMoveError);
				fSentOK = !fMoveError;
				CoTaskMemFree((LPVOID) szReceive);
			}
		}
		if (fSentOK)
		{
			// wait until complete
			fDone = FALSE;
			while(!fDone)
			{
				fDone = this->CheckMoveCompleted(&fNegativeLimit, &fPositiveLimit);
				if (!fDone)
				{
					// yield for messages
					MyYield();
	//				Sleep(10);
				}
			}
			BOOL			fstat=TRUE;
			this->m_pMySciUsbMono->SetAmBusy(FALSE);
/*
			this->SendReceive(TEXT("WAITX"), &szReceive);
			CoTaskMemFree((LPVOID) szReceive);
*/
/*

			this->m_fAmMoving	= TRUE;
			// set the subclass
			this->SetSubclass();
			CoTaskMemFree((LPVOID) szReceive);
*/
		}
		else
		{
//			this->m_pMySciUsbMono->GetM
		}
	}
}

// move difference
long CSciArcus::GetMoveDifference(
								long			NewPosition)
{
	long			currentPosition;

	this->GetCurrentPosition(&currentPosition);
	return NewPosition - currentPosition;
}

BOOL CSciArcus::GetMotorMoving()
{
	return this->m_fAmMoving;
}

BOOL CSciArcus::GetCurrentPosition(long * position)
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive	= NULL;
	BOOL				fSuccess	= FALSE;
	long				lVal;
	TCHAR				szLogString[MAX_PATH];

	*position	= -1;
	StringCchPrintf(szSend, MAX_PATH, TEXT("PX"));
	if (this->SendReceive(szSend, &szReceive))
	{

StringCchPrintf(szLogString, MAX_PATH, TEXT("In GetCurrentPosition Received String = %s"),
				szReceive);
DoLogString(szLogString);



		if (1 == _stscanf_s(szReceive, TEXT("%d"), &lVal))
		{
			*position = lVal;
			fSuccess = TRUE;
		}
		CoTaskMemFree((LPVOID) szReceive);
	}

if (!fSuccess)
{
	DoLogString(TEXT("Failed to translate current position string"));
}

	return fSuccess;
}

// sink events
void CSciArcus::OnError(
								LPCTSTR			Error)
{
	TCHAR				szError[MAX_PATH];

	StringCchPrintf(szError, MAX_PATH, TEXT("Arcus Motor Error: %s:"), Error);
}

void CSciArcus::OnHaveConnectedDevice(
								LPCTSTR			SerialNumber)
{
	// check if this is our device
	if (0 == lstrcmpi(SerialNumber, this->m_szSerialNumber))
	{
//		if (NULL == this->m_hwndSubclass)
//			this->SetSubclass();
//		PostMessage(this->m_hwndSubclass, WM_OnHaveConnectedDevice, 0, 0);
		this->OnHaveConnectedDevice();
	}
}

void CSciArcus::OnHaveConnectedDevice()
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	VARIANT				varModel;
	LPTSTR				szModel		= NULL;

	// set the idle and run currents
	this->SetMotorIdleCurrent(this->m_pMySciUsbMono->GetIdleCurrent());
	this->SetMotorRunCurrent(this->m_pMySciUsbMono->GetRunCurrent());
	// home this motor
	this->m_pMySciUsbMono->GetMonoInfo(0, &varModel);
	VariantToStringAlloc(varModel, &szModel);
	if (NULL != szModel)
	{
		if (0 == lstrcmpi(szModel, TEXT("9030")))
		{
			this->Home9030();
		}
		else if (0 == lstrcmpi(szModel, TEXT("9055")))
		{
			this->Home9055();
		}
		else if (0 == lstrcmpi(szModel, TEXT("9040"))		||
				 0 == lstrcmpi(szModel, TEXT("9490")))
		{
			this->Home9040();
		}
		else if (0 == lstrcmpi(szModel, TEXT("9010")))
		{
			this->Home9010();
		}
		else
		{
			// attempt to home like 9055
			this->Home9055();
		}
		CoTaskMemFree((LPVOID) szModel);
		VariantClear(&varModel);
	}
	// tell the source selector when the device has homed
//	this->m_pMySciUsbMono->OnDeviceConnected();
	this->m_pMySciUsbMono->OnDeviceHomed();
}

// motor IDLE current
long CSciArcus::GetMotorIdleCurrent()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	long				idleCurrent		= 0;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("IdleCurrent"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &varg);
	hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &idleCurrent);
	return idleCurrent;
}

void CSciArcus::SetMotorIdleCurrent(
								long			idleCurrent)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("IdleCurrent"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[1]);
	InitVariantFromInt32(idleCurrent, &avarg[0]);
	Utils_InvokePropertyPut(this->m_pdispSciArcus, dispid, avarg, 2);
}

// motor Run Current
long CSciArcus::GetMotorRunCurrent()
{
	HRESULT				hr;
	DISPID				dispid;
	VARIANTARG			varg;
	VARIANT				varResult;
	long				runCurrent		= 0;
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("RunCurrent"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &varg);
	hr = Utils_InvokePropertyGet(this->m_pdispSciArcus, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &runCurrent);
	return runCurrent;
}

void CSciArcus::SetMotorRunCurrent(
								long			runCurrent)
{
	DISPID				dispid;
	VARIANTARG			avarg[2];
	Utils_GetMemid(this->m_pdispSciArcus, TEXT("RunCurrent"), &dispid);
	InitVariantFromInt32(this->m_deviceIndex, &avarg[1]);
	InitVariantFromInt32(runCurrent, &avarg[0]);
	Utils_InvokePropertyPut(this->m_pdispSciArcus, dispid, avarg, 2);
}

// motor status
long CSciArcus::GetMotorStatus()
{
	// check the motor status
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive			= NULL;
	long				lMotorStatus;
	long				motorStatus = -1;

	StringCchPrintf(szSend, MAX_PATH, TEXT("MST"));
	if (this->SendReceive(szSend, &szReceive))
	{
		if (1 == _stscanf_s(szReceive, TEXT("%d"), &lMotorStatus))
		{
			motorStatus = lMotorStatus;
		}
	}
	return motorStatus;
}


// sink implementation
CSciArcus::CImpISink::CImpISink(CSciArcus * pSciArcus) :
	m_pSciArcus(pSciArcus),
	m_cRefs(0),
	m_dispidError(DISPID_UNKNOWN),
	m_dispidHaveConnectedDevice(DISPID_UNKNOWN)
{
	HRESULT					hr;
	ITypeInfo			*	pTypeInfo;			// sink type info
	TYPEATTR			*	pTypeAttr;			// type attributes
	FUNCDESC			*	pFuncDesc;			// function description
	BSTR					bstrName;			// function name
	LPTSTR					szName;				// function name
	UINT					i;					// index over the functions

	// get the sink type info
	hr = Utils_GetSinkInterfaceID(this->m_pSciArcus->m_pdispSciArcus, &pTypeInfo);
	if (SUCCEEDED(hr))
	{
		hr = pTypeInfo->GetTypeAttr(&pTypeAttr);
		if (SUCCEEDED(hr))
		{
			// store the sink interface id
			this->m_pSciArcus->m_iidSink	= pTypeAttr->guid;
			// loop over the functions to obtain the dispatch ids
			for (i=0; i<pTypeAttr->cFuncs; i++)
			{
				// get the function name
				szName		= NULL;
				bstrName	= NULL;
				hr = pTypeInfo->GetFuncDesc(i, &pFuncDesc);
				if (SUCCEEDED(hr))
				{
					hr = pTypeInfo->GetDocumentation(pFuncDesc->memid, &bstrName, NULL, NULL,
						NULL);
					if (SUCCEEDED(hr))
					{
						//			Utils_UnicodeToAnsi(bstrName, &szName);
						SHStrDup(bstrName, &szName);
						SysFreeString(bstrName);
						if (0 == lstrcmpi(szName, TEXT("Error")))
						{
							this->m_dispidError		= pFuncDesc->memid;
						}
						else if (0 == lstrcmpi(szName, TEXT("HaveConnectedDevice")))
						{
							this->m_dispidHaveConnectedDevice	= pFuncDesc->memid;
						}
						CoTaskMemFree((LPVOID) szName);
					}
					pTypeInfo->ReleaseFuncDesc(pFuncDesc);
				}
			}
			// release the type attributes
			pTypeInfo->ReleaseTypeAttr(pTypeAttr);
		}
		pTypeInfo->Release();
	}
}

CSciArcus::CImpISink::~CImpISink()
{
}

// IUnknown methods
STDMETHODIMP CSciArcus::CImpISink::QueryInterface(
								REFIID			riid,
								LPVOID		*	ppv)
{
	if (IID_IUnknown == riid || IID_IDispatch == riid ||
		riid == this->m_pSciArcus->m_iidSink)
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

STDMETHODIMP_(ULONG) CSciArcus::CImpISink::AddRef()
{
	return ++m_cRefs;
}

STDMETHODIMP_(ULONG) CSciArcus::CImpISink::Release()
{
	ULONG			cRefs;
	cRefs = --m_cRefs;
	if (0 == cRefs)
		delete this;
	return cRefs;
}

// IDispatch methods
STDMETHODIMP CSciArcus::CImpISink::GetTypeInfoCount( 
								PUINT			pctinfo)
{
	*pctinfo	= 0;
	return S_OK;
}

STDMETHODIMP CSciArcus::CImpISink::GetTypeInfo( 
								UINT			iTInfo,         
								LCID			lcid,                   
								ITypeInfo	**	ppTInfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CSciArcus::CImpISink::GetIDsOfNames( 
								REFIID			riid,                  
								OLECHAR		**  rgszNames,  
								UINT			cNames,          
								LCID			lcid,                   
								DISPID		*	rgDispId)
{
	return E_NOTIMPL;
}

STDMETHODIMP CSciArcus::CImpISink::Invoke( 
								DISPID			dispIdMember,      
								REFIID			riid,              
								LCID			lcid,                
								WORD			wFlags,              
								DISPPARAMS	*	pDispParams,  
								VARIANT		*	pVarResult,  
								EXCEPINFO	*	pExcepInfo,  
								PUINT			puArgErr)
{
	HRESULT				hr;
	VARIANTARG			varg;
	UINT				uArgErr;
	VariantInit(&varg);
	if (dispIdMember == this->m_dispidError)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
			LPTSTR			szError		= NULL;
	//		Utils_UnicodeToAnsi(varg.bstrVal, &szError);
			this->m_pSciArcus->OnError(varg.bstrVal);
		//	CoTaskMemFree((LPVOID) szError);
			VariantClear(&varg);
		}
	}
	else if (dispIdMember == this->m_dispidHaveConnectedDevice)
	{
		hr = DispGetParam(pDispParams, 0, VT_BSTR, &varg, &uArgErr);
		if (SUCCEEDED(hr))
		{
		//	LPTSTR			szSerialNumber	= NULL;
		//	Utils_UnicodeToAnsi(varg.bstrVal, &szSerialNumber);
			this->m_pSciArcus->OnHaveConnectedDevice(varg.bstrVal);
		//	CoTaskMemFree((LPVOID) szSerialNumber);
			VariantClear(&varg);
		}
	}
	return S_OK;
}

/*
// forward declaration of callback function
BOOL CALLBACK ETWProc2( HWND hwnd, LPARAM lParam );

// set the subclass
void CSciArcus::SetSubclass()
{
	MY_SUBCLASS		*	pMySubclass;

	// remove any current subclass
	this->RemoveSubclass();
	// subclass the control window
	this->m_hwndSubclass	= this->m_pMySciUsbMono->GetControlWindow();
	if (NULL == this->m_hwndSubclass)
	{
		// get a window to use
		HWND			hwnd = NULL;
		EnumThreadWindows(GetCurrentThreadId(), ETWProc2, (LPARAM)&hwnd );
		this->m_hwndSubclass = hwnd;
	}
	// create a new subclass object
	pMySubclass		= new MY_SUBCLASS;
	pMySubclass->m_pSciArcus	= this;
	// subclass the window
	pMySubclass->m_wpOrig		= (WNDPROC) SetWindowLongPtr(
		this->m_hwndSubclass, GWLP_WNDPROC, (LONG_PTR) SubclassProcSciArcus);
	// send the property
	SetProp(this->m_hwndSubclass, MY_PROP, (HANDLE) pMySubclass);
}*/
/*

BOOL CALLBACK ETWProc2( HWND hwnd, LPARAM lParam ) 
{
    DWORD *pdw = (DWORD *)lParam;
*/
    /*
     * If this window has no parent, then it is a toplevel
     * window for the thread.  Remember the last one we find since it
     * is probably the main window.
     */
/*
    if (GetParent(hwnd) == NULL) 
	{
        *pdw = (DWORD)hwnd;
    }

    return TRUE;
}
*/

/*
// remove the subclass
void CSciArcus::RemoveSubclass()
{
	if (NULL != this->m_hwndSubclass)
	{
		MY_SUBCLASS		*	pMySubclass;
		pMySubclass	= (MY_SUBCLASS *) GetProp(this->m_hwndSubclass, MY_PROP);
		if (NULL != pMySubclass)
		{
			// kill the timer
			KillTimer(this->m_hwndSubclass, TMR_CHECKMOVECOMPLETED);
			// remove the subclass
			SetWindowLongPtr(this->m_hwndSubclass, GWLP_WNDPROC, (LONG_PTR) pMySubclass->m_wpOrig);
			RemoveProp(this->m_hwndSubclass, MY_PROP);
			delete pMySubclass;
			this->m_hwndSubclass	= NULL;
		}
	}
}
*/

/*
// subclass procedure
LRESULT CALLBACK SubclassProcSciArcus(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	MY_SUBCLASS			*	pMySubclass;
	pMySubclass	= (MY_SUBCLASS *) GetProp(hwnd, MY_PROP);
	switch (uMsg)
	{
	case WM_OnHaveConnectedDevice:
		pMySubclass->m_pSciArcus->OnHaveConnectedDevice();
		return 0;
	case WM_TIMER:
		if (TMR_CHECKMOVECOMPLETED == wParam)
		{
			BOOL			fNegativeLimit;
			BOOL			fPositiveLimit;

			// check if move completed
			if (pMySubclass->m_pSciArcus->CheckMoveCompleted(&fNegativeLimit, &fPositiveLimit))
			{
			}
			return 0;
		}
		break;
	case WM_DESTROY:
		{
			WNDPROC				wpOrig		= pMySubclass->m_wpOrig;
			pMySubclass->m_pSciArcus->RemoveSubclass();
			return CallWindowProc(wpOrig, hwnd, WM_DESTROY, 0, 0);
		}
	default:
		break;
	}
	return CallWindowProc(pMySubclass->m_wpOrig, hwnd, uMsg, wParam, lParam);
}
*/

// high speed
long CSciArcus::GetHighSpeed()
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive		= NULL;
	long				highSpeed		= 1000;
	long				lval;
	StringCchPrintf(szSend, MAX_PATH, TEXT("HSPD"));
	if (this->SendReceive(szSend, &szReceive))
	{
		if (1 == _stscanf_s(szReceive, TEXT("%d"), &lval))
		{
			highSpeed = lval;
		}
		CoTaskMemFree((LPVOID) szReceive);
	}
	return highSpeed;
}

void CSciArcus::SetHighSpeed(
								long			highSpeed)
{
	TCHAR				szSend[MAX_PATH];
	LPTSTR				szReceive		= NULL;
	StringCchPrintf(szSend, MAX_PATH, TEXT("HSPD=%1d"), highSpeed);
	if (this->SendReceive(szSend, &szReceive))
	{
		CoTaskMemFree((LPVOID) szReceive);
	}
}

// check for move error
void CSciArcus::CheckMoveError(
								LPCTSTR			szReply,
								BOOL		*	fMoveError)
{
	*fMoveError = 0 == lstrcmpi(szReply, TEXT("?Moving"));
}

void CSciArcus::ClearMoveError()
{
	LPTSTR			szReceive		= NULL;
	if (this->SendReceive(TEXT("CLR"), &szReceive))
	{
		CoTaskMemFree((LPVOID) szReceive);
	}
}

// set minimum position
void CSciArcus::SetMinimumPosition(
								long			minimumPosition)
{
	this->m_fMinimumSet = TRUE;
	this->m_Minimum	= minimumPosition;
}

void CSciArcus::SetMaximumPosition(
								long			maximumPosition)
{
	this->m_fMaximumSet = TRUE;
	this->m_Maximum		= maximumPosition;
}

// check absolute position
BOOL CSciArcus::CheckAbsolutePosition(
								long			position,
								LPTSTR			szError,
								int				nBufferSize)
{
	if (this->m_fMinimumSet && position < this->m_Minimum)
	{
		StringCchPrintf(szError, nBufferSize,
			TEXT("Position %1d is less then minimum %1d"), position, this->m_Minimum);
		return FALSE;
	}
	if (this->m_fMaximumSet && position > this->m_Maximum)
	{
		StringCchPrintf(szError, nBufferSize,
			TEXT("Position %1d is greater then maximum %1d"),
			position, this->m_Maximum);
		return FALSE;
	}
	return TRUE;
}

// wait until move completed
BOOL CSciArcus::WaitUntilMoveCompleted(
						BOOL	*	fPositiveLimit /*= NULL*/)
{
	BOOL				fDone;
	DWORD				lMotorStatus;
	BOOL				fError;
	DWORD				dwTest = 0x01 + 0x02 + 0x04;
	// wait until complete
	fDone = FALSE;
	fError = FALSE;
	while (!fDone)
	{
		lMotorStatus = (DWORD) this->GetMotorStatus();
		fDone = 0 == (lMotorStatus & dwTest);
		if (!fDone)
		{
			MyYield();
		}
	}	
	if (NULL != fPositiveLimit)
		*fPositiveLimit = 0 != (lMotorStatus & 0x20);
	return TRUE;
}

// dialog procedure
LRESULT CALLBACK DlgProcSciArcus(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	return TRUE;
}
