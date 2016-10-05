// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently, but
// are changed infrequently
//

#pragma once

#include "targetver.h"

#define WIN32_LEAN_AND_MEAN             // Exclude rarely-used stuff from Windows headers
// Windows Header Files:
#include <windows.h>
#include <ole2.h>
#include <olectl.h>

#include <shlobj.h>
#include <shlwapi.h>
#include <stdio.h>
#include <tchar.h>
#include <locale.h>
#include <math.h>
#include <Propvarutil.h>

#include <strsafe.h>
#include "G:\Users\Greg\Documents\Visual Studio 2015\Projects\MyUtils\MyUtils\myutils.h"
#include "resource.h"

#define				MAX_CONN_PTS				3
#define				CONN_PT_CUSTOMSINK			0
#define				CONN_PT_PROPNOTIFY			1
#define				CONN_PT__clsIMono			2

#define				STREAM_NAME					TEXT("Contents")

#define				DEF_X_EXTENT				109			// dialog units
#define				DEF_Y_EXTENT				58			// dialog units

#define				FRIENDLY_NAME				TEXT("SciUsbMono")
#define				PROGID						TEXT("Sciencetech.SciUsbMono.1")
#define				VERSIONINDEPENDENTPROGID	TEXT("Sciencetech.SciUsbMono")

#define				MY_WINDOW_CLASS				TEXT("SciUsbMono")

// mono info
#define				MONO_INFO_MODEL				0
#define				MONO_INFO_SERIALNUMBER		1
#define				MONO_INFO_RESOLUTION		2
#define				MONO_INFO_DEFAULTPITCH		3
#define				MONO_INFO_MAXWAVE			4
#define				MONO_INFO_MINWAVE			5
#define				MONO_INFO_STEPRESOLUTION	6
#define				MONO_INFO_INPUTANGLE		7
#define				MONO_INFO_OUTPUTANGLE		8
#define				MONO_INFO_FOCALLENGTH		9
#define				MONO_INFO_INITMODE			10
#define				MONO_INFO_DRIVEMOTOR		11
#define				MONO_INFO_NMDECPLACES		12
#define				MONO_INFO_NUMMOTORS			13
#define				MONO_INFO_DRIVETYPE			14
#define				MONO_INFO_GEARTEETH			15
#define				MONO_INFO_SWITCHMOTOR		16
#define				MONO_INFO_IDLECURRENT		17
#define				MONO_INFO_RUNCURRENT		18
#define				MONO_INFO_HIGHSPEED			19
#define				MONO_INFO_STEPSPERREV		20
#define				MONO_INFO_MAXSTEPS			21
#define				MONO_INFO_MINSTEPS			22

#define				MONO_INFO_NUMGRATINGS		100


// grating info
#define				GRATING_INFO_OFFSETTERM		20
#define				GRATING_INFO_LINEARTERM		21
#define				GRATING_INFO_QUADTERM		22
#define				GRATING_INFO_DUALPASSOFFSET	23			// dual pass offset
#define				GRATING_INFO_ZEROPOSOFFSET	24			// grating zero position offset
#define				GRATING_INFO_TEMPZEROOFFSET	25

// private window messages
#define				WM_OnHaveConnectedDevice	0x0577

// global functions
void		ObjectsUp();
void		ObjectsDown();
ULONG		GetObjectCount();
HINSTANCE	GetOurInstance();
HRESULT		GetTypeLib(
				ITypeLib	**	ppTypeLib);
// yield for messages
void		MyYield();
// log a string
void		DoLogString(
				LPCTSTR			szLogString);
// get working directory
BOOL		GetWorkingDirectory(
				LPTSTR			szWorkingDirectory,
				UINT			nBufferSize);
// named objects
class CNamedObjects;
CNamedObjects* GetNamedObjects();
