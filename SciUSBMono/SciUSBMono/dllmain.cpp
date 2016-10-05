// dllmain.cpp : Defines the entry point for the DLL application.
#include "stdafx.h"
#include <initguid.h>
#include "MyGuids.h"
#include "NamedObjects.h"

// globals
HINSTANCE			g_hInst = NULL;
ULONG				g_cObjects = 0;
CNamedObjects	*	g_pNamedObjects = NULL;

// entry point functions
BOOL WINAPI DllMain(
	HINSTANCE	hinstDLL,
	DWORD		fdwReason,
	LPVOID		lpvReserved
)
{
	switch (fdwReason)
	{
	case DLL_PROCESS_ATTACH:
		g_hInst = hinstDLL;
		{
			TCHAR				szLogFile[MAX_PATH];
			if (GetWorkingDirectory(szLogFile, MAX_PATH))
			{
				HANDLE				hFile;
				PathAppend(szLogFile, TEXT("SciUsbMono.list"));
				hFile = CreateFile(szLogFile, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS,
					FILE_ATTRIBUTE_NORMAL, NULL);
				if (INVALID_HANDLE_VALUE != hFile)
				{
					CloseHandle(hFile);
				}
			}
		}
		g_pNamedObjects = new CNamedObjects;
		break;
	case DLL_PROCESS_DETACH:
		Utils_DELETE_POINTER(g_pNamedObjects);
		break;
	default:
		break;
	}
	return TRUE;
}


void ObjectsUp()
{
	g_cObjects++;
}

void ObjectsDown()
{
	g_cObjects--;
}

ULONG GetObjectCount()
{
	return g_cObjects;
}

HINSTANCE GetOurInstance()
{
	return g_hInst;
}

HRESULT GetTypeLib(
	ITypeLib	**	ppTypeLib)
{
	HRESULT				hr;
	TCHAR				szModule[MAX_PATH];

	hr = LoadRegTypeLib(LIBID_SciUsbMono, 1, 0, LOCALE_USER_DEFAULT, ppTypeLib);
	if (FAILED(hr))
	{
		GetModuleFileName(g_hInst, szModule, MAX_PATH);
		hr = LoadTypeLibEx(szModule, REGKIND_REGISTER, ppTypeLib);
	}
	return hr;
}

CNamedObjects* GetNamedObjects()
{
	return g_pNamedObjects;
}

// yield for messages
void MyYield()
{
	MSG				msg;
	while (PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
	{
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
}


// log a string
void DoLogString(
	LPCTSTR				szLogString)
{
	TCHAR				szLine[MAX_PATH];
	size_t				slen;
	DWORD				dwNWritten;
	TCHAR				szLogFile[MAX_PATH];
	HANDLE				hFile;

	if (GetWorkingDirectory(szLogFile, MAX_PATH))
	{
		PathAppend(szLogFile, TEXT("SciUsbMono.list"));
		hFile = CreateFile(szLogFile, GENERIC_WRITE, 0, NULL, OPEN_ALWAYS,
			FILE_ATTRIBUTE_NORMAL, NULL);
		if (INVALID_HANDLE_VALUE != hFile)
		{
			// set the file pointer to the end of the file
			SetFilePointer(hFile, 0, NULL, FILE_END);
			StringCchPrintf(szLine, MAX_PATH, TEXT("%s\r\n"), szLogString);
			StringCbLength(szLine, MAX_PATH * sizeof(TCHAR), &slen);
			WriteFile(hFile, (LPCVOID)szLine, slen, &dwNWritten, NULL);
			CloseHandle(hFile);
		}
	}
}

// get working directory
BOOL GetWorkingDirectory(
	LPTSTR			szWorkingDirectory,
	UINT			nBufferSize)
{
	HRESULT				hr;
	LPOLESTR			ProgID = NULL;
	CLSID				clsid;
	IUnknown		*	punk;
	IDispatch		*	pdisp;
	DISPID				dispid;
	TCHAR				szFileName[MAX_PATH];
	long				procID = 0;
	VARIANTARG			varg;
	VARIANT				varResult;
	BOOL				fSuccess = FALSE;
	LPTSTR				szString = NULL;

	szWorkingDirectory[0] = '\0';
	GetModuleFileName(GetOurInstance(), szFileName, MAX_PATH);
	PathRemoveFileSpec(szFileName);
	StringCchCopy(szWorkingDirectory, nBufferSize, szFileName);
	return TRUE;
	//	return FALSE;
	/*
	// get the AppWorkingDirectory object, if it exists
	Utils_AnsiToUnicode(TEXT("Sci.AppWorkingDirectory.1"), &ProgID);
	hr = CLSIDFromProgID(ProgID, &clsid);
	CoTaskMemFree((LPVOID) ProgID);
	if (SUCCEEDED(hr))
	{
	hr = GetActiveObject(clsid, NULL, &punk);
	}
	if (SUCCEEDED(hr))
	{
	hr = punk->QueryInterface(IID_IDispatch, (LPVOID*) &pdisp);
	punk->Release();
	}
	if (SUCCEEDED(hr))
	{
	// get the file name for the calling executable
	GetModuleFileName(NULL, szFileName, MAX_PATH);
	PathRemoveExtension(szFileName);
	// get the process id
	Utils_GetMemid(pdisp, TEXT("GetProcessID"), &dispid);
	InitVariantFromString(PathFindFileName(szFileName), &varg);
	hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
	VariantClear(&varg);
	if (SUCCEEDED(hr)) VariantToInt32(varResult, &procID);
	// get the working folder
	Utils_GetMemid(pdisp, TEXT("GetWorkingDirectory"), &dispid);
	InitVariantFromInt32(procID, &varg);
	hr = Utils_InvokeMethod(pdisp, dispid, &varg, 1, &varResult);
	if (SUCCEEDED(hr))
	{
	VariantToStringAlloc((varResult, &szString);
	StringCchCopy(szWorkingDirectory, nBufferSize, szString);
	fSuccess	= TRUE;
	CoTaskMemFree((LPVOID) szString);
	VariantClear(&varResult);
	}
	pdisp->Release();
	}
	if (!fSuccess)
	{
	GetModuleFileName(g_hInst, szWorkingDirectory, nBufferSize);
	PathRemoveFileSpec(szWorkingDirectory);
	fSuccess = TRUE;
	}
	return fSuccess;
	*/
}