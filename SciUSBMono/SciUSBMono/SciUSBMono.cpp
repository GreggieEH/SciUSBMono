// SciUSBMono.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include "MyFactory.h"
#include "MyGuids.h"
#include <MyCatid.h>

#define				MY_CATID			CATID_SciMono

// local functions
BOOL SetRegKeyValue(
	LPTSTR pszKey,
	LPTSTR pszSubkey,
	LPTSTR pszValue);
HRESULT MyStringFromCLSID(
	REFGUID		refGuid,
	LPTSTR		szCLSID);
void RegisterPropPage(
	REFCLSID		clsid,
	LPTSTR			szName,
	LPTSTR			szProgID);
void UnregisterPropPage(
	REFCLSID		clsid,
	LPTSTR			szName,
	LPTSTR			szProgID);
HRESULT	RegisterCatid();
HRESULT UnRegisterCatid();
BOOL CheckCatRegistered();		// check if our category is registered
BOOL RegisterCat();

STDAPI DllCanUnloadNow(void)
{
	return 0 == GetObjectCount() ? S_OK : S_FALSE;
}

STDAPI DllGetClassObject(
	REFCLSID		rclsid,
	REFIID			riid,
	LPVOID		*	ppv)
{
	CMyFactory		*	pMyFactory = NULL;
	HRESULT				hr = CLASS_E_CLASSNOTAVAILABLE;

	*ppv = NULL;
	if (CLSID_SciUsbMono == rclsid)
	{
		// main COM object
		pMyFactory = new CMyFactory(FALSE);
	}
	else if (CLSID_PropPageSciUsbMono == rclsid)
	{
		// property page creation
		pMyFactory = new CMyFactory(TRUE);
	}
	else
	{
		return hr;
	}
	if (NULL != pMyFactory)
	{
		ObjectsUp();
		hr = pMyFactory->QueryInterface(riid, ppv);
		if (FAILED(hr))
		{
			ObjectsDown();
			delete pMyFactory;
		}
	}
	else
		hr = E_OUTOFMEMORY;
	return hr;
}

STDAPI DllRegisterServer(void)
{
	HRESULT			hr = NOERROR;
	TCHAR			szID[MAX_PATH];
	TCHAR			szCLSID[MAX_PATH];
	TCHAR			szModulePath[MAX_PATH];
	WCHAR			wszModulePath[MAX_PATH];
	ITypeLib	*	pITypeLib;
	TCHAR			szString[MAX_PATH];
	TCHAR			szStr2[MAX_PATH];
	WCHAR			wszID[MAX_PATH];
	LPTSTR			szTypeLib = NULL;

	// Obtain the path to this module's executable file for later use.
	GetModuleFileName(GetOurInstance(), szModulePath, sizeof(szModulePath) / sizeof(TCHAR));

	// Create some base key strings.
	MyStringFromCLSID(CLSID_SciUsbMono, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCat(szCLSID, MAX_PATH, szID);

	// Create ProgID keys.
	SetRegKeyValue(PROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(PROGID, TEXT("CLSID"), szID);

	// Create VersionIndependentProgID keys.
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CurVer"), PROGID);
	SetRegKeyValue(VERSIONINDEPENDENTPROGID, TEXT("CLSID"), szID);
	// Create entries under CLSID.
	SetRegKeyValue(szCLSID, NULL, FRIENDLY_NAME);
	SetRegKeyValue(szCLSID, TEXT("ProgID"), PROGID);
	SetRegKeyValue(szCLSID, TEXT("VersionIndependentProgID"), VERSIONINDEPENDENTPROGID);
	SetRegKeyValue(szCLSID, TEXT("NotInsertable"), NULL);
	SetRegKeyValue(szCLSID, TEXT("InprocServer32"), szModulePath);
	StringCchPrintf(szString, MAX_PATH, TEXT("%s,%7d"), szModulePath,
		IDI_ICONSciUsbMono);
	SetRegKeyValue(szCLSID, TEXT("DefaultIcon"), szString);
	StringCchPrintf(szString, MAX_PATH, TEXT("%s,%7d"), szModulePath,
		IDB_BITMAPSciUsbMono);
	SetRegKeyValue(szCLSID, TEXT("ToolboxBitmap32"), szString);
	SetRegKeyValue(szCLSID, TEXT("Control"), NULL);
	// miscellaneous status
	SetRegKeyValue(szCLSID, TEXT("MiscStatus"), TEXT("0"));
	StringCchPrintf(szStr2, MAX_PATH, TEXT("%s\\MiscStatus"), szCLSID);
	StringCchPrintf(szString, MAX_PATH, TEXT("%10d"),
		OLEMISC_RECOMPOSEONRESIZE |
		OLEMISC_INSIDEOUT |
		OLEMISC_ACTIVATEWHENVISIBLE |
		OLEMISC_SETCLIENTSITEFIRST);
	SetRegKeyValue(szStr2, TEXT("1"), szString);
	// type library
	StringFromGUID2(LIBID_SciUsbMono, wszID, MAX_PATH);
	//	WideCharToMultiByte(CP_ACP, 0, wszID, -1, szID, MAX_PATH * sizeof(TCHAR), NULL, NULL);
	SetRegKeyValue(szCLSID, TEXT("TypeLib"), wszID);
	// version number
	SetRegKeyValue(szCLSID, TEXT("Version"), TEXT("1.0"));
	// type library
	StringFromGUID2(LIBID_SciUsbMono, wszID, MAX_PATH);
	//	Utils_UnicodeToAnsi(wszID, &szTypeLib);
	//	if (NULL != szTypeLib)
	//	{
	SetRegKeyValue(szCLSID, TEXT("TypeLib"), wszID);
	//		CoTaskMemFree((LPVOID) szTypeLib);
	//	}
	// version number
	SetRegKeyValue(szCLSID, TEXT("Version"), TEXT("1.0"));
	// register the type library
	//	MultiByteToWideChar(
	//		CP_ACP,
	//		0,
	//		szModulePath,
	//		-1,
	//		wszModulePath,
	//		MAX_PATH);
	hr = LoadTypeLibEx(szModulePath, REGKIND_REGISTER, &pITypeLib);
	if (SUCCEEDED(hr))
		pITypeLib->Release();
	// register our catid
	RegisterCatid();

	// register the property page
	RegisterPropPage(CLSID_PropPageSciUsbMono,
		TEXT("PropPageSciUsbMono"), TEXT("Sciencetech.PropPageSciUsbMono"));

	return hr;
}

STDAPI DllUnregisterServer(void)
{
	HRESULT  hr = S_OK;
	TCHAR    szID[MAX_PATH];
	TCHAR    szCLSID[MAX_PATH];
	TCHAR    szTemp[MAX_PATH];

	//Create some base key strings.
	MyStringFromCLSID(CLSID_SciUsbMono, szID);
	StringCchCopy(szCLSID, MAX_PATH, TEXT("CLSID\\"));
	StringCchCopy(szCLSID, MAX_PATH, szID);

	// un register the entries under Version independent prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CurVer"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), VERSIONINDEPENDENTPROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, VERSIONINDEPENDENTPROGID);

	// delete the entries under prog ID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), PROGID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, PROGID);

	// delete the entries under CLSID
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("ProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("VersionIndependentProgID"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("NotInsertable"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\%s"), szCLSID, TEXT("InprocServer32"));
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);

	RegDeleteKey(HKEY_CLASSES_ROOT, szCLSID);

	// unregister the type library
	hr = UnRegisterTypeLib(LIBID_SciUsbMono, 1, 0, 0x09, SYS_WIN32);

	// unregister cat
	UnRegisterCatid();

	// unregister the property page
	UnregisterPropPage(CLSID_PropPageSciUsbMono,
		TEXT("PropPageSciUsbMono"), TEXT("Sciencetech.PropPageSciUsbMono"));

	return S_OK;
}

void RegisterPropPage(
	REFCLSID		clsid,
	LPTSTR			szName,
	LPTSTR			szProgID)
{
	TCHAR		szID[MAX_PATH];
	WCHAR		wszID[MAX_PATH];
	TCHAR		szCLSID[MAX_PATH];
	TCHAR		szModulePath[MAX_PATH];

	// Create some base key strings.
	StringFromGUID2(clsid, wszID, MAX_PATH);
	//	SHUnicodeToAnsi(wszID, szID, MAX_PATH);
	StringCchPrintf(szCLSID, MAX_PATH, TEXT("CLSID\\%s"), wszID);
	SetRegKeyValue(szCLSID, NULL, szName);
	// Obtain the path to this module's executable file for later use.
	GetModuleFileName(GetOurInstance(), szModulePath, sizeof(szModulePath) / sizeof(TCHAR));
	SetRegKeyValue(szCLSID, TEXT("InprocServer32"), szModulePath);
	// Create ProgID keys.
	SetRegKeyValue(szProgID, NULL, szName);
	SetRegKeyValue(szProgID, TEXT("CLSID"), szID);
}

void UnregisterPropPage(
	REFCLSID		clsid,
	LPTSTR			szName,
	LPTSTR			szProgID)
{
	TCHAR				szTemp[MAX_PATH];
	WCHAR				wszID[MAX_PATH];
	TCHAR				szID[MAX_PATH];
	TCHAR				szCLSID[MAX_PATH];

	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\CLSID"), szProgID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szProgID);
	StringFromGUID2(clsid, wszID, MAX_PATH);
	//	SHUnicodeToAnsi(wszID, szID, MAX_PATH);
	StringCchPrintf(szCLSID, MAX_PATH, TEXT("CLSID\\%s"), wszID);
	StringCchPrintf(szTemp, MAX_PATH, TEXT("%s\\InprocServer32"), szCLSID);
	RegDeleteKey(HKEY_CLASSES_ROOT, szTemp);
	RegDeleteKey(HKEY_CLASSES_ROOT, szCLSID);
}


/*F+F++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function: SetRegKeyValue

Summary:  Internal utility function to set a Key, Subkey, and value
in the system Registry under HKEY_CLASSES_ROOT.

Args:     LPTSTR pszKey,
LPTSTR pszSubkey,
LPTSTR pszValue)

Returns:  BOOL
TRUE if success; FALSE if not.
------------------------------------------------------------------------F-F*/
BOOL SetRegKeyValue(
	LPTSTR pszKey,
	LPTSTR pszSubkey,
	LPTSTR pszValue)
{
	BOOL bOk = FALSE;
	LONG ec;
	HKEY hKey;
	TCHAR szKey[MAX_PATH];

	StringCchCopy(szKey, MAX_PATH, pszKey);

	if (NULL != pszSubkey)
	{
		StringCchCat(szKey, MAX_PATH, TEXT("\\"));
		StringCchCat(szKey, MAX_PATH, pszSubkey);
	}

	ec = RegCreateKeyEx(
		HKEY_CLASSES_ROOT,
		szKey,
		0,
		NULL,
		REG_OPTION_NON_VOLATILE,
		KEY_ALL_ACCESS,
		NULL,
		&hKey,
		NULL);

	if (ERROR_SUCCESS == ec)
	{
		if (NULL != pszValue)
		{
			ec = RegSetValueEx(
				hKey,
				NULL,
				0,
				REG_SZ,
				(BYTE *)pszValue,
				(lstrlen(pszValue) + 1) * sizeof(TCHAR));
		}
		if (ERROR_SUCCESS == ec)
			bOk = TRUE;
		RegCloseKey(hKey);
	}
	return bOk;
}

HRESULT MyStringFromCLSID(
	REFGUID		refGuid,
	LPTSTR		szCLSID)
{
	HRESULT				hr;
	LPOLESTR			wszCLSID;
	LPTSTR				szTemp;

	hr = StringFromCLSID(refGuid, &wszCLSID);
	if (SUCCEEDED(hr))
	{
		// copy to output string
		//		Utils_UnicodeToAnsi(wszCLSID, &szTemp);
		StringCchCopy(szCLSID, MAX_PATH, wszCLSID);
		//		CoTaskMemFree((LPVOID) szTemp);
		CoTaskMemFree((LPVOID)wszCLSID);
	}
	return hr;
}


HRESULT	RegisterCatid()
{
	HRESULT				hr;
	ICatRegister	*	pCatRegister = NULL;
	CATID				arrCatid[1];

	// make sure our category is registered
	if (!CheckCatRegistered())
	{
		if (!RegisterCat()) return E_FAIL;
	}
	// get ICatRegister
	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID*)&pCatRegister);
	if (SUCCEEDED(hr))
	{
		arrCatid[0] = MY_CATID;
		hr = pCatRegister->RegisterClassImplCategories(CLSID_SciUsbMono, 1, arrCatid);
		pCatRegister->Release();
	}
	return hr;
}
HRESULT UnRegisterCatid()
{
	HRESULT				hr;
	ICatRegister	*	pCatRegister = NULL;
	CATID				arrCatid[1];

	// check if category registered
	if (!CheckCatRegistered()) return S_FALSE;
	// get ICatRegister
	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID*)&pCatRegister);
	if (SUCCEEDED(hr))
	{
		arrCatid[0] = MY_CATID;
		hr = pCatRegister->UnRegisterClassImplCategories(CLSID_SciUsbMono, 1, arrCatid);
		pCatRegister->Release();
	}
	return hr;
}

BOOL CheckCatRegistered()		// check if our category is registered
{
	HRESULT				hr;
	ICatInformation	*	pCatInfo = NULL;
	IEnumCATEGORYINFO*	pEnum = NULL;
	CATEGORYINFO		catInfo;
	BOOL				fSuccess = FALSE;

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatInformation, (LPVOID*)&pCatInfo);
	if (SUCCEEDED(hr))
	{
		hr = pCatInfo->EnumCategories(GetUserDefaultLCID(), &pEnum);
		if (SUCCEEDED(hr))
		{
			fSuccess = FALSE;
			while (!fSuccess && S_OK == pEnum->Next(1, &catInfo, NULL))
			{
				// check the cat id
				if (catInfo.catid == MY_CATID)
				{
					fSuccess = TRUE;
				}
			}
			pEnum->Release();
		}
		pCatInfo->Release();
	}
	return fSuccess;
}

BOOL RegisterCat()					// register our category
{
	HRESULT				hr;
	ICatRegister	*	pCatRegister = NULL;
	CATEGORYINFO		catInfo;

	hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr, NULL, CLSCTX_INPROC_SERVER, IID_ICatRegister, (LPVOID*)&pCatRegister);
	if (SUCCEEDED(hr))
	{
		catInfo.catid = MY_CATID;
		catInfo.lcid = LOCALE_USER_DEFAULT;
		StringCchCopy(catInfo.szDescription, 128, L"Sciencetech Monochromator Objects");
		hr = pCatRegister->RegisterCategories(1, &catInfo);
		pCatRegister->Release();
	}
	return SUCCEEDED(hr);
}


