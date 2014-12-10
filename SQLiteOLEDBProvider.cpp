#pragma once

#include "stdafx.h"
#include "resource.h"
#include "SQLiteOLEDBProvider.h"
#include "SQLiteOLEDBProvider_i.c"
#include "SQLiteDataSource.hpp"
#include <initguid.h>

CComModule _Module;
CRegExp m_RegExp;

/////////////////////////////////////////////////////////////////////////////
BEGIN_OBJECT_MAP(ObjectMap)
	OBJECT_ENTRY(CLSID_SQLiteOLEDB, CSQLiteDataSource)
END_OBJECT_MAP()

/////////////////////////////////////////////////////////////////////////////
class CSQLiteOLEDBProviderApp : public CWinApp
{
public:

	// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSQLiteOLEDBProviderApp)
	public:
    virtual BOOL InitInstance();
    virtual int ExitInstance();
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CSQLiteOLEDBProviderApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

/////////////////////////////////////////////////////////////////////////////
BEGIN_MESSAGE_MAP(CSQLiteOLEDBProviderApp, CWinApp)
//{{AFX_MSG_MAP(CSQLiteOLEDBProviderApp)
	// NOTE - the ClassWizard will add and remove mapping macros here.
	//    DO NOT EDIT what you see in these blocks of generated code!
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CSQLiteOLEDBProviderApp theApp;

/////////////////////////////////////////////////////////////////////////////
BOOL CSQLiteOLEDBProviderApp::InitInstance()
{
	_Module.Init(ObjectMap, m_hInstance, &LIBID_SQLiteOLEDBProviderLib);
	return CWinApp::InitInstance();
}

/////////////////////////////////////////////////////////////////////////////
int CSQLiteOLEDBProviderApp::ExitInstance()
{
	_Module.Term();
	return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
	return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	return _Module.UnregisterServer(TRUE);
}


