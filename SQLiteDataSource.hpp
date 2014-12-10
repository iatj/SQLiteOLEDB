#pragma once

#include "SQLiteOLEDB.h"
#include "SQLiteColumnsRowset.hpp"
#include "SQLiteCommand.hpp"
#include "SQLiteDataSource.hpp"
#include "SQLiteRowset.hpp"
#include "SQLiteSession.hpp"

// ==================================================================================================================================
//	   ___________ ____    __    _ __       ____        __        _____                          
//	  / ____/ ___// __ \  / /   (_) /____  / __ \____ _/ /_____ _/ ___/____  __  _______________ 
//	 / /    \__ \/ / / / / /   / / __/ _ \/ / / / __ `/ __/ __ `/\__ \/ __ \/ / / / ___/ ___/ _ \
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/ /_/ / /_/ / /_/ /_/ /___/ / /_/ / /_/ / /  / /__/  __/
//	\____//____/\___\_\/_____/_/\__/\___/_____/\__,_/\__/\__,_//____/\____/\__,_/_/   \___/\___/ 
//	                                                                                             
// ==================================================================================================================================

/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE CSQLiteDataSource : 
	public CComObjectRootEx<THREAD_MODEL>,
	public CComCoClass<CSQLiteDataSource, &CLSID_SQLiteOLEDB>,
	public IDBCreateSessionImpl<CSQLiteDataSource, CSQLiteSession>,
	public IDBInitializeImpl<CSQLiteDataSource>,
	public IDBPropertiesImpl<CSQLiteDataSource>,
	public IPersistImpl<CSQLiteDataSource>,
	public IInternalConnectionImpl<CSQLiteDataSource>
{
public:

	CComBSTR SQLiteDataBaseFile;

	HRESULT FinalConstruct()
	{
		return FInit();
	}

	/////////////////////////////////////////////////////////////////////////////
	DECLARE_REGISTRY_RESOURCEID(IDR_SQLiteOLEDB)
	
	/////////////////////////////////////////////////////////////////////////////

	BEGIN_PROPSET_MAP(CSQLiteDataSource)

		BEGIN_PROPERTY_SET(DBPROPSET_DATASOURCEINFO)
			PROPERTY_INFO_ENTRY(ACTIVESESSIONS)
			PROPERTY_INFO_ENTRY(DATASOURCEREADONLY)			
			PROPERTY_INFO_ENTRY(OUTPUTPARAMETERAVAILABILITY)
			PROPERTY_INFO_ENTRY(PROVIDEROLEDBVER)
			PROPERTY_INFO_ENTRY(DSOTHREADMODEL)
			PROPERTY_INFO_ENTRY(SUPPORTEDTXNISOLEVELS)
			PROPERTY_INFO_ENTRY(USERNAME)
			PROPERTY_INFO_ENTRY_VALUE(SCHEMATERM, OLESTR(""))
			PROPERTY_INFO_ENTRY_VALUE(SCHEMAUSAGE, 0)
			PROPERTY_INFO_ENTRY_VALUE(DSOTHREADMODEL, DBPROPVAL_RT_APTMTTHREAD)			
			PROPERTY_INFO_ENTRY_VALUE(PROVIDEROLEDBVER, OLESTR("02.80"))
		END_PROPERTY_SET(DBPROPSET_DATASOURCEINFO)

		BEGIN_PROPERTY_SET(DBPROPSET_DBINIT)
			PROPERTY_INFO_ENTRY(INIT_DATASOURCE)
			PROPERTY_INFO_ENTRY(INIT_HWND)
			PROPERTY_INFO_ENTRY(INIT_LCID)
			PROPERTY_INFO_ENTRY_VALUE(INIT_MODE,DB_MODE_READWRITE)
			PROPERTY_INFO_ENTRY(INIT_PROMPT)
			PROPERTY_INFO_ENTRY(INIT_PROVIDERSTRING)
			PROPERTY_INFO_ENTRY(INIT_TIMEOUT)
			PROPERTY_INFO_ENTRY_VALUE(INIT_LOCATION, OLESTR("LOCALHOST"))
		END_PROPERTY_SET(DBPROPSET_DBINIT)

		CHAIN_PROPERTY_SET(CSQLiteRowset)

	END_PROPSET_MAP()

	/////////////////////////////////////////////////////////////////////////////
	BEGIN_COM_MAP(CSQLiteDataSource)
		COM_INTERFACE_ENTRY(IDBCreateSession)
		COM_INTERFACE_ENTRY(IDBInitialize)
		COM_INTERFACE_ENTRY(IDBProperties)
		COM_INTERFACE_ENTRY(IPersist)
		COM_INTERFACE_ENTRY(IInternalConnection)
	END_COM_MAP()

	/////////////////////////////////////////////////////////////////////////////
	STDMETHODIMP Initialize(void)
	{
		HRESULT hr;

		if (FAILED(hr = IDBInitializeImpl<CSQLiteDataSource>::Initialize()))
			return hr;

		// Get the database property from the OLE DB properties
		DBPROPIDSET propIDSet;
		DBPROPID    propID = DBPROP_INIT_DATASOURCE;

		propIDSet.rgPropertyIDs   = &propID;
		propIDSet.cPropertyIDs    = 1;
		propIDSet.guidPropertySet = DBPROPSET_DBINIT;

		ULONG nProps;
		DBPROPSET* propSet = 0;

		if (FAILED(hr = GetProperties(1, &propIDSet, &nProps, &propSet)))
			return E_FAIL;

		// SQLite DataBase Filename
		SQLiteDataBaseFile = CComBSTR(propSet->rgProperties[0].vValue.bstrVal);

		::VariantClear(&propSet->rgProperties[0].vValue);
		::CoTaskMemFree(propSet->rgProperties);
		::CoTaskMemFree(propSet);

		return hr;
	}

	/////////////////////////////////////////////////////////////////////////////
	STDMETHODIMP Uninitialize(void)
	{
		return IDBInitializeImpl<CSQLiteDataSource>::Uninitialize();
	}

	/////////////////////////////////////////////////////////////////////////////
	STDMETHODIMP CreateSession(IUnknown *pUnkOuter, REFIID riid, IUnknown **ppDBSession)
	{
		HRESULT hr = IDBCreateSessionImpl<CSQLiteDataSource, CSQLiteSession>::CreateSession(pUnkOuter, riid, ppDBSession);
		if (FAILED(hr)) return hr;
		
		// Get DataSource object
		CComPtr<IGetDataSource> ipGDS;
		(*ppDBSession)->QueryInterface(IID_IGetDataSource, (void **)&ipGDS );
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);

		// Open the SQLite Database
		std::string fileName = BSTR_to_STR(SQLiteDataBaseFile);
		hr = pSess->OpenSQLiteDataBase(fileName.c_str());		
		return hr;
	}
};


