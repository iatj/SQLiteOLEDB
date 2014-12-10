#pragma once

#include "SQLiteOLEDB.h"
#include "md5.h"

class CSQLiteCommand;
class CSQLiteRowset;
class CSRS_DBSCHEMA_TABLES;
class CSRS_DBSCHEMA_COLUMNS;
class CSRS_DBSCHEMA_FOREIGN_KEYS;

// ==================================================================================================================================
//	   ___________ ____    __    _ __      _____                _           
//	  / ____/ ___// __ \  / /   (_) /____ / ___/___  __________(_)___  ____ 
//	 / /    \__ \/ / / / / /   / / __/ _ \\__ \/ _ \/ ___/ ___/ / __ \/ __ \
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/__/ /  __(__  |__  ) / /_/ / / / /
//	\____//____/\___\_\/_____/_/\__/\___/____/\___/____/____/_/\____/_/ /_/ 
//	                                                                        
// ==================================================================================================================================

/////////////////////////////////////////////////////////////////////////////
class ATL_NO_VTABLE CSQLiteSession : 
	public CComObjectRootEx<THREAD_MODEL>,
	public IGetDataSourceImpl<CSQLiteSession>,
	public IOpenRowsetImpl<CSQLiteSession>,
	public ISessionPropertiesImpl<CSQLiteSession>,
	public IObjectWithSiteSessionImpl<CSQLiteSession>,
	public IDBSchemaRowsetImpl<CSQLiteSession>,
	public IDBCreateCommandImpl<CSQLiteSession, CSQLiteCommand>
	//public ITransactionLocal,
{
public:

	sqlite3 *db;

	SESSION_PROPSET_MAP(CSQLiteSession)

	/////////////////////////////////////////////////////////////////////////////
	BEGIN_COM_MAP(CSQLiteSession)
		COM_INTERFACE_ENTRY(IGetDataSource)
		COM_INTERFACE_ENTRY(IOpenRowset)
		COM_INTERFACE_ENTRY(ISessionProperties)
		COM_INTERFACE_ENTRY(IObjectWithSite)
		COM_INTERFACE_ENTRY(IDBCreateCommand)
		COM_INTERFACE_ENTRY(IDBSchemaRowset)
	END_COM_MAP()

	/////////////////////////////////////////////////////////////////////////////
	BEGIN_SCHEMA_MAP(CSQLiteSession)
		SCHEMA_ENTRY(DBSCHEMA_TABLES, CSRS_DBSCHEMA_TABLES)
		SCHEMA_ENTRY(DBSCHEMA_COLUMNS, CSRS_DBSCHEMA_COLUMNS)
		SCHEMA_ENTRY(DBSCHEMA_FOREIGN_KEYS, CSRS_DBSCHEMA_FOREIGN_KEYS)		
	END_SCHEMA_MAP()

	/////////////////////////////////////////////////////////////////////////////
	CSQLiteSession() : db(0)
	{
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT OpenSQLiteDataBase(const char* fileName)
	{		
		int rc = sqlite3_open_v2(fileName, &db, SQLITE_OPEN_READWRITE|SQLITE_OPEN_FULLMUTEX, NULL);			
		if(rc==SQLITE_OK)
		{
			rc = sqlite3_create_function(db, "BASE64_DECODE", -1, SQLITE_UTF8, NULL, &sqlite3_unbase64_fn, NULL, NULL);
		}
		return (rc==0 ? S_OK : E_FAIL);
	}

	/////////////////////////////////////////////////////////////////////////////
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return FInit();
	}

	/////////////////////////////////////////////////////////////////////////////
	void FinalRelease() 
	{		
		if(db) sqlite3_close(db);
	}

	/////////////////////////////////////////////////////////////////////////////
	STDMETHODIMP OpenRowset(IUnknown *pUnk, DBID *pTID, DBID *pInID, REFIID riid,	ULONG cSets, DBPROPSET rgSets[], IUnknown **ppRowset)
	{
		CSQLiteRowset* pRowset;
		return CreateRowset(pUnk, pTID, pInID, riid, cSets, rgSets, ppRowset, pRowset);
	}
};

// ==================================================================================================================================
//	   _____      __                            ____                           __          __      
//	  / ___/_____/ /_  ___  ____ ___  ____ _   / __ \___  _________  _________/ /_______  / /______
//	  \__ \/ ___/ __ \/ _ \/ __ `__ \/ __ `/  / /_/ / _ \/ ___/ __ \/ ___/ __  / ___/ _ \/ __/ ___/
//	 ___/ / /__/ / / /  __/ / / / / / /_/ /  / _, _/  __/ /__/ /_/ / /  / /_/ (__  )  __/ /_(__  ) 
//	/____/\___/_/ /_/\___/_/ /_/ /_/\__,_/  /_/ |_|\___/\___/\____/_/   \__,_/____/\___/\__/____/  
//	                                                                                               
// ==================================================================================================================================

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The TABLES rowset contains the TABLES (including views) of the database
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
class CSRS_DBSCHEMA_TABLES : 	
	public CRowsetImpl<CSRS_DBSCHEMA_TABLES, CTABLESRow, CSQLiteSession, CAtlArray<CTABLESRow>, CSimpleRow, IRowsetLocateImpl<CSRS_DBSCHEMA_TABLES, IRowsetLocate>>
{
public:

	ROWSET_PROPSET_MAP(CSRS_DBSCHEMA_TABLES)

	BEGIN_COM_MAP(CSRS_DBSCHEMA_TABLES)
		COM_INTERFACE_ENTRY(IRowsetLocate)
		COM_INTERFACE_ENTRY_CHAIN(_RowsetBaseClass)
	END_COM_MAP()

	DBSTATUS GetDBStatus(CSimpleRow*, ATLCOLUMNINFO* pInfo)
	{
		return DBSTATUS_S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT Execute(LONG* pcRowsAffected, ULONG, const VARIANT* params)
	{
		HRESULT hr;
		CComPtr<IGetDataSource> ipGDS;
		if (FAILED(hr = GetSpecification(IID_IGetDataSource, (IUnknown **)&ipGDS))) return hr;		
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);

		// Get list of Tables from SQLite Database
		char *zErrMsg = 0;	
		std::string sql("SELECT name as [TABLE] FROM sqlite_master WHERE type='table';");
		int rc = sqlite3_exec(pSess->db, sql.c_str(), sqlite_callback, this, &zErrMsg);		

		// Affected Records
		ULONG record_count = m_rgRowData.GetCount();
		*pcRowsAffected = record_count;

		// Bookmarks
		m_rgBookmarks.SetCount(record_count+3);
		m_rgBookmarks[0] = m_rgBookmarks[1] = m_rgBookmarks[2] = -1;
		for(ULONG i=3;i<record_count+3;i++) { m_rgBookmarks[i] = i-2; }

		return S_OK;
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	static int sqlite_callback(void *NotUsed, int argc, char **argv, char **azColName)	
	{
		CSRS_DBSCHEMA_TABLES* objPtr = (CSRS_DBSCHEMA_TABLES*) NotUsed;
		CTABLESRow trData;
		std::string TABLE(argv[0]);		
		TABLE = "[" + TABLE + "]";
		BSTR name = UTF8_to_BSTR(TABLE.c_str());		

		trData.m_guid = MD5Guid(std::string(argv[0]));

		//::wcscpy(trData.m_szType, OLESTR("TABLE"));	//ALIAS,SYNONYM,SYSTEM TABLE,VIEW,GLOBAL TEMPORARY,LOCAL TEMPORARY,SYSTEM VIEW
		//::wcscpy(trData.m_szTable, name);
		wsprintf(trData.m_szType, L"%s", L"TABLE"); //ALIAS,SYNONYM,SYSTEM TABLE,VIEW,GLOBAL TEMPORARY,LOCAL TEMPORARY,SYSTEM VIEW		
		wsprintf(trData.m_szTable, L"%s", name);		

		objPtr->m_rgRowData.Add(trData);		
		SysFreeString(name);
		return 0;
	}
};

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The COLUMNS rowset contains the columns of TABLES (including views) of the database
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
class CSRS_DBSCHEMA_COLUMNS : 
	public CRowsetImpl<CSRS_DBSCHEMA_COLUMNS, CCOLUMNSRow, CSQLiteSession, CAtlArray<CCOLUMNSRow>, CSimpleRow, IRowsetLocateImpl<CSRS_DBSCHEMA_COLUMNS, IRowsetLocate>>
{
public:

	ROWSET_PROPSET_MAP(CSRS_DBSCHEMA_COLUMNS)
	
	ULONG record_count;
	LONG ordinal_pos;
	std::string table;	
	sqlite3* db;

	/////////////////////////////////////////////////////////////////////////////
	DBSTATUS GetDBStatus(CSimpleRow*, ATLCOLUMNINFO* pInfo)
	{
		return DBSTATUS_S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT Execute(LONG* pcRowsAffected, ULONG, const VARIANT*)
	{
		HRESULT hr;
		CComPtr<IGetDataSource> ipGDS;
		if (FAILED(hr = GetSpecification(IID_IGetDataSource, (IUnknown **)&ipGDS))) return hr;		
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);

		// Step1: Create an SQL statement to get columns for all TABLES
		std::vector<std::string> TABLES = SQLiteTables(pSess->db);

		// Step2: Execute the SQL statement
		record_count=0;		
		db = pSess->db;
		for(size_t i=0;i<TABLES.size();i++)
		{
			ordinal_pos=0;
			table = TABLES[i];
			std::string sql = std::string("PRAGMA table_info(") + table + ")";
			char *zErrMsg = 0;
			int rc = sqlite3_exec(pSess->db, sql.c_str(), sqlite_callback_columns, this, &zErrMsg);
			if(rc) return E_FAIL;
		}
		*pcRowsAffected = record_count;		

		// Bookmarks
		m_rgBookmarks.SetCount(record_count+3);
		m_rgBookmarks[0] = m_rgBookmarks[1] = m_rgBookmarks[2] = -1;
		for(ULONG i=3;i<record_count+3;i++) { m_rgBookmarks[i] = i-2; }

		return S_OK;
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	static int sqlite_callback_columns(void *NotUsed, int argc, char **argv, char **azColName)	
	{		
		CSRS_DBSCHEMA_COLUMNS* objPtr = (CSRS_DBSCHEMA_COLUMNS*) NotUsed;
		CCOLUMNSRow trData;						

		// This is the information we need to fill in for DBSCHEMA_COLUMNS.
		// 0:cid, 1:name, 2:type, 3:notnull, 4:dflt_value, 5:pk

		CSQLiteColumnsRowsetRow TR;
		ATLCOLUMNINFO col;
		DBBYTEOFFSET offset=0;

		std::string TABLE(objPtr->table);
		std::string COLUMN(argv[1]);
		TR.InitColumnMetadata(objPtr->db, TABLE, COLUMN, ++(objPtr->ordinal_pos), col, offset);

		trData.m_guidType				= GUID_NULL;
		trData.m_ulOrdinalPosition		= TR.m_DBCOLUMN_NUMBER;
		trData.m_guidColumn				= TR.m_DBCOLUMN_GUID;		
		trData.m_nDataType				= TR.m_DBCOLUMN_TYPE;
		trData.m_ulColumnFlags			= TR.m_DBCOLUMN_FLAGS;		
		trData.m_nNumericPrecision		= TR.m_DBCOLUMN_PRECISION;
		trData.m_ulDateTimePrecision	= TR.m_DBCOLUMN_DATETIMEPRECISION;
		trData.m_nNumericScale			= TR.m_DBCOLUMN_SCALE;
		trData.m_ulCharOctetLength		= NULL;
		trData.m_bIsNullable			= TR.m_DBCOLUMN_SQLITE_CAN_BE_NULL;
		trData.m_bColumnHasDefault		= TR.m_DBCOLUMN_HASDEFAULT;

		::wcscpy(trData.m_szTableName, TR.m_DBCOLUMN_BASETABLENAME);
		::wcscpy(trData.m_szColumnName, TR.m_DBCOLUMN_BASECOLUMNNAME);
		
		if(argv[4])
		{
			BSTR defVal = UTF8_to_BSTR(argv[4]);		
			::wcscpy(trData.m_szColumnDefault, defVal);
			SysFreeString(defVal);
		}
		
		objPtr->record_count++;
		objPtr->m_rgRowData.Add(trData);		
		return 0;
	}

};

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// The FOREIGN_KEYS rowset contains the foreign keys of TABLES of the database
// Note: We need to define CForeignKeysRow class because it is not defined in atldb.h like CTABLESRow and CCOLUMNSRow are.
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

class CForeignKeysRow
{
public:

	WCHAR PKTableCatalog[129];
	WCHAR PKTableSchema[129];
	WCHAR PKTableName[129];
	WCHAR PKColumnName[129];
	GUID  PKColumnGUID;
	ULONG PKColumnPropID;
	WCHAR FKTableCatalog[129];
	WCHAR FKTableSchema[129];
	WCHAR FKTableName[129];
	WCHAR FKColumnName[129];
	GUID  FKColumnGUID;
	ULONG FKColumnPropID;
	ULONG ColumnOrdinal;
	WCHAR UpdateRule[129];
	WCHAR DeleteRule[129];
	WCHAR foreignKeyName[129];
	WCHAR primaryKeyName[129];
	ULONG KeySeq;

	CForeignKeysRow()
	{
		ClearMembers();
	}

	void ClearMembers()
	{
		PKTableCatalog[0]			= NULL;
		PKTableSchema[0]			= NULL;
		PKTableName[0]				= NULL;
		PKColumnName[0]				= NULL;

		FKTableCatalog[0]			= NULL;
		FKTableSchema[0]			= NULL;
		FKTableName[0]				= NULL;
		FKColumnName[0]				= NULL;

		UpdateRule[0]				= NULL;
		DeleteRule[0]				= NULL;

		foreignKeyName[0]			= NULL;
		primaryKeyName[0]			= NULL;

		FKColumnGUID = GUID_NULL;
		PKColumnGUID = GUID_NULL;

		PKColumnPropID = 0;
		FKColumnPropID = 0;
		ColumnOrdinal = 0;
		KeySeq=0;
	}

	~CForeignKeysRow()
	{
	}

	BEGIN_PROVIDER_COLUMN_MAP(CForeignKeysRow)
		PROVIDER_COLUMN_ENTRY_WSTR("PK_TABLE_CATALOG", 1, PKTableCatalog)
		PROVIDER_COLUMN_ENTRY_WSTR("PK_TABLE_SCHEMA", 2, PKTableSchema)
		PROVIDER_COLUMN_ENTRY_WSTR("PK_TABLE_NAME", 3, PKTableName)
		PROVIDER_COLUMN_ENTRY_WSTR("PK_COLUMN_NAME", 4, PKColumnName)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("PK_COLUMN_GUID", 5, DBTYPE_GUID, 0xFF, 0xFF, PKColumnGUID)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("PK_COLUMN_PROPID", 6, DBTYPE_UI4, 10, ~0, PKColumnPropID)
		PROVIDER_COLUMN_ENTRY_WSTR("FK_TABLE_CATALOG", 7, FKTableCatalog)
		PROVIDER_COLUMN_ENTRY_WSTR("FK_TABLE_SCHEMA", 8, FKTableSchema)
		PROVIDER_COLUMN_ENTRY_WSTR("FK_TABLE_NAME", 9, FKTableName)
		PROVIDER_COLUMN_ENTRY_WSTR("FK_COLUMN_NAME", 10, FKColumnName)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("FK_COLUMN_GUID", 11, DBTYPE_GUID, 0xFF, 0xFF, FKColumnGUID)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("FK_COLUMN_PROPID", 12, DBTYPE_UI4, 10, ~0, FKColumnPropID)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("ORDINAL", 13, DBTYPE_UI4, 10, ~0, ColumnOrdinal)
		PROVIDER_COLUMN_ENTRY_WSTR("UPDATE_RULE", 14, UpdateRule)
		PROVIDER_COLUMN_ENTRY_WSTR("DELETE_RULE", 15, DeleteRule)
		PROVIDER_COLUMN_ENTRY_TYPE_PS("KEY_SEQ", 16, DBTYPE_UI4, 10, ~0, KeySeq)
		PROVIDER_COLUMN_ENTRY_WSTR("FK_NAME", 17, foreignKeyName)
		PROVIDER_COLUMN_ENTRY_WSTR("PK_NAME", 18, primaryKeyName)
	END_PROVIDER_COLUMN_MAP()
};

/////////////////////////////////////////////////////////////////////////////
class CSRS_DBSCHEMA_FOREIGN_KEYS :
	public CSchemaRowsetImpl<CSRS_DBSCHEMA_FOREIGN_KEYS, CForeignKeysRow, CSQLiteSession, CAtlArray<CForeignKeysRow>, CSimpleRow, IRowsetLocateImpl<CSRS_DBSCHEMA_FOREIGN_KEYS, IRowsetLocate>>

{
public:

	ROWSET_PROPSET_MAP(CSRS_DBSCHEMA_FOREIGN_KEYS)

	ULONG record_count;
	LONG ordinal_pos;
	std::string table;
	sqlite3* db;

	/////////////////////////////////////////////////////////////////////////////
	DBSTATUS GetDBStatus(CSimpleRow*, ATLCOLUMNINFO* pInfo)
	{
		return DBSTATUS_S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT Execute(LONG* pcRowsAffected, ULONG, const VARIANT*)
	{
		HRESULT hr;
		CComPtr<IGetDataSource> ipGDS;
		if (FAILED(hr = GetSpecification(IID_IGetDataSource, (IUnknown **)&ipGDS))) return hr;		
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);

		// Step1: Create an SQL statement to get columns for all TABLES
		std::vector<std::string> TABLES = SQLiteTables(pSess->db);

		// Step2: Execute the SQL statement
		record_count=0;		 
		db = pSess->db;
		for(size_t i=0;i<TABLES.size();i++)
		{
			ordinal_pos=0;
			table = TABLES[i];
			if(table=="sqlite_sequence") continue;
			std::string sql = std::string("PRAGMA table_info(") + table + ")";
			char *zErrMsg = 0;
			int rc = sqlite3_exec(pSess->db, sql.c_str(), sqlite_callback_foreign_keys, this, &zErrMsg);
			if(rc) return E_FAIL;
		}
		*pcRowsAffected = record_count;		

		// Bookmarks
		m_rgBookmarks.SetCount(record_count+3);
		m_rgBookmarks[0] = m_rgBookmarks[1] = m_rgBookmarks[2] = -1;
		for(ULONG i=3;i<record_count+3;i++) { m_rgBookmarks[i] = i-2; }

		return S_OK;
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	static int sqlite_callback_foreign_keys(void *NotUsed, int argc, char **argv, char **azColName)	
	{		
		CSRS_DBSCHEMA_FOREIGN_KEYS* objPtr = (CSRS_DBSCHEMA_FOREIGN_KEYS*) NotUsed;		

		// This is the information we need to fill in for DBSCHEMA_COLUMNS.
		// 0:cid, 1:name, 2:type, 3:notnull, 4:dflt_value, 5:pk

		CSQLiteColumnsRowsetRow TR_PK;		
		ATLCOLUMNINFO col;
		DBBYTEOFFSET offset=0;

		std::string TABLE(objPtr->table);
		std::string COLUMN(argv[1]);

		TR_PK.InitColumnMetadata(objPtr->db, TABLE, COLUMN, ++(objPtr->ordinal_pos), col, offset);

		if(TR_PK.m_DBCOLUMN_SQLITE_FKEYCOLUMN==VARIANT_TRUE)
		{
			std::string FK(CW2A(TR_PK.m_DBCOLUMN_SQLITE_FKEYCOLUMN_NAME));
			std::string FT(CW2A(TR_PK.m_DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME));

			CSQLiteColumnsRowsetRow TR_FK;
			TR_FK.InitColumnMetadata(objPtr->db, FT, FK, ++(objPtr->ordinal_pos), col, offset);

			CForeignKeysRow trData;
			
			BSTR PK_NAME = UTF8_to_BSTR(TR_PK.FILED_NAME.c_str());
			BSTR FK_NAME = UTF8_to_BSTR(TR_FK.FILED_NAME.c_str());

			trData.PKColumnGUID				= TR_FK.m_DBCOLUMN_GUID;			
			::wcscpy(trData.PKColumnName,	TR_FK.m_DBCOLUMN_BASECOLUMNNAME);
			::wcscpy(trData.PKTableCatalog, TR_FK.m_DBCOLUMN_BASECATALOGNAME);
			::wcscpy(trData.PKTableName,	TR_FK.m_DBCOLUMN_BASETABLENAME);
			::wcscpy(trData.PKTableSchema,	TR_FK.m_DBCOLUMN_BASESCHEMANAME);

			trData.FKColumnGUID				= TR_PK.m_DBCOLUMN_GUID;			
			::wcscpy(trData.FKColumnName,	TR_PK.m_DBCOLUMN_BASECOLUMNNAME);
			::wcscpy(trData.FKTableCatalog,	TR_PK.m_DBCOLUMN_BASECATALOGNAME);
			::wcscpy(trData.FKTableName,	TR_PK.m_DBCOLUMN_BASETABLENAME);
			::wcscpy(trData.FKTableSchema,	TR_PK.m_DBCOLUMN_BASESCHEMANAME);

			trData.ColumnOrdinal			= TR_PK.m_DBCOLUMN_NUMBER;
			::wcscpy(trData.primaryKeyName,	L"PrimaryKey");
			::wcscpy(trData.foreignKeyName,	L"Reference");

			trData.PKColumnPropID = 1;				
			trData.FKColumnPropID = 1;
			trData.KeySeq		  = 1;

			::wcscpy(trData.UpdateRule,	L"");
			::wcscpy(trData.DeleteRule,	L"");

			SysFreeString(PK_NAME);
			SysFreeString(FK_NAME);

			objPtr->record_count++;
			objPtr->m_rgRowData.Add(trData);		
		}

		return 0;
	}
};
