#pragma once

#include "SQLiteOLEDB.h"
#include "SQLiteSession.hpp"

///////////////////////////////////////////////////////////////////////////
// class ICommandWithPamametersImpl
static const struct StandardDataType
{
	const OLECHAR *typeName;
	DBTYPE dbType;
} standardDataTypes[] =
{
	{ OLESTR("DBTYPE_I2"), DBTYPE_I2 },
	{ OLESTR("DBTYPE_UI2"), DBTYPE_I2 },
	{ OLESTR("DBTYPE_I4"), DBTYPE_I4 },
	{ OLESTR("DBTYPE_UI4"), DBTYPE_UI4 },
	{ OLESTR("DBTYPE_R4"), DBTYPE_R4 },
	{ OLESTR("DBTYPE_R8"), DBTYPE_R8 },
	{ OLESTR("DBTYPE_BOOL"), DBTYPE_I2 },
	{ OLESTR("DBTYPE_VARIANT"), DBTYPE_VARIANT },
	{ OLESTR("DBTYPE_IUNKNOWN"), DBTYPE_IUNKNOWN },
	{ OLESTR("DBTYPE_DATE"), DBTYPE_DATE },
	{ OLESTR("DBTYPE_BSTR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_CHAR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_WSTR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_VARCHAR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_LONGVARCHAR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_WCHAR"), DBTYPE_BSTR },
	{ OLESTR("DBTYPE_BINARY"), DBTYPE_IUNKNOWN },
	{ OLESTR("DBTYPE_VARBINARY"), DBTYPE_IUNKNOWN },
	{ OLESTR("DBTYPE_LONGVARBINARY"), DBTYPE_IUNKNOWN },
	{ OLESTR("DBTYPE_GEOMETRY"), DBTYPE_IUNKNOWN },
	{ OLESTR("DBTYPE_VARIANT"), DBTYPE_VARIANT }
};

/////////////////////////////////////////////////////////////////////////////
template <class T>
class ATL_NO_VTABLE ICommandWithPamametersImpl :
	public ICommandWithParameters
{
public:
	ICommandWithPamametersImpl(void) :
	  
	  m_nSetParams(0)
	  {
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  STDMETHOD(GetParameterInfo)(ULONG *pcParams, DBPARAMINFO **prgParamInfo, OLECHAR **ppNamesBuffer)
	  {
		  HRESULT hr = E_FAIL;
		  if (!pcParams || !prgParamInfo || !ppNamesBuffer) return E_INVALIDARG;
		  T *pT = 0;
		  *pcParams = 0;
		  return DB_E_PARAMUNAVAILABLE;
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  STDMETHOD(MapParameterNames)(ULONG cParamNames, const OLECHAR *rgParamNames[], LONG rgParamOrdinals[])
	  {
		  if(cParamNames==0) return S_OK;
		  if (cParamNames > 0 && (rgParamNames == 0 || rgParamOrdinals == 0)) return E_INVALIDARG;
		  for (ULONG i = 0; i < cParamNames; i++)
			  rgParamOrdinals[i] = GetSpatialParamOrdinal(rgParamNames[i]);
		  return S_OK;
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  STDMETHOD(SetParameterInfo)(ULONG cParams, const ULONG rgParamOrdinals[], const DBPARAMBINDINFO rgParamBindInfo[])
	  {
		  HRESULT hr;
		  if (cParams == 0)
		  {
			  m_paramInfo.RemoveAll();
			  m_nSetParams = 0;
			  return S_OK;
		  }

		  if (cParams > 0 && rgParamOrdinals == 0) return E_INVALIDARG;
		  if (rgParamBindInfo != 0)
		  {
			  bool allParamNamesSet = true, oneParamNameSet = false;
			  for (ULONG i = 0; i < cParams; i++)
			  {
				  //  We do not handle default parameter conversion
				  if(rgParamBindInfo[i].pwszDataSourceType == 0)	return E_INVALIDARG;
				  DBTYPE dbType;
				  if((dbType = CheckDataType(rgParamBindInfo[i].pwszDataSourceType)) == DBTYPE_EMPTY) return DB_E_BADTYPENAME;
				  if( rgParamBindInfo[i].pwszName == 0 )
					  allParamNamesSet = false;
				  else
					  oneParamNameSet = true;
				  if( rgParamBindInfo[i].dwFlags & ~(DBPARAMIO_INPUT | DBPARAMIO_OUTPUT) )
					  return E_INVALIDARG;
			  }
			  if( oneParamNameSet && ! allParamNamesSet ) return DB_E_BADPARAMETERNAME;
		  }

		  for (ULONG i = 0; i < cParams; i++)
		  {
			  //  Look for the parameter already in the list
			  for (int j = m_nSetParams; --j >= 0;)
			  {
				  if( rgParamOrdinals[i] == m_paramInfo[j].iOrdinal )
					  break;
			  }

			  if (j >= 0)
			  {
				  //  Discard the type info. for this parameter
				  if( rgParamBindInfo == 0 )
				  {
					  m_paramInfo[j].~ParamInfo();
					  while( j < (int)m_nSetParams )
					  {
						  m_paramInfo[j] = m_paramInfo[j + 1];
						  j++;
					  }
					  --m_nSetParams;
				  }
				  //  Change parameter type info???
				  else
				  {
				  }
			  }
			  else if (rgParamBindInfo != 0)
			  {
				  ParamInfo tempInfo;
				  m_paramInfo.SetAtGrow( m_nSetParams, tempInfo );
				  ParamInfo &paramInfo = m_paramInfo[m_nSetParams];
				  if (FAILED(hr = paramInfo.Set( rgParamOrdinals[i], rgParamBindInfo[i]))) return hr;
				  m_nSetParams++;
			  }
		  }
		  return S_OK;
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  static DBTYPE CheckDataType( const OLECHAR *pDataTypeName )
	  {
		  const StandardDataType *pDataTypes = standardDataTypes;
		  ATLASSERT(pDataTypeName);
		  for (int i = sizeof(standardDataTypes)/sizeof(standardDataTypes[0]); --i >= 0;)
		  {
			  if (_wcsicmp( pDataTypes->typeName, pDataTypeName) == 0)
				  return pDataTypes->dbType;
			  pDataTypes++;
		  }
		  return DBTYPE_EMPTY;
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  static ULONG GetSpatialParamOrdinal(const OLECHAR *paramName)
	  {
		  static const OLECHAR *paramNames[] =
		  {
			  OLESTR( "FILTER" ),
			  OLESTR( "OPERATOR" ),
			  OLESTR( "COL_NAME" )
		  };

		  for (int i = 0; i < sizeof(paramNames)/sizeof(paramNames[0]); i++)
		  {
			  if (_wcsicmp(paramNames[i], paramName) == 0)
			  {
				  return i + 1;
			  }
		  }

		  return 0;
	  }

	  /////////////////////////////////////////////////////////////////////////////
	  struct ParamInfo : public DBPARAMINFO
	  {
		  ParamInfo(void)
		  {
			  this->iOrdinal = 0;
			  this->pwszName = 0;
			  this->ulParamSize = 0;
			  this->dwFlags = DBPARAMFLAGS_ISINPUT;
			  this->bPrecision = 0;
			  this->bScale = 0;
			  this->wType = 0;
			  this->pTypeInfo = 0;
		  }

		  ~ParamInfo( void )
		  {
			  if (this->pwszName != 0)
				  delete [] this->pwszName;
		  }

		  HRESULT Set(ULONG ordinal, const DBPARAMBINDINFO &rParamBindInfo)
		  {
			  this->iOrdinal = ordinal;
			  if (this->pwszName != 0)
			  {
				  delete [] this->pwszName;
				  this->pwszName = 0;
			  }

			  if (rParamBindInfo.pwszName != 0)
			  {
				  this->pwszName = new OLECHAR[::wcslen( rParamBindInfo.pwszName ) + 1];
				  if (this->pwszName == 0)
					  return E_OUTOFMEMORY;
				  ::wcscpy( this->pwszName, rParamBindInfo.pwszName );
			  }

			  ATLASSERT(rParamBindInfo.pwszDataSourceType);
			  this->wType = CheckDataType(rParamBindInfo.pwszDataSourceType);
			  this->ulParamSize = rParamBindInfo.ulParamSize;
			  this->dwFlags = rParamBindInfo.dwFlags; 
			  this->bPrecision = rParamBindInfo.bPrecision;
			  this->bScale = rParamBindInfo.bScale;

			  return S_OK;
		  }
	  };

	  CArray<ParamInfo, ParamInfo &> m_paramInfo;
	  ULONG                          m_nSetParams;
};


// ==================================================================================================================================
//	   ___________ ____    __    _ __       ______                                          __
//	  / ____/ ___// __ \  / /   (_) /____  / ____/___  ____ ___  ____ ___  ____ _____  ____/ /
//	 / /    \__ \/ / / / / /   / / __/ _ \/ /   / __ \/ __ `__ \/ __ `__ \/ __ `/ __ \/ __  / 
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/ /___/ /_/ / / / / / / / / / / / /_/ / / / / /_/ /  
//	\____//____/\___\_\/_____/_/\__/\___/\____/\____/_/ /_/ /_/_/ /_/ /_/\__,_/_/ /_/\__,_/   
//	                                                                                          
// ==================================================================================================================================

class ATL_NO_VTABLE CSQLiteCommand : 
	public CComObjectRootEx<THREAD_MODEL>,
	public IAccessorImpl<CSQLiteCommand>,
	public ICommandTextImpl<CSQLiteCommand>,
	public ICommandPropertiesImpl<CSQLiteCommand>,
	public IObjectWithSiteImpl<CSQLiteCommand>,
	public IConvertTypeImpl<CSQLiteCommand>,
	public IColumnsInfoImpl<CSQLiteCommand>,
	public ICommandWithPamametersImpl<CSQLiteCommand>
{
public:

	std::string SQL;
	sqlite3 *db;

	COMMAND_PROPSET_MAP(CSQLiteCommand)

	/////////////////////////////////////////////////////////////////////////////
	BEGIN_COM_MAP(CSQLiteCommand)
		COM_INTERFACE_ENTRY(IAccessor)
		COM_INTERFACE_ENTRY(ICommand)
		COM_INTERFACE_ENTRY2(ICommandText, ICommand)
		COM_INTERFACE_ENTRY(IObjectWithSite)		
		COM_INTERFACE_ENTRY(ICommandProperties)		
		COM_INTERFACE_ENTRY(IColumnsInfo)
		COM_INTERFACE_ENTRY(IConvertType)
		COM_INTERFACE_ENTRY(ICommandWithParameters)
	END_COM_MAP()

	/////////////////////////////////////////////////////////////////////////////
	HRESULT FinalConstruct()
	{
		HRESULT hr = CConvertHelper::FinalConstruct();
		if (FAILED (hr)) return hr;

		hr = IAccessorImpl<CSQLiteCommand>::FinalConstruct();
		if (FAILED(hr)) return hr;

		return CUtlProps<CSQLiteCommand>::FInit();
	}

	/////////////////////////////////////////////////////////////////////////////
	void FinalRelease()
	{
		IAccessorImpl<CSQLiteCommand>::FinalRelease();
	}
	
/////////////////////////////////////////////////////////////////////////////
// ICommand
/////////////////////////////////////////////////////////////////////////////

	static ATLCOLUMNINFO* GetColumnInfo(CSQLiteCommand* pv, ULONG* pcInfo)
	{
		return NULL;
	}

	/////////////////////////////////////////////////////////////////////////////
	STDMETHODIMP SetCommandText(REFGUID rguidDialect, LPCOLESTR pwszCommand)
	{						
		USES_CONVERSION;
		
		BSTR bstr = SysAllocString(pwszCommand);
		SQL = BSTR_to_UTF8(bstr);
		SysFreeString(bstr);

		// Special care for @@IDENTIRY even though it 
		// is not thread-safe and can lead to errors
		// if used with multiple client cursors on
		// the same database file.
		if(SQL=="SELECT @@IDENTITY")
		{
			SQL = "SELECT last_insert_rowid()";
			pwszCommand = A2COLE(SQL.c_str());
		}

		SQL = SQL + "\n";
		OutputDebugStringA(SQL.c_str());
		
		CComPtr<IGetDataSource> ipGDS;
		m_spUnkSite->QueryInterface(IID_IGetDataSource, (void **)&ipGDS);
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);
		db = pSess->db;
		
		HRESULT hr = ICommandTextImpl<CSQLiteCommand>::SetCommandText(rguidDialect, pwszCommand);
		if (FAILED(hr)) return hr;

		return S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT WINAPI Execute(IUnknown * pUnkOuter, REFIID riid, DBPARAMS * pParams, LONG * pcRowsAffected, IUnknown ** ppRowset)
	{
		HRESULT hr = S_OK;		

		// Execute a SELECT query and return a Rowset
		if(riid!=GUID_NULL)
		{
			CSQLiteRowset* pRowset;
			return CreateRowset(pUnkOuter, riid, pParams, pcRowsAffected, ppRowset, pRowset);			
		}

		// === Execute an UPDATE/DELETE query and return affected rows == //
		// This is called by Microsoft Client Cursor Engine for updating client Recordsets.
		// Batch update flag must be specified when opening the ADO Recordset.

		*pcRowsAffected = 0;
		
		// Get accessor bindings
		ATLBINDINGS *pBinding = NULL;
		if(!m_rgBindings.Lookup((ULONG)pParams->hAccessor, pBinding) || pBinding==NULL) return DB_E_BADACCESSORHANDLE;
		if(!(pBinding->dwAccessorFlags & DBACCESSOR_PARAMETERDATA)) return DB_E_BADACCESSORTYPE;

		// Prepare an SQL statement
		int rc = 0;
		sqlite3_stmt* stmt;
		rc = sqlite3_prepare_v2(db, SQL.c_str(), -1, &stmt, nullptr);
		if(rc!=SQLITE_OK) return DB_E_DIALECTNOTSUPPORTED;
		DBCOUNTITEM params = (DBCOUNTITEM) sqlite3_bind_parameter_count(stmt);

		// The Microsoft Client Cursor Engine generates an update SQL statement
		// in the form: UPDATE <table> SET <value> WHERE <id>=? AND <old_values>
		// This is causing a severe problem with FLOAT numbers in the WHERE 
		// section and therefore we need to do some heuristics here and remove
		// the portion of the SQL after the AND if floats are detected.		

		bool sql_has_floats = false;
		for(DBCOUNTITEM i=0;i<params;i++)
		{
			DBBINDING &pBindCur = pBinding->pBindings[i];
			switch(pBindCur.wType)
			{
			case DBTYPEENUM::DBTYPE_R4:
			case DBTYPEENUM::DBTYPE_R8:
				sql_has_floats=true;
				break;
			}
		}
		sql_has_floats=true;
		if(sql_has_floats)
		{
			std::size_t found = SQL.find(" AND ");
			if(found!=std::string::npos) 
			{
				SQL = SQL.substr(0, found);			
				rc = sqlite3_finalize(stmt);
				rc = sqlite3_prepare_v2(db, SQL.c_str(), -1, &stmt, nullptr);
				if(rc!=SQLITE_OK) return DB_E_DIALECTNOTSUPPORTED;
			}
		}						

		// Loop on bindings and bind the parameters to SQL statement		
		params = (DBCOUNTITEM) sqlite3_bind_parameter_count(stmt);
		for(DBCOUNTITEM i=0;i<params;i++)
		{
			sqlite3_reset(stmt);
				
			// Get a binding
			DBBINDING &pBindCur = pBinding->pBindings[i];
			DWORD dwPart = pBindCur.dwPart;

			// Check param value for null
			DBSTATUS* dbStat = (dwPart & DBPART_STATUS) ? (DBSTATUS *)((BYTE*)pParams->pData+pBindCur.obStatus) : NULL;
			if(dbStat && *dbStat==DBSTATUS_S_ISNULL)
			{
				rc = sqlite3_bind_null(stmt, i+1);
				if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				continue;
			}

			// Get param data type
			DBTYPE wType = pBindCur.wType & ~DBTYPE_BYREF;

			// Get param value offset
			void *pDataValue = (dwPart & DBPART_VALUE) ? ((BYTE*)pParams->pData+pBindCur.obValue) : NULL;
			if((pBindCur.wType & DBTYPE_BYREF) && pDataValue) pDataValue = *(void **)pDataValue;

			// Get param data length
			DBLENGTH cbBytesLen = 0;
			if(dwPart & DBPART_LENGTH)
			{
				cbBytesLen = *(DBLENGTH *)((BYTE*)pParams->pData+pBindCur.obLength);
			}

			switch(wType)
			{
			case DBTYPEENUM::DBTYPE_I1:
				{
					INT8 v = *(INT8*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_I2:		
				{
					INT16 v = *(INT16*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_I4:	
				{
					INT32 v = *(INT32*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_I8:		
				{
					INT64 v = *(INT64*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int64(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_UI1:	
				{
					UINT8 v = *(UINT8*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_UI2:	
				{
					UINT16 v = *(UINT16*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_UI4:	
				{
					UINT32 v = *(UINT32*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_UI8:	
				{
					UINT64 v = *(UINT64*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int64(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_R4:		
				{
					double v = (double) *(float*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_double(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_R8:		
				{
					double v = *(double*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_double(stmt, i+1, v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_CY:
				{
					CY	cy = *(CY*)pDataValue;
					double v = double(cy.int64 / 10000.0);
					rc = sqlite3_bind_double(stmt, i+1,  v);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_DECIMAL:
				{
					DECIMAL	v = *(DECIMAL*)pDataValue;
					BSTR pSrcData = SysAllocString(L"");
					hr = m_spConvert->DataConvert(DBTYPEENUM::DBTYPE_DECIMAL, DBTYPEENUM::DBTYPE_BSTR, cbBytesLen, &cbBytesLen, (void*)&v, (void*)&pSrcData, pBindCur.cbMaxLen, *dbStat, dbStat, pBindCur.bPrecision, pBindCur.bScale, 0);
					if(hr==S_OK) 
					{							
						std::string value = BSTR_to_UTF8(pSrcData);
						rc = sqlite3_bind_text(stmt, i+1, value.c_str(), value.size(), SQLITE_TRANSIENT);
						if(rc!=SQLITE_OK) hr = DB_E_UNSUPPORTEDCONVERSION;
					}
					SysFreeString(pSrcData);
				}
				break;

			case DBTYPEENUM::DBTYPE_BOOL:	
				{
					VARIANT_BOOL v = *(VARIANT_BOOL*)((BYTE*)(pDataValue));
					rc = sqlite3_bind_int(stmt, i+1, v==VARIANT_TRUE ? 1 : 0);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;
				}
				break;

			case DBTYPEENUM::DBTYPE_DBDATE:
				{
					DBDATE dt = *(DBDATE*)((BYTE*)(pDataValue)); 
					char v[80];
					sprintf(v, "%04d-%02d-%02d\0", dt.year, dt.month, dt.day);
					rc = sqlite3_bind_text(stmt, i+1, v, 80, SQLITE_TRANSIENT);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;						
				}
				break;

			case DBTYPEENUM::DBTYPE_DBTIME:
				{
					DBTIME dt = *(DBTIME*)((BYTE*)(pDataValue)); 
					char v[80];
					sprintf(v, "%02d:%02d:%02d\0", dt.hour, dt.minute, dt.second);
					rc = sqlite3_bind_text(stmt, i+1, v, 80, SQLITE_TRANSIENT);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;						
				}
				break;

			case DBTYPEENUM::DBTYPE_DBTIMESTAMP:
				{
					DBTIMESTAMP dt = *(DBTIMESTAMP*)((BYTE*)(pDataValue)); 						
					char v[80];
					sprintf(v, "%04d-%02d-%02d %02d:%02d:%02d\0", dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second);
					rc = sqlite3_bind_text(stmt, i+1, v, 80, SQLITE_TRANSIENT);
					if(rc!=SQLITE_OK) return DB_E_UNSUPPORTEDCONVERSION;						
				}
				break;

			case DBTYPE_BLOB:
				{
					VARIANT* var = (VARIANT *)pDataValue;					
					switch(var->vt)
					{
					case VT_NULL:						
						rc = sqlite3_bind_null(stmt, i+1);
						hr = S_OK;
						break;

					case VT_ARRAY|VT_UI1:												
						{
							SAFEARRAY* psa = V_ARRAY(var);
							long lBound, uBound;
							SafeArrayGetLBound(psa, 1, &lBound);
							SafeArrayGetUBound(psa, 1, &uBound);						
							cbBytesLen = uBound - lBound + 1;
							char *blob;
							SafeArrayAccessData(psa, (void**)&blob);
							rc = sqlite3_bind_blob(stmt, i+1, blob, cbBytesLen, SQLITE_TRANSIENT);
							SafeArrayUnaccessData(psa);
							hr = S_OK;
						}
						break;

					default:
						hr = DB_E_UNSUPPORTEDCONVERSION;
					}
				}
				break;

			default:
				{
					// Read from ISequentialStream					
					if(pBindCur.pObject && pBindCur.pObject->iid==IID_ISequentialStream)
					{
						IUnknown* pUnknown = *(IUnknown**)(pDataValue);
						ISequentialStream* pSequentialStream = NULL;
						hr = pUnknown->QueryInterface(pBindCur.pObject->iid, (void**)&pSequentialStream);
						if(hr==S_OK)
						{
							// TODO: We should probably loop and read in chunks.
							WCHAR Buffer[MAX_COLUMN_BYTES];							
							pSequentialStream->Read(Buffer, cbBytesLen, &cbBytesLen);																					
							Buffer[cbBytesLen/2] = 0;
							BSTR pSrcData = SysAllocString(Buffer);
							std::string value = BSTR_to_UTF8(pSrcData);
							SysFreeString(pSrcData);
							rc = sqlite3_bind_text(stmt, i+1, value.c_str(), value.size(), SQLITE_TRANSIENT);
							if(rc!=SQLITE_OK) hr = DB_E_UNSUPPORTEDCONVERSION;
						}
					}

					// Read from buffer
					else
					{
						BSTR pSrcData = SysAllocString(L"");
						hr = m_spConvert->DataConvert(pBindCur.wType, DBTYPEENUM::DBTYPE_BSTR, cbBytesLen, &cbBytesLen, pDataValue, (void*)&pSrcData, pBindCur.cbMaxLen, *dbStat, dbStat, pBindCur.bPrecision, pBindCur.bScale, 0);
						if(hr==S_OK) 
						{							
							std::string value = BSTR_to_UTF8(pSrcData);
							rc = sqlite3_bind_text(stmt, i+1, value.c_str(), value.size(), SQLITE_TRANSIENT);
							if(rc!=SQLITE_OK) hr = DB_E_UNSUPPORTEDCONVERSION;
						}
						SysFreeString(pSrcData);
					}

					// Check for errors
					if(hr!=S_OK)
					{
						rc = sqlite3_finalize(stmt);
						return hr;
					}
				}
				break;
			}
		}

		// Finally execute the statement
		rc = sqlite3_step(stmt);
		if(rc!=SQLITE_DONE)
			return DB_E_ERRORSINCOMMAND;

		rc = sqlite3_finalize(stmt);
		if(rc!=SQLITE_OK)
			return DB_E_ERRORSINCOMMAND;

		*pcRowsAffected = 1;
		return hr;
	}
};

