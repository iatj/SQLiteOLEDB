#pragma once

#include "SQLiteOLEDB.h"
 
// ==================================================================================================================================
//	   ___________ ____    __    _ __       ______      __                           ____                          __  ____                          __ 
//	  / ____/ ___// __ \  / /   (_) /____  / ____/___  / /_  ______ ___  ____  _____/ __ \____ _      __________  / /_/ __ \____ _      __________  / /_
//	 / /    \__ \/ / / / / /   / / __/ _ \/ /   / __ \/ / / / / __ `__ \/ __ \/ ___/ /_/ / __ \ | /| / / ___/ _ \/ __/ /_/ / __ \ | /| / / ___/ _ \/ __/
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/ /___/ /_/ / / /_/ / / / / / / / / (__  ) _, _/ /_/ / |/ |/ (__  )  __/ /_/ _, _/ /_/ / |/ |/ (__  )  __/ /_  
//	\____//____/\___\_\/_____/_/\__/\___/\____/\____/_/\__,_/_/ /_/ /_/_/ /_/____/_/ |_|\____/|__/|__/____/\___/\__/_/ |_|\____/|__/|__/____/\___/\__/  
//	                                                                                                                                                    
// ==================================================================================================================================

#define DBPRECISION	10
#define SQLITE_REQUIRED_METADATA_COLUMNS (UINT)11
#define SQLITE_OPTIONAL_METADATA_COLUMNS (UINT)24
	
/////////////////////////////////////////////////////////////////////////////
// SQLite Specific ColumnsRowset Columns
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_FKEYCOLUMN				= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)42};
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME	= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)43};
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_FKEYCOLUMN_NAME			= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)44};
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_DATATYPE				= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)45};
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_DECLATED_DATATYPE		= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)46};
extern const OLEDBDECLSPEC DBID DBCOLUMN_SQLITE_CAN_BE_NULL				= {DBCIDGUID, DBKIND_GUID_PROPID, (LPOLESTR)47};

/////////////////////////////////////////////////////////////////////////////
#define SQLITE_COLUMN_ENTRY(Ordinal, dbid, Name, Size, Type, Precision, Scale, Flags, Variable) \
{ \
	(LPOLESTR)OLESTR(Name), \
	NULL, \
	(DBORDINAL)Ordinal, \
	Flags, \
	Size, \
	Type, \
	(BYTE)Precision, \
	(BYTE)Scale, \
	{ \
		EXPANDGUID(dbid.uGuid.guid), \
		dbid.eKind, \
		dbid.uName.pwszName  \
	}, \
	offsetof(_Class, Variable) \
},

/////////////////////////////////////////////////////////////////////////////
const struct
{
	std::string sName;
	DBTYPE wType;
} OLEDB_TYPE_NAMES[] = 
{
	{ "DBTYPE_EMPTY"		, DBTYPE_EMPTY       },
	{ "DBTYPE_ARRAY"        , DBTYPE_ARRAY       },
	{ "DBTYPE_BOOL"         , DBTYPE_BOOL        },
	{ "DBTYPE_BSTR"         , DBTYPE_BSTR        },
	{ "DBTYPE_BYREF"        , DBTYPE_BYREF       },
	{ "DBTYPE_BYTES"        , DBTYPE_BYTES       },
	{ "DBTYPE_CY"	        , DBTYPE_CY	         },
	{ "DBTYPE_DATE"         , DBTYPE_DATE        },
	{ "DBTYPE_DBDATE"       , DBTYPE_DBDATE      },
	{ "DBTYPE_DBTIME"       , DBTYPE_DBTIME      },
	{ "DBTYPE_DBTIMESTAMP"  , DBTYPE_DBTIMESTAMP },
	{ "DBTYPE_DECIMAL"      , DBTYPE_DECIMAL     },
	{ "DBTYPE_ERROR"        , DBTYPE_ERROR       },
	{ "DBTYPE_FILETIME"     , DBTYPE_FILETIME    },
	{ "DBTYPE_GUID"         , DBTYPE_GUID        },
	{ "DBTYPE_HCHAPTER"     , DBTYPE_HCHAPTER    },
	{ "DBTYPE_I1"           , DBTYPE_I1          },
	{ "DBTYPE_I2"	        , DBTYPE_I2	         },
	{ "DBTYPE_I4"	        , DBTYPE_I4	         },
	{ "DBTYPE_I8"           , DBTYPE_I8          },
	{ "DBTYPE_IDISPATCH"    , DBTYPE_IDISPATCH   },
	{ "DBTYPE_IUNKNOWN"     , DBTYPE_IUNKNOWN    },
	{ "DBTYPE_NULL"	        , DBTYPE_NULL	     },
	{ "DBTYPE_NUMERIC"      , DBTYPE_NUMERIC     },
	{ "DBTYPE_PROPVARIANT"  , DBTYPE_PROPVARIANT },
	{ "DBTYPE_R4"	        , DBTYPE_R4	         },
	{ "DBTYPE_R8"	        , DBTYPE_R8	         },
	{ "DBTYPE_RESERVED"     , DBTYPE_RESERVED    },
	{ "DBTYPE_STR"          , DBTYPE_STR         },
	{ "DBTYPE_UDT"          , DBTYPE_UDT         },
	{ "DBTYPE_UI1"          , DBTYPE_UI1         },
	{ "DBTYPE_UI2"          , DBTYPE_UI2         },
	{ "DBTYPE_UI4"          , DBTYPE_UI4         },
	{ "DBTYPE_UI8"          , DBTYPE_UI8         },
	{ "DBTYPE_VARIANT"      , DBTYPE_VARIANT     },
	{ "DBTYPE_VARNUMERIC"   , DBTYPE_VARNUMERIC  },
	{ "DBTYPE_VECTOR"       , DBTYPE_VECTOR      },
	{ "DBTYPE_WSTR"         , DBTYPE_WSTR        },

	{ "DBTYPE_ARRAY|DBTYPE_BYTES" , DBTYPE_BYTES|DBTYPE_ARRAY }
};	  

/////////////////////////////////////////////////////////////////////////////
std::string OLEDB_TYPENAME(int wType)
{
	wType &= ~DBTYPE_BYREF;
	
	for(int i=1;i<sizeof(OLEDB_TYPE_NAMES)/sizeof(*OLEDB_TYPE_NAMES);i++)
	{
		if(OLEDB_TYPE_NAMES[i].wType==wType)
			return OLEDB_TYPE_NAMES[i].sName;
	}
	return "(UNKNOWN)";
}
  																																		
/////////////////////////////////////////////////////////////////////////////
class CSQLiteColumnsRowsetRow
{
public:

	std::string	FILED_NAME;
	std::string BASECOLUMNNAME;
	std::string BASETABLENAME;

	// Mandatory columns

	WCHAR			   	m_DBCOLUMN_IDNAME[128];		// Column name. This column, together with the DBCOLUMN_GUID and DBCOLUMN_PROPID columns, forms the ID of the column. One or more (but not all) of these columns will be NULL, depending on which elements of the DBID structure the provider uses. The column ID of a base table should be invariant under views.	
	GUID		   		m_DBCOLUMN_GUID;			// Column GUID
	UINT32			   	m_DBCOLUMN_PROPID;			// Column property ID (for ColumnRowset Columns - not for Data Columns)
	WCHAR			   	m_DBCOLUMN_NAME[128];		// The name of the column; this might not be unique. If this cannot be determined, a NULL is returned. The name can be different from the value returned in DBCOLUMN_IDNAME if the column has been renamed by the command text. This name always reflects the most recent renaming of the column in the current view or command text.
	ULONG				m_DBCOLUMN_NUMBER;			// The ordinal of the column. This is zero for the bookmark column of the row, if any. Other columns are numbered starting with one. This column cannot contain a NULL value.
	DBTYPEENUM		   	m_DBCOLUMN_TYPE;			// The indicator of the column's data type. If the data type of the column varies from row to row, this must be DBTYPE_VARIANT. This column cannot contain a NULL value. For a list of valid type indicators, see http://msdn.microsoft.com/en-us/library/windows/desktop/ms711251(v=vs.85).aspx
	IUnknown*		   	m_DBCOLUMN_TYPEINFO;		// Reserved for future use. 
		
	ULONG	    		m_DBCOLUMN_COLUMNSIZE;		// The maximum possible length of a value in the column. For columns that use a fixed-length data type, this is the size of the data type. For columns that use a variable-length data type, this is one of the following:
													// The maximum length of the column in characters (for DBTYPE_STR and DBTYPE_WSTR) or in bytes (for DBTYPE_BYTES and DBTYPE_VARNUMERIC), if one is defined. For example, a CHAR(5) column in an SQL table has a maximum length of 5.
													// The maximum length of the data type in characters (for DBTYPE_STR and DBTYPE_WSTR) or in bytes (for DBTYPE_BYTES and DBTYPE_VARNUMERIC), if the column does not have a defined length.
													// ~0 (bitwise, the value is not 0; all bits are set to 1) if neither the column nor the data type has a defined maximum length.
													// For data types that do not have a length, this is set to ~0 (bitwise, the value is not 0; all bits are set to 1).

	UINT16			   	m_DBCOLUMN_PRECISION;		// If DBCOLUMN_TYPE is a numeric data type, this is the maximum precision of the column. The precision of columns with a data type of DBTYPE_DECIMAL or DBTYPE_NUMERIC depends on the definition of the column.
													// For the precision of all other numeric data types, see http://msdn.microsoft.com/en-us/library/windows/desktop/ms715867(v=vs.85).aspx. If DBCOLUMN_TYPE is not a numeric data type, this is NULL.

	UINT16				m_DBCOLUMN_SCALE;			// If DBCOLUMN_TYPE is DBTYPE_DECIMAL or DBTYPE_NUMERIC, this is the number of digits to the right of the decimal point. Otherwise, this is NULL.			   
	DWORD			   	m_DBCOLUMN_FLAGS;

	// Optional columns

	WCHAR			   	m_DBCOLUMN_BASECATALOGNAME[128];	   
	WCHAR			   	m_DBCOLUMN_BASECOLUMNNAME[128];	   
	WCHAR			   	m_DBCOLUMN_BASESCHEMANAME[128];	   
	WCHAR			   	m_DBCOLUMN_BASETABLENAME[128];	   

	GUID		   		m_DBCOLUMN_CLSID;			   
	INT32				m_DBCOLUMN_COLLATINGSEQUENCE;   
	INT32				m_DBCOLUMN_COMPUTEMODE;		   
	UINT32			   	m_DBCOLUMN_DATETIMEPRECISION;   

	VARIANT_BOOL		m_DBCOLUMN_HASDEFAULT;		     
	VARIANT				m_DBCOLUMN_DEFAULTVALUE;		   

	VARIANT_BOOL		m_DBCOLUMN_ISAUTOINCREMENT;	     
	VARIANT_BOOL		m_DBCOLUMN_ISCASESENSITIVE;	     
	UINT32			   	m_DBCOLUMN_ISSEARCHABLE;		   
	VARIANT_BOOL		m_DBCOLUMN_ISUNIQUE;

	VARIANT_BOOL		m_DBCOLUMN_MAYSORT;			     	
	ULONG	    		m_DBCOLUMN_OCTETLENGTH;		   
	VARIANT_BOOL		m_DBCOLUMN_KEYCOLUMN;	
	UINT64			   	m_DBCOLUMN_BASETABLEVERSION;

	// Additional SQLite Specific Columns

	VARIANT_BOOL		m_DBCOLUMN_SQLITE_FKEYCOLUMN;	
	WCHAR			   	m_DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME[128];		   
	WCHAR			   	m_DBCOLUMN_SQLITE_FKEYCOLUMN_NAME[128];
	UINT32				m_DBCOLUMN_SQLITE_DATATYPE;
	WCHAR			   	m_DBCOLUMN_SQLITE_DECLATED_DATATYPE[128];		   
	VARIANT_BOOL		m_DBCOLUMN_SQLITE_CAN_BE_NULL;

	/////////////////////////////////////////////////////////////////////////////
	CSQLiteColumnsRowsetRow()
	{
		m_DBCOLUMN_GUID								= GUID_NULL;
		m_DBCOLUMN_CLSID							= GUID_NULL;
		m_DBCOLUMN_TYPEINFO							= NULL;

		m_DBCOLUMN_HASDEFAULT						= ATL_VARIANT_FALSE;		     
		m_DBCOLUMN_ISAUTOINCREMENT	     			= ATL_VARIANT_FALSE;
		m_DBCOLUMN_ISCASESENSITIVE	     			= ATL_VARIANT_FALSE;
		m_DBCOLUMN_ISUNIQUE			     			= ATL_VARIANT_FALSE;
		m_DBCOLUMN_KEYCOLUMN		     			= ATL_VARIANT_FALSE;
		m_DBCOLUMN_MAYSORT			     			= ATL_VARIANT_FALSE;
		m_DBCOLUMN_SQLITE_FKEYCOLUMN				= ATL_VARIANT_FALSE;
		m_DBCOLUMN_SQLITE_CAN_BE_NULL				= ATL_VARIANT_FALSE;

		m_DBCOLUMN_COLLATINGSEQUENCE				= 0; 
		m_DBCOLUMN_COMPUTEMODE		   				= 0;
		m_DBCOLUMN_PRECISION		   				= 0;
		m_DBCOLUMN_TYPE				   				= DBTYPEENUM::DBTYPE_EMPTY;
		m_DBCOLUMN_SCALE			   				= 0;
		m_DBCOLUMN_DATETIMEPRECISION   				= 0;
		m_DBCOLUMN_FLAGS			   				= 0;
		m_DBCOLUMN_ISSEARCHABLE		   				= 0;
		m_DBCOLUMN_PROPID			   				= 0;
		m_DBCOLUMN_BASETABLEVERSION	   				= 0;
		m_DBCOLUMN_COLUMNSIZE		   				= 0;
		m_DBCOLUMN_OCTETLENGTH		   				= 0;
		m_DBCOLUMN_NUMBER			   				= 0;
		m_DBCOLUMN_SQLITE_DATATYPE					= 0;

		m_DBCOLUMN_BASECATALOGNAME[0]				= NULL;
		m_DBCOLUMN_BASECOLUMNNAME[0]				= NULL;	   
		m_DBCOLUMN_BASESCHEMANAME[0]				= NULL;	   
		m_DBCOLUMN_BASETABLENAME[0]					= NULL;	   
		m_DBCOLUMN_IDNAME[0]						= NULL;			   
		m_DBCOLUMN_NAME[0]							= NULL;	
		m_DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME[0]	= NULL;
		m_DBCOLUMN_SQLITE_FKEYCOLUMN_NAME[0]		= NULL;
		m_DBCOLUMN_SQLITE_DECLATED_DATATYPE[0]		= NULL;

		::VariantInit(&m_DBCOLUMN_DEFAULTVALUE);
	}

	/////////////////////////////////////////////////////////////////////////////
	~CSQLiteColumnsRowsetRow()
	{
		VariantClear(&m_DBCOLUMN_DEFAULTVALUE);
	}

	/////////////////////////////////////////////////////////////////////////////
	bool InitBookmarkColumnMetadata(ATLCOLUMNINFO* m_ColumnsMetadata, /*out*/ DBBYTEOFFSET& Offset)
	{		
		wsprintf(m_DBCOLUMN_IDNAME, L"%s", L"Bookmark");		

		//BUG: Bookmarks should be DBTYPE_UI4 but oledb.h uses them as DBTYPE_BYTES
		//     This obviously allows 256 records per Recordset. Should fix this.
		
		m_DBCOLUMN_NUMBER						= 0;
		m_DBCOLUMN_TYPE							= DBTYPE_BYTES;
		m_DBCOLUMN_ISSEARCHABLE					= DB_UNSEARCHABLE;
		m_DBCOLUMN_FLAGS						= DBCOLUMNFLAGS_ISBOOKMARK|DBCOLUMNFLAGS_ISFIXEDLENGTH;
		m_DBCOLUMN_PRECISION					= 10;
		m_DBCOLUMN_SCALE						= 0;
		m_DBCOLUMN_PROPID						= 2;
		m_DBCOLUMN_COLUMNSIZE					= sizeof(DWORD);
		m_DBCOLUMN_ISCASESENSITIVE				= VARIANT_FALSE;		
		m_DBCOLUMN_HASDEFAULT					= VARIANT_FALSE;
		m_DBCOLUMN_ISUNIQUE						= VARIANT_FALSE;
		m_DBCOLUMN_DATETIMEPRECISION			= ~0;
		m_DBCOLUMN_OCTETLENGTH					= 0;
		m_DBCOLUMN_MAYSORT						= VARIANT_TRUE;

		memcpy(&m_DBCOLUMN_GUID, &DBCOL_SPECIALCOL, sizeof(GUID));
		
		ATLCOLUMNINFO& COLUMN_INFO				= m_ColumnsMetadata[0];
		COLUMN_INFO.cbOffset					= 0;
		COLUMN_INFO.iOrdinal					= 0;
		COLUMN_INFO.columnid.eKind				= DBKIND_NAME;
		COLUMN_INFO.columnid.uGuid.guid			= GUID_NULL;
		COLUMN_INFO.columnid.uName.pwszName		= OLESTR("Bookmark");
		COLUMN_INFO.pwszName					= OLESTR("Bookmark");
		COLUMN_INFO.wType						= m_DBCOLUMN_TYPE;
		COLUMN_INFO.ulColumnSize				= m_DBCOLUMN_COLUMNSIZE;
		COLUMN_INFO.dwFlags						= m_DBCOLUMN_FLAGS;				
		COLUMN_INFO.pTypeInfo					= (ITypeInfo*)NULL;

		Offset+=m_DBCOLUMN_COLUMNSIZE;
		return true;
	}

	/////////////////////////////////////////////////////////////////////////////
	bool InitColumnMetadata(sqlite3* db, std::string UTF8_TableName, std::string UTF8_ColumnName, ULONG Ordinal, ATLCOLUMNINFO& COLUMN_INFO, /*out*/ DBBYTEOFFSET& Offset)
	{
		int rc;
		sqlite3_stmt* stmt;

		NO_BRACKETS(UTF8_TableName);
		NO_BRACKETS(UTF8_ColumnName);

		// Set the Ordinal
		m_DBCOLUMN_NUMBER				= Ordinal;
		m_DBCOLUMN_GUID.Data1			= Ordinal;
		m_DBCOLUMN_TYPEINFO				= (ITypeInfo*) NULL;					
		m_DBCOLUMN_FLAGS				= Ordinal>0 ? DBCOLUMNFLAGS_WRITE : 0;
		m_DBCOLUMN_MAYSORT				= VARIANT_TRUE;
		m_DBCOLUMN_ISSEARCHABLE			= DB_ALL_EXCEPT_LIKE;
		m_DBCOLUMN_COMPUTEMODE			= 0;//DBCOMPUTEDMODE_NOTCOMPUTED;
		m_DBCOLUMN_ISCASESENSITIVE		= VARIANT_FALSE;
		m_DBCOLUMN_COLLATINGSEQUENCE	= GetSystemDefaultLCID();
		m_DBCOLUMN_BASETABLEVERSION		= 1;
		m_DBCOLUMN_SCALE				= NULL;

		// Create an SQLite Statement selecting only the column we are interested in.
		// NOTE: Column and Table in std::string are UTF8 encoded for save some string convertions.		
		std::string SQL("select [" + UTF8_ColumnName + "] from [" + UTF8_TableName + "] where 1=0;");
		rc = sqlite3_prepare_v2(db, SQL.c_str(), -1, &stmt, nullptr);
		if(rc!=SQLITE_OK) return false;

		// Get the query name of the column (eg. SELECT [a] AS [b], this holds [b])
		std::string ColumnName = UTF8_to_STR(sqlite3_column_name(stmt,0));
		BSTR sColumnName = STR_to_BSTR(ColumnName); 		 
		wsprintf(m_DBCOLUMN_NAME, L"%s", sColumnName);		

		// Get the base name of the column (eg. SELECT [a] AS [b], this holds [a])
		BASECOLUMNNAME = UTF8_to_STR(sqlite3_column_origin_name(stmt,0));
		BASECOLUMNNAME = "[" + BASECOLUMNNAME + "]";
		BSTR sBaseColumnName = STR_to_BSTR(BASECOLUMNNAME); 
		wsprintf(m_DBCOLUMN_BASECOLUMNNAME, L"%s", sBaseColumnName);		

		// Get the base table name of the column		
		BASETABLENAME = UTF8_to_STR(sqlite3_column_table_name(stmt,0));
		BASETABLENAME = "[" + BASETABLENAME + "]";
		BSTR sBaseTableName = STR_to_BSTR(BASETABLENAME); 
		wsprintf(m_DBCOLUMN_BASETABLENAME, L"%s", sBaseTableName);
		
		// Create a unique Column Identifier (GUID)		
		FILED_NAME = BASETABLENAME + "." + BASECOLUMNNAME;
		m_DBCOLUMN_GUID = MD5Guid(FILED_NAME);	
		BSTR sColumnQualifiedName = STR_to_BSTR(FILED_NAME); 		
		wsprintf(m_DBCOLUMN_IDNAME, L"%s", sColumnQualifiedName);		
			
		// Get Data Type
		const char* dt = sqlite3_column_decltype(stmt,0);
		if(dt==NULL) dt = "INTEGER";
		std::string DeclaredDatatype = UTF8_to_STR(dt);
		BSTR sDeclaredDatatype = STR_to_BSTR(DeclaredDatatype); 
		wsprintf(m_DBCOLUMN_SQLITE_DECLATED_DATATYPE, L"%s", sDeclaredDatatype);		

		// Get Column Constrains
		const char* pzDataType; const char* pzCollSeq; int pNotNull; int pPrimaryKey; int pAutoinc;		
		rc = sqlite3_table_column_metadata(db, NULL, sqlite3_column_table_name(stmt,0), sqlite3_column_origin_name(stmt,0), &pzDataType, &pzCollSeq, &pNotNull, &pPrimaryKey, &pAutoinc);		
		if(rc!=SQLITE_OK) return false;

		// Free SQLite statement
		rc = sqlite3_reset(stmt);
		rc = sqlite3_finalize(stmt);

		// Is it Null able?
		m_DBCOLUMN_SQLITE_CAN_BE_NULL = (pNotNull!=0 ? VARIANT_FALSE : VARIANT_TRUE);
		if(m_DBCOLUMN_SQLITE_CAN_BE_NULL)
		{
			m_DBCOLUMN_FLAGS |= DBCOLUMNFLAGS_ISNULLABLE | DBCOLUMNFLAGS_MAYBENULL;
		}

		// Is it Key Column?
		m_DBCOLUMN_KEYCOLUMN = (pPrimaryKey!=0 ? VARIANT_TRUE : VARIANT_FALSE);
		if(m_DBCOLUMN_KEYCOLUMN)
		{			
			m_DBCOLUMN_FLAGS |= DBCOLUMNFLAGS_ISROWID;
			m_DBCOLUMN_FLAGS |= DBCOLUMNFLAGS_KEYCOLUMN;
			m_DBCOLUMN_ISUNIQUE = VARIANT_TRUE;
			m_DBCOLUMN_SQLITE_CAN_BE_NULL = VARIANT_FALSE;
			pNotNull=0;
		}
		else
		{
			m_DBCOLUMN_FLAGS |= DBCOLUMNFLAGS_MAYBENULL;			
		}

		// Is it AutoIncrement?
		m_DBCOLUMN_ISAUTOINCREMENT = (pAutoinc!=0 ? VARIANT_TRUE : VARIANT_FALSE);
		if(m_DBCOLUMN_ISAUTOINCREMENT)
		{
			//TODO: Additional information about auto sequence?
		}

		// Column Defaults
		if(false)
		{
			m_DBCOLUMN_HASDEFAULT = VARIANT_TRUE;
			m_DBCOLUMN_DEFAULTVALUE.intVal = 100;
		}

		// Get Extended Datatype Information
		SQLiteColumnMetadata(DeclaredDatatype);
		
		// Parse CREATE TABLE SQL for Foreign Key Information
		// NOTE: There seems to be a bug in std::regexp and it does not parse the SQL.
		// I had to fall-back to Microsoft Scripting Engine which might be memory leaky.
		SQL = "SELECT [sql] FROM [sqlite_master] WHERE [tbl_name]='" + UTF8_TableName + "'";
		rc = sqlite3_prepare_v2(db, SQL.c_str(), -1, &stmt, nullptr);
		if(rc!=SQLITE_OK) return false;		
		for(;;)
		{
			rc = sqlite3_step(stmt);			
			char* csql = (char*)sqlite3_column_text(stmt, 0);
			if(csql!=NULL)
			{
				std::string create_sql(csql);
				std::string pattern(",\\s*(\\w+|(?:\\[[\\w\\s]+\\]))\\s*[^,]*?\\s+REFERENCES\\s+(\\w+|(?:\\[[\\w\\s]+\\]))\\s*\\(\\s*(\\w+|(?:\\[[\\w\\s]+\\]))\\s*\\)");
				ULONG matches = m_RegExp.Parse(create_sql, pattern);
				for(ULONG i=0;i<matches;i++)
				{
					std::string col = m_RegExp.SubMatch(i, 0);
					std::string ft = m_RegExp.SubMatch(i, 1);
					std::string fk = m_RegExp.SubMatch(i, 2);
					NO_BRACKETS(col);
					NO_BRACKETS(ft);
					NO_BRACKETS(fk);

					if(col==UTF8_ColumnName)
					{
						//char debug[1024];
						//sprintf(debug, "[%s].[%s] -> [%s].[%s]\n", _TableName.c_str(), col.c_str(), ft.c_str(), fk.c_str());
						//OutputDebugStringA(debug);						

						BSTR FT = UTF8_to_BSTR(ft.c_str());
						BSTR FK = UTF8_to_BSTR(fk.c_str());
						m_DBCOLUMN_SQLITE_FKEYCOLUMN = VARIANT_TRUE;
						wsprintf(m_DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME, L"%s",FT);
						wsprintf(m_DBCOLUMN_SQLITE_FKEYCOLUMN_NAME, L"%s", FK);
						SysFreeString(FT);
						SysFreeString(FK);
					}
				}
			}
			if(rc==SQLITE_DONE)
				break;
		}
		rc = sqlite3_finalize(stmt);

		// Set COLUMN_INFO attributes
		SET_COLUMN_INFO(COLUMN_INFO, Offset, sColumnName);
			
		// Clean up
		SysFreeString(sColumnName);
		SysFreeString(sBaseTableName);
		SysFreeString(sBaseColumnName);
		SysFreeString(sColumnQualifiedName);

		// Increase Offset and return it
		Offset += m_DBCOLUMN_COLUMNSIZE;		
		return true;
	}

	/////////////////////////////////////////////////////////////////////////////
	bool InitExpressionColumnMetadata(sqlite3* db, sqlite3_stmt* stmt, ULONG Ordinal, ATLCOLUMNINFO& COLUMN_INFO, /*out*/ DBBYTEOFFSET& Offset)
	{
		// Set the Ordinal
		m_DBCOLUMN_NUMBER				= Ordinal;
		m_DBCOLUMN_GUID.Data1			= Ordinal;
		m_DBCOLUMN_TYPEINFO				= (ITypeInfo*) NULL;					
		m_DBCOLUMN_FLAGS				= Ordinal>0 ? DBCOLUMNFLAGS_WRITE : 0;
		m_DBCOLUMN_MAYSORT				= VARIANT_TRUE;
		m_DBCOLUMN_ISSEARCHABLE			= DB_ALL_EXCEPT_LIKE;
		m_DBCOLUMN_COMPUTEMODE			= DBCOMPUTEMODE_COMPUTED;
		m_DBCOLUMN_ISCASESENSITIVE		= VARIANT_FALSE;
		m_DBCOLUMN_COLLATINGSEQUENCE	= GetSystemDefaultLCID();
		m_DBCOLUMN_BASETABLEVERSION		= 1;
		m_DBCOLUMN_SCALE				= NULL;

		char cOrdinal[10];
		sprintf(cOrdinal, "%d", Ordinal);
		std::string sOrdinal(cOrdinal);

		// Get the query name of the column (eg. SELECT [a] AS [b], this holds [b])
		std::string ColumnName("Expr" + sOrdinal);
		BSTR sColumnName = UTF8_to_BSTR(ColumnName.c_str()); 		 
		wsprintf(m_DBCOLUMN_NAME, L"%s", sColumnName);		

		// Get the base name of the column (eg. SELECT [a] AS [b], this holds [a])
		BASECOLUMNNAME = "EXPR" + sOrdinal;
		BSTR sBaseColumnName = UTF8_to_BSTR(BASECOLUMNNAME.c_str()); 
		wsprintf(m_DBCOLUMN_BASECOLUMNNAME, L"%s", sBaseColumnName);		

		// Get the base table name of the column		
		BASETABLENAME = "SQLiteExpression";
		BSTR sBaseTableName = UTF8_to_BSTR(BASETABLENAME.c_str()); 
		wsprintf(m_DBCOLUMN_BASETABLENAME, L"%s", sBaseTableName);

		// Create a unique Column Identifier (GUID)		
		FILED_NAME = BASETABLENAME + "." + BASECOLUMNNAME;
		m_DBCOLUMN_GUID = MD5Guid(FILED_NAME);	
		BSTR sColumnQualifiedName = UTF8_to_BSTR(FILED_NAME.c_str()); 		
		wsprintf(m_DBCOLUMN_IDNAME, L"%s", sColumnQualifiedName);		

		// Get Data Type
		m_DBCOLUMN_SQLITE_DATATYPE = sqlite3_column_type(stmt, 0);
		std::string DeclaredDatatype("TEXT");
		switch(m_DBCOLUMN_SQLITE_DATATYPE)
		{
		case SQLITE_INTEGER: DeclaredDatatype = "INTEGER"; break;
		case SQLITE_FLOAT:	 DeclaredDatatype = "FLOAT"; break;
		case SQLITE_BLOB:	 DeclaredDatatype = "BLOB"; break;
		}
		BSTR sDeclaredDatatype = UTF8_to_BSTR(DeclaredDatatype.c_str()); 
		wsprintf(m_DBCOLUMN_SQLITE_DECLATED_DATATYPE, L"%s", sDeclaredDatatype);		

		// Get Extended Datatype Information
		SQLiteColumnMetadata(DeclaredDatatype);

		// Set COLUMN_INFO attributes
		SET_COLUMN_INFO(COLUMN_INFO, Offset, sColumnName);

		// Clean up
		SysFreeString(sColumnName);
		SysFreeString(sBaseTableName);
		SysFreeString(sBaseColumnName);
		SysFreeString(sColumnQualifiedName);

		// Increase Offset and return it
		Offset += m_DBCOLUMN_COLUMNSIZE;		
		return true;
	}

	/////////////////////////////////////////////////////////////////////////////
	void SET_COLUMN_INFO(ATLCOLUMNINFO& COLUMN_INFO, DBBYTEOFFSET Offset, BSTR sColumnName)
	{																																	
		COLUMN_INFO.iOrdinal					= m_DBCOLUMN_NUMBER;																					
		COLUMN_INFO.cbOffset					= Offset;																					
		COLUMN_INFO.dwFlags						= m_DBCOLUMN_FLAGS;																			
		COLUMN_INFO.pTypeInfo					= (ITypeInfo*)NULL;																			
		COLUMN_INFO.wType						= m_DBCOLUMN_TYPE;																			
		COLUMN_INFO.ulColumnSize				= m_DBCOLUMN_COLUMNSIZE	== (ULONG)~0  ? (DBLENGTH)~0	: m_DBCOLUMN_COLUMNSIZE;		
		COLUMN_INFO.bPrecision					= m_DBCOLUMN_PRECISION	== (UINT16)~0 ? -1		: m_DBCOLUMN_PRECISION;			
		COLUMN_INFO.bScale						= m_DBCOLUMN_SCALE		== (UINT16)~0 ? -1		: m_DBCOLUMN_SCALE;				

		memset(&(COLUMN_INFO.columnid), 0, sizeof(DBID));	

		COLUMN_INFO.columnid.eKind				= DBKIND_GUID_NAME;																			
		COLUMN_INFO.columnid.uGuid.guid			= m_DBCOLUMN_GUID;																			
		COLUMN_INFO.columnid.uName.pwszName		= _wcsdup(m_DBCOLUMN_IDNAME);																
		COLUMN_INFO.pwszName					= _wcsdup(sColumnName);																		

		char colinfo[512];
		sprintf(colinfo, "Column: %s type: %s\n", FILED_NAME.c_str(), OLEDB_TYPENAME(m_DBCOLUMN_TYPE).c_str());
		OutputDebugStringA(colinfo);
	}

	/////////////////////////////////////////////////////////////////////////////
	// This function was a hell to get right and I have tested it with dxDBGrid
	// using all shorts of datatypes in test.db. I have made excessive tests 
	// using both Server side cursor Recordsets and Client side cursor Recordsets
	// and especially for dxDBGrid I tested Lookup Columns. It seems that using
	// BSTR for strings, even though it makes life and memory management easier
	// it is not compatible with the Lookup mechanism in dxDBGrid and I had to 
	// fall-back to using WSTR. Using non-fixed-length BSTR in disconnected 
	// Recordsets forces Microsoft Client Cursor Engine to use ISequentialSteam
	// on the consumer side which also makes the code a mess when using the 
	// provider with dxDBGrid. Anyway, I got everything fixed and works!!
	/////////////////////////////////////////////////////////////////////////////
	bool SQLiteColumnMetadata(std::string SQLiteDeclaredType)
	{
		if(SQLiteDeclaredType.size()==0)
			return DBTYPEENUM::DBTYPE_EMPTY;	

		// -------------------------------------------------------------------------------------------------------------------------
		// Extract precision and scale (eg. DECIMAL(19,2) )
		// -------------------------------------------------------------------------------------------------------------------------

		std::string DECLARED_TYPE;
		ULONG DECLARED_SIZE = 0;
		UINT16 DECLARED_PRECISION = 0;
		UINT16 DECLARED_SCALE = 0;

		// We will use a regular expression to extract precision (or size) and scale from the SQLite declared datatype.
		std::transform(SQLiteDeclaredType.begin(), SQLiteDeclaredType.end(), SQLiteDeclaredType.begin(), ::toupper);
		const std::regex pattern("^\\s*(\\w+)(?:\\s*\\(\\s*(\\d+)\\s*(?:,\\s*(\\d+)\\s*)?\\)\\s*)?$");		
		const std::sregex_iterator end;
		for(std::sregex_iterator m(SQLiteDeclaredType.cbegin(), SQLiteDeclaredType.cend(), pattern); m!=end; ++m)
		{	
			DECLARED_TYPE = ((*m)[1]);		
			std::string precision((*m)[2]);
			std::string scale((*m)[3]);
			DECLARED_SIZE = DECLARED_PRECISION = (precision!="" ? atoi(precision.c_str()) : 0);
			DECLARED_SCALE = scale!="" ? atoi(scale.c_str()) : 0;
		}

		// -------------------------------------------------------------------------------------------------------------------------
		// Map SQLite declared type to DBTYPEENUM
		// -------------------------------------------------------------------------------------------------------------------------
		//
		// Special cases that either have their own UDT or need additional flags and specification:
		//
		//	DBTYPE_DECIMAL			DBDECIMAL
		//	DBTYPE_NUMERIC			DBNUMERIC				Implementation falls back to DECIMAL
		//	DBTYPE_DBDATE			DBDATE					Persisted as ISO8601 string
		//	DBTYPE_DBTIME			DBTIME					Persisted as ISO8601 string
		//	DBTYPE_TIMESTAMP		DBTIMESTAMP				Persisted as ISO8601 string
		//	DBTYPE_DATE				DATE					Persisted as ISO8601 string
		//	DBTYPE_BOOL				VARIANT_BOOL			Note: size is 2 bytes
		//	DBTYPE_BSTR				BSTR
		//	DBTYPE_STR				char[cbMaxLen]
		//	DBTYPE_WSTR				wchar_t[cbMaxLen]
		//	DBTYPE_VARIANT			VARIANT					Not implemented 
		//	DBTYPE_ARRAY			SAFEARRAY *				Not implemented
		//	DBTYPE_VECTOR			DBVECTOR				Not implemented
		//	DBTYPE_BYTES			BYTE[cbMaxlen]			Not implemented (BLOBs)

			 if(DECLARED_TYPE == "INTEGER"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I4;					}
		else if(DECLARED_TYPE == "INT1"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I1;					}
		else if(DECLARED_TYPE == "INT2"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I2;					}
		else if(DECLARED_TYPE == "INT4"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I4;					}
		else if(DECLARED_TYPE == "INT8"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I8;					}
		else if(DECLARED_TYPE == "INT"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I4;					}
		else if(DECLARED_TYPE == "TINYINT"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I1;					}
		else if(DECLARED_TYPE == "SMALLINT"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I2;					}
		else if(DECLARED_TYPE == "MEDIUMINT"		)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I4;					}
		else if(DECLARED_TYPE == "BIGINT"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I8;					}
		else if(DECLARED_TYPE == "LONG"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_I8;					}

		else if(DECLARED_TYPE == "UNSIGNED BIG INT"	)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_UI8;					}		
		else if(DECLARED_TYPE == "UINT1"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_UI1;					}		
		else if(DECLARED_TYPE == "UINT2"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_UI2;					}		
		else if(DECLARED_TYPE == "UINT4"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_UI8;					} //BUG: UINT4 does not work with dxDBGrid
		else if(DECLARED_TYPE == "UINT8"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_UI8;					}		

		else if(DECLARED_TYPE == "REAL"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_R8;					}
		else if(DECLARED_TYPE == "DOUBLE PRECISION"	)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_R8;					}
		else if(DECLARED_TYPE == "DOUBLE"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_R8;					}
		else if(DECLARED_TYPE == "FLOAT"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_R8;					}

		else if(DECLARED_TYPE == "BOOLEAN"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_BOOL;				}
		else if(DECLARED_TYPE == "BOOL"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_BOOL;				}

		else if(DECLARED_TYPE == "SMALLDATETIME"	)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DBDATE;				}
		else if(DECLARED_TYPE == "DATETIME"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DBTIMESTAMP;			}
		else if(DECLARED_TYPE == "DATE"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DBDATE;				}
		else if(DECLARED_TYPE == "TIMESTAMP"		)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DBTIMESTAMP;			} // or DBTYPE_BYTES with DBCOLUMNFLAGS_ISROWVER set
		else if(DECLARED_TYPE == "TIME"				)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DBTIME;				}

		else if(DECLARED_TYPE == "SMALLMONEY"		)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_CY; 					}
		else if(DECLARED_TYPE == "MONEY"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_CY; 					}
		else if(DECLARED_TYPE == "CURRENCY"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_CY; 					}

		else if(DECLARED_TYPE == "IMAGE"			)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}
		else if(DECLARED_TYPE == "BLOB"				)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}
		else if(DECLARED_TYPE == "BINARY"			)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}
		else if(DECLARED_TYPE == "VARBINARY"		)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}
		else if(DECLARED_TYPE == "BYTES"			)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}
		else if(DECLARED_TYPE == "VARBYTES"			)	{ m_DBCOLUMN_TYPE = DBTYPE_BLOB;								}

		else if(DECLARED_TYPE == "DECIMAL"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DECIMAL; 			}
		else if(DECLARED_TYPE == "NUMERIC"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DECIMAL; 			} // Numeric needs some work
		else if(DECLARED_TYPE == "NUMBER"			)	{ m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_DECIMAL; 			}

		else
		{
			// All unrecognized declared types fall back to string.
			if(DECLARED_SIZE>255 || DECLARED_SIZE==0)
				m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_BSTR;
			else
				m_DBCOLUMN_TYPE = DBTYPEENUM::DBTYPE_WSTR;
		}

		// -------------------------------------------------------------------------------------------------------------------------
		// Set fixed size
		// -------------------------------------------------------------------------------------------------------------------------

		if(IsFixedType((DBTYPE)m_DBCOLUMN_TYPE))
			m_DBCOLUMN_FLAGS |= DBCOLUMNFLAGS_ISFIXEDLENGTH;

		// -------------------------------------------------------------------------------------------------------------------------
		// Set Datatype Size, Scale, Precision
		// -------------------------------------------------------------------------------------------------------------------------

		// Initialize with defaults:

		m_DBCOLUMN_COLUMNSIZE			= 0;
		m_DBCOLUMN_OCTETLENGTH			= NULL;				// NULL for non-character and non-binary types.
		m_DBCOLUMN_PRECISION			= NULL;				// If DBCOLUMN_TYPE is not a numeric data type, this is NULL.
		m_DBCOLUMN_SCALE				= NULL;				// If DBCOLUMN_TYPE is not a numeric data type, this is NULL.
		m_DBCOLUMN_DATETIMEPRECISION	= (UINT32) ~0;
		m_DBCOLUMN_ISSEARCHABLE			= DB_SEARCHABLE;

		// Here is all the information required for understanding the following code:
		// http://msdn.microsoft.com/en-us/library/windows/desktop/ms714373(v=vs.85).aspx
		// http://msdn.microsoft.com/en-us/library/windows/desktop/ms715968(v=vs.85).aspx
		// http://msdn.microsoft.com/en-us/library/windows/desktop/ms723069(v=vs.85).aspx
		// http://msdn.microsoft.com/en-us/library/windows/desktop/ms712925(v=vs.85).aspx

		// DBCOLUMN_COLUMNSIZE:
		// ====================
		// The maximum possible length of a value in the column. 
		// For columns that use a fixed-length data type, this is the size of the data type. 
		// For columns that use a variable-length data type, this is one of the following:
		// - The maximum length of the column in characters (for DBTYPE_STR and DBTYPE_WSTR) 
		//	 or in bytes (for DBTYPE_BYTES and DBTYPE_VARNUMERIC), if one is defined. 
		//	 For example, a CHAR(5) column in an SQL table has a maximum length of 5.
		// - The maximum length of the data type in characters (for DBTYPE_STR and DBTYPE_WSTR)
		//   or in bytes (for DBTYPE_BYTES and DBTYPE_VARNUMERIC), if the column does not have a defined length.
		// - ~0 if neither the column nor the data type has a defined maximum length.

		// DBCOLUMN_OCTETLENGTH:
		// =====================
		// If the column is a character or binary type:
		// - The maximum length in octets (bytes) of the column
		// - Zero if column has no maximum length
		// NULL for all other types of columns.

		// If you set DBCOLUMNFLAGS_ISLONG flag you specify that this provider supports treating BLOBs with ISequentialStream. 

		// Having said all that, here we go:

		switch(m_DBCOLUMN_TYPE)
		{

		// ======= NUMBERS =====================================================================

		case DBTYPEENUM::DBTYPE_I1:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(INT8);   m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION =  3; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_I2:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(INT16);  m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION =  5; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_I4:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(INT32);  m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION = 10; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_I8:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(INT64);  m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION = 19; m_DBCOLUMN_SCALE = NULL; break; }		
		case DBTYPEENUM::DBTYPE_UI1:	{ m_DBCOLUMN_COLUMNSIZE = sizeof(UINT8);  m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION =  3; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_UI2:	{ m_DBCOLUMN_COLUMNSIZE = sizeof(UINT16); m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION =  5; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_UI4:	{ m_DBCOLUMN_COLUMNSIZE = sizeof(UINT32); m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION = 10; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_UI8:	{ m_DBCOLUMN_COLUMNSIZE = sizeof(UINT64); m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION = 20; m_DBCOLUMN_SCALE = NULL; break; }		
		case DBTYPEENUM::DBTYPE_R4:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(FLOAT);  m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION =  7; m_DBCOLUMN_SCALE = NULL; break; }
		case DBTYPEENUM::DBTYPE_R8:		{ m_DBCOLUMN_COLUMNSIZE = sizeof(DOUBLE); m_DBCOLUMN_OCTETLENGTH = NULL; m_DBCOLUMN_PRECISION = 16; m_DBCOLUMN_SCALE = NULL; break; }

		case DBTYPEENUM::DBTYPE_CY:		
			m_DBCOLUMN_COLUMNSIZE			= sizeof(CY);
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 19;
			m_DBCOLUMN_SCALE				= DECLARED_SCALE ? DECLARED_SCALE : 8;
			break;

		case DBTYPEENUM::DBTYPE_DECIMAL:						
			m_DBCOLUMN_COLUMNSIZE			= sizeof(DECIMAL);
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 28;
			m_DBCOLUMN_SCALE				= DECLARED_SCALE ? DECLARED_SCALE : 4;
			break;

		case DBTYPEENUM::DBTYPE_NUMERIC:						
			m_DBCOLUMN_COLUMNSIZE			= sizeof(DB_NUMERIC);
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 38;
			m_DBCOLUMN_SCALE				= DECLARED_SCALE ? DECLARED_SCALE : 4;
			break;

		case DBTYPEENUM::DBTYPE_BOOL:	
			m_DBCOLUMN_COLUMNSIZE			= sizeof(VARIANT_BOOL);
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= NULL;
			m_DBCOLUMN_SCALE				= NULL;
			break; 			

		// ======= DATES =====================================================================

		case DBTYPEENUM::DBTYPE_DBDATE:
			m_DBCOLUMN_COLUMNSIZE			= sizeof(DBDATE);				
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 15;
			m_DBCOLUMN_SCALE				= NULL;
			break;

		case DBTYPEENUM::DBTYPE_DBTIME:
			m_DBCOLUMN_COLUMNSIZE			= sizeof(DBTIME);				
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 8;
			m_DBCOLUMN_SCALE				= NULL;
			break;

		case DBTYPEENUM::DBTYPE_DBTIMESTAMP:
			m_DBCOLUMN_COLUMNSIZE			= sizeof(DBTIMESTAMP);				
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= 15;
			m_DBCOLUMN_SCALE				= NULL;
			break;

		// ======= BLOBs =====================================================================
		// Blobs are treated as VARIANT(ARRAY|UI1)
		case DBTYPE_BLOB:
			m_DBCOLUMN_COLUMNSIZE			= sizeof(VARIANT);
			m_DBCOLUMN_OCTETLENGTH			= NULL;
			m_DBCOLUMN_PRECISION			= NULL;
			m_DBCOLUMN_SCALE				= NULL;
			m_DBCOLUMN_FLAGS				&= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
			m_DBCOLUMN_ISSEARCHABLE			= DB_UNSEARCHABLE;
			break;

		// ======= STRINGS =====================================================================

		// Long Strings
		case DBTYPEENUM::DBTYPE_BSTR:
			m_DBCOLUMN_COLUMNSIZE			= DECLARED_SIZE ? DECLARED_SIZE : ~0;
			m_DBCOLUMN_OCTETLENGTH			= DECLARED_SIZE ? DECLARED_SIZE : 0;
			m_DBCOLUMN_PRECISION			= NULL;
			m_DBCOLUMN_SCALE				= NULL;
			m_DBCOLUMN_FLAGS				&= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;
			m_DBCOLUMN_ISSEARCHABLE			= DB_ALL_EXCEPT_LIKE;			
			break;


		// Defined-size Strings < 255 characters
		case DBTYPEENUM::DBTYPE_WSTR:

			if(DECLARED_SIZE==0)
			{
				m_DBCOLUMN_COLUMNSIZE		= ~0;
				m_DBCOLUMN_OCTETLENGTH		= 0;			
				m_DBCOLUMN_FLAGS			&= ~DBCOLUMNFLAGS_ISFIXEDLENGTH;				
				m_DBCOLUMN_ISSEARCHABLE		= DB_ALL_EXCEPT_LIKE;
			}
			else
			{
				m_DBCOLUMN_COLUMNSIZE		= DECLARED_SIZE;
				m_DBCOLUMN_OCTETLENGTH		= DECLARED_SIZE;
				m_DBCOLUMN_FLAGS			|= DBCOLUMNFLAGS_ISFIXEDLENGTH;
			}		
			break;
		}

		// We need to properly set the generic SQLite type.
		switch(m_DBCOLUMN_TYPE)
		{
			case DBTYPEENUM::DBTYPE_I1:	
			case DBTYPEENUM::DBTYPE_I2:	
			case DBTYPEENUM::DBTYPE_I4:	
			case DBTYPEENUM::DBTYPE_I8:	
			case DBTYPEENUM::DBTYPE_UI1:
			case DBTYPEENUM::DBTYPE_UI2:
			case DBTYPEENUM::DBTYPE_UI4:
			case DBTYPEENUM::DBTYPE_UI8:
			case DBTYPEENUM::DBTYPE_BOOL:

				m_DBCOLUMN_SQLITE_DATATYPE = SQLITE_INTEGER;
				break;

			case DBTYPEENUM::DBTYPE_R4:	
			case DBTYPEENUM::DBTYPE_R8:	
			case DBTYPEENUM::DBTYPE_CY:		
			case DBTYPEENUM::DBTYPE_DECIMAL:						
			case DBTYPEENUM::DBTYPE_NUMERIC:						

				m_DBCOLUMN_SQLITE_DATATYPE = SQLITE_FLOAT;
				break;

			case DBTYPE_BLOB:

				m_DBCOLUMN_SQLITE_DATATYPE = SQLITE_BLOB;
				break;

			case DBTYPEENUM::DBTYPE_DBDATE:
			case DBTYPEENUM::DBTYPE_DBTIME:
			case DBTYPEENUM::DBTYPE_DBTIMESTAMP:
			case DBTYPEENUM::DBTYPE_BSTR:
			case DBTYPEENUM::DBTYPE_WSTR:

				m_DBCOLUMN_SQLITE_DATATYPE = SQLITE_TEXT;
				break;

			default:
				// Should NOT happen.
				char error[256];
				sprintf(error, "ERROR: Missing SQLite type mapping for %d (%s).\n", m_DBCOLUMN_TYPE, OLEDB_TYPENAME(m_DBCOLUMN_TYPE));
				OutputDebugStringA(error);
				break;
		}	

		return true;
	}

	/////////////////////////////////////////////////////////////////////////////

	//http://msdn.microsoft.com/en-us/library/windows/desktop/ms712925(v=vs.85).aspx

	BEGIN_PROVIDER_COLUMN_MAP(CSQLiteColumnsRowsetRow)

		//					Ordinal			Guid									Name										Size						Type					  Precision	      Scale		 Flags		Variable			
		
		// Mandatory Columns																																																						  

		SQLITE_COLUMN_ENTRY(	  1,		DBCOLUMN_IDNAME,						"DBCOLUMN_IDNAME",							128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_IDNAME								)
		SQLITE_COLUMN_ENTRY(	  2,		DBCOLUMN_GUID,							"DBCOLUMN_GUID",							sizeof(GUID),				DBTYPE_GUID,					255,		255,	   112,		m_DBCOLUMN_GUID									)
		SQLITE_COLUMN_ENTRY(	  3,		DBCOLUMN_PROPID,						"DBCOLUMN_PROPID",							sizeof(ULONG),				DBTYPE_UI4,						 10,		255,	   112,		m_DBCOLUMN_PROPID								)
		SQLITE_COLUMN_ENTRY(	  4,		DBCOLUMN_NAME,							"DBCOLUMN_NAME",							128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_NAME									)
		SQLITE_COLUMN_ENTRY(	  5,		DBCOLUMN_NUMBER,						"DBCOLUMN_NUMBER",							sizeof(ULONG_PTR),			DBTYPEFOR_DBORDINAL,	DBPRECISION,		255,		80,		m_DBCOLUMN_NUMBER								)
		SQLITE_COLUMN_ENTRY(	  6,		DBCOLUMN_TYPE,							"DBCOLUMN_TYPE",							sizeof(USHORT),				DBTYPE_UI2,						  5,		255,		80,		m_DBCOLUMN_TYPE									)
		SQLITE_COLUMN_ENTRY(	  7,		DBCOLUMN_TYPEINFO,						"DBCOLUMN_TYPEINFO",						sizeof(IUnknown),			DBTYPE_IUNKNOWN,				255,		255,	   112,		m_DBCOLUMN_TYPEINFO								)
		SQLITE_COLUMN_ENTRY(	  8,		DBCOLUMN_COLUMNSIZE,					"DBCOLUMN_COLUMNSIZE",						sizeof(ULONG_PTR),			DBTYPEFOR_DBLENGTH,		DBPRECISION,		255,		80,		m_DBCOLUMN_COLUMNSIZE							)
		SQLITE_COLUMN_ENTRY(	  9,		DBCOLUMN_PRECISION,						"DBCOLUMN_PRECISION",						sizeof(USHORT),				DBTYPE_UI2,						  5,		255,	   112,		m_DBCOLUMN_PRECISION							)
		SQLITE_COLUMN_ENTRY(	 10,		DBCOLUMN_SCALE,							"DBCOLUMN_SCALE",							sizeof(SHORT),				DBTYPE_I2,						  5,		255,	   112,		m_DBCOLUMN_SCALE								)
		SQLITE_COLUMN_ENTRY(	 11,		DBCOLUMN_FLAGS,							"DBCOLUMN_FLAGS",							sizeof(ULONG),				DBTYPE_UI4,						 10,		255,		80,		m_DBCOLUMN_FLAGS								)

		// Optional Columns																																															  		

		SQLITE_COLUMN_ENTRY(	 12,		DBCOLUMN_BASECATALOGNAME,				"DBCOLUMN_BASECATALOGNAME",					128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_BASECATALOGNAME						)
		SQLITE_COLUMN_ENTRY(	 13,		DBCOLUMN_BASECOLUMNNAME,				"DBCOLUMN_BASECOLUMNNAME",					128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_BASECOLUMNNAME						)
		SQLITE_COLUMN_ENTRY(	 14,		DBCOLUMN_BASESCHEMANAME,				"DBCOLUMN_BASESCHEMANAME",					128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_BASESCHEMANAME						)
		SQLITE_COLUMN_ENTRY(	 15,		DBCOLUMN_BASETABLENAME,					"DBCOLUMN_BASETABLENAME",					128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_BASETABLENAME						)
		SQLITE_COLUMN_ENTRY(	 16,		DBCOLUMN_CLSID,							"DBCOLUMN_CLSID",							sizeof(GUID),				DBTYPE_GUID,					255,		255,	   112,		m_DBCOLUMN_CLSID								)
		SQLITE_COLUMN_ENTRY(	 17,		DBCOLUMN_COLLATINGSEQUENCE,				"DBCOLUMN_COLLATINGSEQUENCE",				sizeof(LONG),				DBTYPE_I4,						 10,		255,	   112,		m_DBCOLUMN_COLLATINGSEQUENCE					)
		SQLITE_COLUMN_ENTRY(	 18,		DBCOLUMN_COMPUTEMODE,					"DBCOLUMN_COMPUTEMODE",						sizeof(LONG),				DBTYPE_I4,						 10,		255,	   112,		m_DBCOLUMN_COMPUTEMODE							)
		SQLITE_COLUMN_ENTRY(	 19,		DBCOLUMN_DATETIMEPRECISION,				"DBCOLUMN_DATETIMEPRECISION",				sizeof(LONG),				DBTYPE_UI4,						 10,		255,	   112,		m_DBCOLUMN_DATETIMEPRECISION					)
		SQLITE_COLUMN_ENTRY(	 20,		DBCOLUMN_DEFAULTVALUE,					"DBCOLUMN_DEFAULTVALUE",					sizeof(VARIANT),			DBTYPE_VARIANT,					255,		255,	   112,		m_DBCOLUMN_DEFAULTVALUE							)
		SQLITE_COLUMN_ENTRY(	 21,		DBCOLUMN_HASDEFAULT,					"DBCOLUMN_HASDEFAULT",						sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	   112,		m_DBCOLUMN_HASDEFAULT							)
		SQLITE_COLUMN_ENTRY(	 22,		DBCOLUMN_ISAUTOINCREMENT,				"DBCOLUMN_ISAUTOINCREMENT",					sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	    16,		m_DBCOLUMN_ISAUTOINCREMENT						)
		SQLITE_COLUMN_ENTRY(	 23,		DBCOLUMN_ISCASESENSITIVE,				"DBCOLUMN_ISCASESENSITIVE",					sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	   112,		m_DBCOLUMN_ISCASESENSITIVE						)
		SQLITE_COLUMN_ENTRY(	 24,		DBCOLUMN_ISSEARCHABLE,					"DBCOLUMN_ISSEARCHABLE",					sizeof(ULONG),				DBTYPE_UI4,						 10,		255,	   112,		m_DBCOLUMN_ISSEARCHABLE							)
		SQLITE_COLUMN_ENTRY(	 25,		DBCOLUMN_ISUNIQUE,						"DBCOLUMN_ISUNIQUE",						sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	   112,		m_DBCOLUMN_ISUNIQUE								)
		SQLITE_COLUMN_ENTRY(	 26,		DBCOLUMN_MAYSORT,						"DBCOLUMN_MAYSORT",							sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	    16,		m_DBCOLUMN_MAYSORT								)
		SQLITE_COLUMN_ENTRY(	 27,		DBCOLUMN_OCTETLENGTH,					"DBCOLUMN_OCTETLENGTH",						sizeof(ULONG_PTR),			DBTYPEFOR_DBLENGTH,		DBPRECISION,		255,	   112,		m_DBCOLUMN_OCTETLENGTH							)
		SQLITE_COLUMN_ENTRY(	 28,		DBCOLUMN_KEYCOLUMN,						"DBCOLUMN_KEYCOLUMN",						sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	    16,		m_DBCOLUMN_KEYCOLUMN							)
		SQLITE_COLUMN_ENTRY(	 29,		DBCOLUMN_BASETABLEVERSION,				"DBCOLUMN_BASETABLEVERSION",				sizeof(ULARGE_INTEGER),		DBTYPE_UI8,						 20,		255,	    16,		m_DBCOLUMN_BASETABLEVERSION						)

		// SQLite Specific Optional Columns

		SQLITE_COLUMN_ENTRY(	 30,		DBCOLUMN_SQLITE_FKEYCOLUMN,				"DBCOLUMN_SQLITE_FKEYCOLUMN",				sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	    16,		m_DBCOLUMN_SQLITE_FKEYCOLUMN					)
		SQLITE_COLUMN_ENTRY(	 31,		DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME,	"DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME",	128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME			)
		SQLITE_COLUMN_ENTRY(	 32,		DBCOLUMN_SQLITE_FKEYCOLUMN_NAME,		"DBCOLUMN_SQLITE_FKEYCOLUMN_NAME",			128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_SQLITE_FKEYCOLUMN_NAME				)

		SQLITE_COLUMN_ENTRY(	 33,		DBCOLUMN_SQLITE_DATATYPE,				"DBCOLUMN_SQLITE_DATATYPE",					sizeof(LONG),				DBTYPE_I4,						 10,		255,	   112,		m_DBCOLUMN_SQLITE_DATATYPE						)
		SQLITE_COLUMN_ENTRY(	 34,		DBCOLUMN_SQLITE_DECLATED_DATATYPE,		"DBCOLUMN_SQLITE_DECLATED_DATATYPE",		128,						DBTYPE_WSTR,					255,		255,	   112,		m_DBCOLUMN_SQLITE_DECLATED_DATATYPE				)
		SQLITE_COLUMN_ENTRY(	 35,		DBCOLUMN_SQLITE_CAN_BE_NULL,			"DBCOLUMN_SQLITE_CAN_BE_NULL",				sizeof(VARIANT_BOOL),		DBTYPE_BOOL,					255,		255,	    16,		m_DBCOLUMN_SQLITE_CAN_BE_NULL					)

	END_PROVIDER_COLUMN_MAP()

};

// ==================================================================================================================================
//	   ___________ ____    __    _ __       ______      __                           ____                          __ 
//	  / ____/ ___// __ \  / /   (_) /____  / ____/___  / /_  ______ ___  ____  _____/ __ \____ _      __________  / /_
//	 / /    \__ \/ / / / / /   / / __/ _ \/ /   / __ \/ / / / / __ `__ \/ __ \/ ___/ /_/ / __ \ | /| / / ___/ _ \/ __/
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/ /___/ /_/ / / /_/ / / / / / / / / (__  ) _, _/ /_/ / |/ |/ (__  )  __/ /_  
//	\____//____/\___\_\/_____/_/\__/\___/\____/\____/_/\__,_/_/ /_/ /_/_/ /_/____/_/ |_|\____/|__/|__/____/\___/\__/  
//	                                                                                                                  
// ==================================================================================================================================

/////////////////////////////////////////////////////////////////////////////
template <class CreatorClass>
class CSQLiteColumnsRowset : 
	public CSchemaRowsetImpl<CSQLiteColumnsRowset<CreatorClass>, CSQLiteColumnsRowsetRow, CreatorClass>	
{
public:
	
	typedef CSQLiteColumnsRowset<CreatorClass> _RowsetClass;

	bool m_bIsTable;
	DBORDINAL m_cColumns;
	ATLCOLUMNINFO *m_rgColumns;

	/////////////////////////////////////////////////////////////////////////////
	BEGIN_PROPSET_MAP(CSQLiteColumnsRowset)
		BEGIN_PROPERTY_SET(DBPROPSET_ROWSET)
			PROPERTY_INFO_ENTRY(IAccessor)
			PROPERTY_INFO_ENTRY(IColumnsInfo)
			PROPERTY_INFO_ENTRY(IConvertType)
			PROPERTY_INFO_ENTRY(IRowset)
			PROPERTY_INFO_ENTRY(IRowsetIdentity)
			PROPERTY_INFO_ENTRY(IRowsetInfo)
			PROPERTY_INFO_ENTRY(IRowsetLocate)																									\
			PROPERTY_INFO_ENTRY(IRowsetScroll)																									\
			PROPERTY_INFO_ENTRY(CANFETCHBACKWARDS)
			PROPERTY_INFO_ENTRY(CANHOLDROWS)
			PROPERTY_INFO_ENTRY(CANSCROLLBACKWARDS)
			PROPERTY_INFO_ENTRY_VALUE(MAXOPENROWS, 0)
			PROPERTY_INFO_ENTRY_VALUE(MAXROWS, 0)
		END_PROPERTY_SET(DBPROPSET_ROWSET)
	END_PROPSET_MAP()	

	/////////////////////////////////////////////////////////////////////////////
	CSQLiteColumnsRowset()
	{
		m_rgColumns = NULL;
		m_cColumns = SQLITE_REQUIRED_METADATA_COLUMNS + SQLITE_OPTIONAL_METADATA_COLUMNS;
	}

	/////////////////////////////////////////////////////////////////////////////
	~CSQLiteColumnsRowset()
	{
		delete [] m_rgColumns;
	}

	/////////////////////////////////////////////////////////////////////////////
	static ATLCOLUMNINFO *GetColumnInfo(CSQLiteColumnsRowset *pv, DBORDINAL *pcCols)
	{
		*pcCols = pv->m_cColumns;
		return pv->m_rgColumns;
	}
	
	/////////////////////////////////////////////////////////////////////////////
	DBSTATUS GetDBStatus(CSimpleRow *pRow, ATLCOLUMNINFO *pInfo)
	{
		CSQLiteColumnsRowsetRow &TR = m_rgRowData[pRow->m_iRowset];

		/////////////////////////////////////////////////////////////////////////////
		// Mandatory metadata columns
		/////////////////////////////////////////////////////////////////////////////

		switch(pInfo->iOrdinal)
		{
		case  1: return DBSTATUS_S_OK;		// DBCOLUMN_IDNAME
		case  2: return DBSTATUS_S_OK;		// DBCOLUMN_GUID
		case  3: return DBSTATUS_S_ISNULL;	// DBCOLUMN_PROPID
		case  4: return DBSTATUS_S_OK;		// DBCOLUMN_NAME
		case  5: return DBSTATUS_S_OK;		// DBCOLUMN_NUMBER
		case  6: return DBSTATUS_S_OK;		// DBCOLUMN_TYPE				
		case  7: return DBSTATUS_S_ISNULL;	// DBCOLUMN_TYPEINFO

		case  8: // DBCOLUMN_COLUMNSIZE
			if(TR.m_DBCOLUMN_COLUMNSIZE==(ULONG)~0 || TR.m_DBCOLUMN_COLUMNSIZE==0)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;

		case 9: // DBCOLUMN_PRECISION				
			if(TR.m_DBCOLUMN_PRECISION==(UINT16)~0)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;

		case 10: // DBCOLUMN_SCALE
			if(TR.m_DBCOLUMN_SCALE==(UINT16)~0)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;

		case 11: return DBSTATUS_S_OK;		// DBCOLUMN_FLAGS
		}

		/////////////////////////////////////////////////////////////////////////////
		// Optional metadata columns
		/////////////////////////////////////////////////////////////////////////////
		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_DATETIMEPRECISION, sizeof(DBID))==0)
		{
			return DBSTATUS_S_OK;
			/*
			if(TR.m_DBCOLUMN_DATETIMEPRECISION==(UINT32)~0)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;
			*/			
		}

		if(memcmp(&pInfo->columnid, &DBCOLUMN_OCTETLENGTH, sizeof(DBID))==0)
		{
			/*
			if(TR.m_DBCOLUMN_OCTETLENGTH==(ULONG)~0)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;
			*/
			return DBSTATUS_S_OK;
		}

		if(memcmp(&pInfo->columnid, &DBCOLUMN_DEFAULTVALUE, sizeof(DBID))==0)
		{
			if(V_VT(&TR.m_DBCOLUMN_DEFAULTVALUE)==VT_EMPTY)
				return DBSTATUS_S_ISNULL;
			else
				return DBSTATUS_S_OK;
		}

		if(memcmp(&pInfo->columnid, &DBCOLUMN_HASDEFAULT, sizeof(DBID))==0)							return DBSTATUS_S_OK;		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_ISSEARCHABLE, sizeof(DBID))==0)						return DBSTATUS_S_OK;		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_ISCASESENSITIVE, sizeof(DBID))==0)					return DBSTATUS_S_OK;		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_ISAUTOINCREMENT, sizeof(DBID))==0)					return DBSTATUS_S_OK;		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_ISUNIQUE, sizeof(DBID))==0)							return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_MAYSORT, sizeof(DBID))==0)							return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_KEYCOLUMN, sizeof(DBID))==0)							return DBSTATUS_S_OK;

		if(memcmp(&pInfo->columnid, &DBCOLUMN_BASECATALOGNAME, sizeof(DBID))==0)					return DBSTATUS_S_OK;		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_BASECOLUMNNAME, sizeof(DBID))==0)						return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_BASETABLENAME, sizeof(DBID))==0)						return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_BASESCHEMANAME, sizeof(DBID))==0)						return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_BASETABLEVERSION, sizeof(DBID))==0)					return DBSTATUS_S_OK;
		
		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_FKEYCOLUMN, sizeof(DBID))==0)					return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_FKEYCOLUMN_TABLE_NAME, sizeof(DBID))==0)		return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_FKEYCOLUMN_NAME, sizeof(DBID))==0)				return DBSTATUS_S_OK;

		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_DATATYPE, sizeof(DBID))==0)					return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_DECLATED_DATATYPE, sizeof(DBID))==0)			return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_SQLITE_CAN_BE_NULL, sizeof(DBID))==0)					return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_CLSID, sizeof(DBID))==0)								return DBSTATUS_S_OK;
		if(memcmp(&pInfo->columnid, &DBCOLUMN_COLLATINGSEQUENCE, sizeof(DBID))==0)					return DBSTATUS_S_OK;
		
		return DBSTATUS_S_ISNULL;
	}
};


