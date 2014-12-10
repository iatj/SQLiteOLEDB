#pragma once

#pragma warning(disable : 4482)

#include "Resource.h"							
#include "sqlite3.h"
#include "md5.h"
#include "vbscript.h"

#include <atlcom.h>								
#include <atldb.h>								
#include <comutil.h>							
#include <stdint.h>								
#include <string>								
#include <vector>
#include <algorithm>
#include <string>
#include <atlenc.h>

#include <string>
#include <regex>

/////////////////////////////////////////////////////////////////////////////
//#define THREAD_MODEL CComSingleThreadModel
#define THREAD_MODEL CComMultiThreadModel
#define NULL_DATA_VALUE "FF707534-E45E-4294-A242-E7B798BF96A7" 
#define MAX_COLUMN_BYTES (1024*10)
#define DBTYPE_BLOB DBTYPEENUM(DBTYPE_VARIANT)

/////////////////////////////////////////////////////////////////////////////
//http://devzone.advantagedatabase.com/dz/webhelp/Advantage7.1/mergedProjects/adsoledb/adsoledb/rowset_properties.htm

#define ROWSET_PROPSET_MAP(Class)																																\
BEGIN_PROPSET_MAP(Class)																																		\
	BEGIN_PROPERTY_SET(DBPROPSET_ROWSET)																														\
																																								\
		PROPERTY_INFO_ENTRY						(IAccessor)																										\
		PROPERTY_INFO_ENTRY						(IColumnsInfo)																									\
		PROPERTY_INFO_ENTRY_VALUE				(IColumnsRowset, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY						(IConvertType)																									\
		PROPERTY_INFO_ENTRY						(IRowset)																										\
		PROPERTY_INFO_ENTRY						(IRowsetIdentity)																								\
		PROPERTY_INFO_ENTRY						(IRowsetInfo)																									\
		PROPERTY_INFO_ENTRY_VALUE				(IRowsetLocate, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(IRowsetExactScroll, ATL_VARIANT_TRUE)																			\
		PROPERTY_INFO_ENTRY_VALUE				(IRowsetScroll, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(IRowsetChange, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(IRowsetUpdate, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(ISequentialStream, ATL_VARIANT_TRUE)																			\
																																								\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(ABORTPRESERVE, ATL_VARIANT_TRUE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(BOOKMARKS, ATL_VARIANT_TRUE)																					\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(BOOKMARKSKIPPED, ATL_VARIANT_TRUE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(BOOKMARKTYPE, DBPROPVAL_BMK_NUMERIC, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)									\
		PROPERTY_INFO_ENTRY_VALUE				(BLOCKINGSTORAGEOBJECTS, ATL_VARIANT_TRUE)																		\
																																								\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(COMMITPRESERVE, ATL_VARIANT_FALSE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
		PROPERTY_INFO_ENTRY_VALUE				(CANFETCHBACKWARDS, ATL_VARIANT_TRUE) 																			\
		PROPERTY_INFO_ENTRY_VALUE				(CANSCROLLBACKWARDS, ATL_VARIANT_TRUE)																			\
		PROPERTY_INFO_ENTRY_VALUE				(CANHOLDROWS, ATL_VARIANT_FALSE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(CHANGEINSERTEDROWS, VARIANT_TRUE)																				\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(IMMOBILEROWS, VARIANT_TRUE)																					\
																																								\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(LITERALBOOKMARKS, ATL_VARIANT_TRUE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(LITERALIDENTITY, ATL_VARIANT_TRUE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(MAXROWS, 0)																									\
		PROPERTY_INFO_ENTRY_VALUE				(MAXROWSIZE, 0)																									\
		PROPERTY_INFO_ENTRY_VALUE				(MAXOPENROWS, 0)																								\
		PROPERTY_INFO_ENTRY_VALUE				(MAXTABLESINSELECT, 0)																							\
		PROPERTY_INFO_ENTRY_VALUE				(MAXROWSIZEINCLUDESBLOB, VARIANT_FALSE)																			\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(MAXPENDINGROWS, 0, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)														\
		PROPERTY_INFO_ENTRY_VALUE				(MULTIPLESTORAGEOBJECTS, VARIANT_TRUE)																			\
		PROPERTY_INFO_ENTRY_VALUE				(MULTITABLEUPDATE, VARIANT_FALSE)																				\
																																								\
		PROPERTY_INFO_ENTRY_EX					(OWNINSERT, VT_BOOL, DBPROPFLAGS_ROWSET|DBPROPFLAGS_READ, VARIANT_TRUE, 0)										\
		PROPERTY_INFO_ENTRY_EX					(OWNUPDATEDELETE, VT_BOOL, DBPROPFLAGS_ROWSET|DBPROPFLAGS_READ, VARIANT_TRUE, 0)								\
		PROPERTY_INFO_ENTRY_EX					(OTHERINSERT, VT_BOOL, DBPROPFLAGS_ROWSET|DBPROPFLAGS_READ, ATL_VARIANT_FALSE, 0)								\
		PROPERTY_INFO_ENTRY_EX					(OTHERUPDATEDELETE, VT_BOOL, DBPROPFLAGS_ROWSET|DBPROPFLAGS_READ, ATL_VARIANT_FALSE, 0)							\
		PROPERTY_INFO_ENTRY_VALUE_FLAGS			(ORDEREDBOOKMARKS, ATL_VARIANT_TRUE, DBPROPFLAGS_ROWSET | DBPROPFLAGS_READ)										\
		PROPERTY_INFO_ENTRY_VALUE				(OLEOBJECTS, DBPROPVAL_OO_BLOB)																					\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(PERSISTENTIDTYPE, DBPROPVAL_PT_GUID_NAME)																		\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(QUICKRESTART, VARIANT_TRUE)																					\
																																								\
		PROPERTY_INFO_ENTRY_EX					(REMOVEDELETED, VT_BOOL, DBPROPFLAGS_ROWSET|DBPROPFLAGS_READ, VARIANT_FALSE, 0)									\
		PROPERTY_INFO_ENTRY_VALUE				(REPORTMULTIPLECHANGES, VARIANT_FALSE)																			\
		PROPERTY_INFO_ENTRY_VALUE				(ROWRESTRICT, VARIANT_FALSE)																					\
		PROPERTY_INFO_ENTRY_VALUE				(ROWTHREADMODEL, DBPROPVAL_RT_APTMTTHREAD)																		\
		PROPERTY_INFO_ENTRY_VALUE				(REENTRANTEVENTS, ATL_VARIANT_FALSE)																			\
		PROPERTY_INFO_ENTRY_VALUE				(RETURNPENDINGINSERTS, VARIANT_FALSE)																			\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(SERVERCURSOR, ATL_VARIANT_TRUE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(STRONGIDENTITY, ATL_VARIANT_FALSE)																				\
		PROPERTY_INFO_ENTRY_VALUE				(SQLSUPPORT, DBPROPVAL_SQL_ODBC_MINIMUM |DBPROPVAL_SQL_ANSI92_ENTRY |DBPROPVAL_SQL_ESCAPECLAUSES)				\
		PROPERTY_INFO_ENTRY_VALUE				(STRUCTUREDSTORAGE, DBPROPVAL_SS_ISEQUENTIALSTREAM)																\
																																								\
		PROPERTY_INFO_ENTRY_VALUE				(UPDATABILITY, DBPROPVAL_UP_CHANGE|DBPROPVAL_UP_INSERT|DBPROPVAL_UP_DELETE)										\
																																								\
	END_PROPERTY_SET(DBPROPSET_ROWSET)																															\
END_PROPSET_MAP()																																				\

#define PROVIDER_PROPSET_MAP(Class)	ROWSET_PROPSET_MAP(Class)
#define SESSION_PROPSET_MAP(Class)	ROWSET_PROPSET_MAP(Class)
#define COMMAND_PROPSET_MAP(Class)	ROWSET_PROPSET_MAP(Class)

/////////////////////////////////////////////////////////////////////////////
BOOL IsVariableType(DBTYPE wType)
{
	//According to OLE DB Spec Appendix A (Variable-Length Data Types)
	switch(wType) 
	{
	case DBTYPE_STR:
	case DBTYPE_WSTR:
	case DBTYPE_BYTES:
	case DBTYPE_VARNUMERIC:
		return TRUE;
	}
	return FALSE;
}

/////////////////////////////////////////////////////////////////////////////
BOOL IsFixedType(DBTYPE wType)
{
	return !IsVariableType(wType);
}

/////////////////////////////////////////////////////////////////////////////
BOOL IsNumericType(DBTYPE wType)
{
	//According to OLE DB Spec Appendix A (Numeric Data Types)
	switch(wType) 
	{
	case DBTYPE_I1:
	case DBTYPE_I2:
	case DBTYPE_I4:
	case DBTYPE_I8:
	case DBTYPE_UI1:
	case DBTYPE_UI2:
	case DBTYPE_UI4:
	case DBTYPE_UI8:
	case DBTYPE_R4:
	case DBTYPE_R8:
	case DBTYPE_CY:
	case DBTYPE_DECIMAL:
	case DBTYPE_NUMERIC:
		return TRUE;
	}
	return FALSE;
}

/////////////////////////////////////////////////////////////////////////////
BSTR UTF8_to_BSTR(const char* utf8, int len=-1)
{	
	if(len==-1) len = strlen(utf8);
	if(len==0) return SysAllocString(L"");
	int size_needed = MultiByteToWideChar(CP_UTF8, 0, utf8, len, NULL, 0);
	std::wstring wstrTo(size_needed, 0);
	MultiByteToWideChar(CP_UTF8, 0, utf8, len, &wstrTo[0], size_needed);	
	return SysAllocString((BSTR)(wstrTo.data()));
}

/////////////////////////////////////////////////////////////////////////////
std::string BSTR_to_UTF8(BSTR& value, int len=-1)
{
	int wlen = wcslen(value);
	//int blen = SysStringLen(value);
	UINT slen = len !=-1 ? len : wlen;
	if(slen==0) return std::string("");
	int size_needed = WideCharToMultiByte(CP_UTF8, 0, value, (int)slen, NULL, 0, NULL, NULL);
	std::string strTo(size_needed, 0);
	WideCharToMultiByte(CP_UTF8, 0, value, (int)slen, &strTo[0], size_needed, NULL, NULL);
	return strTo;
}

/////////////////////////////////////////////////////////////////////////////
std::string ConvertWCSToMBS(const wchar_t* pstr, long wslen)
{
	int len = ::WideCharToMultiByte(CP_ACP, 0, pstr, wslen, NULL, 0, NULL, NULL);
	std::string dblstr(len, '\0');
	len = ::WideCharToMultiByte(CP_ACP, 0, pstr, wslen,	&dblstr[0], len, NULL, NULL);
	return dblstr;
}

/////////////////////////////////////////////////////////////////////////////
BSTR STR_to_BSTR(const std::string& str)
{
	int wslen = MultiByteToWideChar(CP_ACP, 0, str.data(), str.length(), NULL, 0);
	BSTR wsdata = SysAllocStringLen(NULL, wslen);
	MultiByteToWideChar(CP_ACP, 0, str.data(), str.length(), wsdata, wslen);
	return wsdata;
}

/////////////////////////////////////////////////////////////////////////////
std::string BSTR_to_STR(BSTR bstr)
{
	int wslen = SysStringLen(bstr);
	return ConvertWCSToMBS((wchar_t*)bstr, wslen);
}

/////////////////////////////////////////////////////////////////////////////
std::string UTF8_to_STR(const char* utf8, int len=-1)
{
	BSTR b = UTF8_to_BSTR(utf8, len);
	std::string s = BSTR_to_STR(b);
	SysFreeString(b);
	return s;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
char* STR_TO_UTF8(std::string& u)
{
	BSTR b = STR_to_BSTR(u);
	std::string s = BSTR_to_UTF8(b);
	char* c = (char*)malloc(s.size()+1);
	memset(c, 0, s.size()+1);
	memcpy(c, s.c_str(), s.size());
	return c;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
inline BSTR str2bstr(std::string& s)
{	
	return UTF8_to_BSTR(s.c_str());
}

/////////////////////////////////////////////////////////////////////////////
inline void NO_BRACKETS(std::string& v)
{
	if(v.c_str()[0]=='[' && v.c_str()[v.size()-1]==']')
		v = v.substr(1, v.size()-2);
}

/////////////////////////////////////////////////////////////////////////////
void NO_BRACKETS(WCHAR v[])
{
	ULONG sz = wcslen(v);
	std::string s = BSTR_to_UTF8(v,sz);
	NO_BRACKETS(s);
	BSTR b = UTF8_to_BSTR(s.c_str());	
	::wcscpy(v, b);	
	SysFreeString(b);	
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
static int sqlite_callback_tables(void *NotUsed, int argc, char **argv, char **azColName)	
{
	std::vector<std::string>* TABLES = (std::vector<std::string>*)(NotUsed);
	std::string TABLE(argv[0]);	
	NO_BRACKETS(TABLE);
	TABLES->push_back(TABLE);
	return 0;
}
std::vector<std::string> SQLiteTables(sqlite3* db)
{
	std::vector<std::string> TABLES;
	char *zErrMsg = 0;	
	int rc = sqlite3_exec(db, "SELECT name as [TABLE] FROM sqlite_master WHERE type='table';", sqlite_callback_tables, &TABLES, &zErrMsg);
	return TABLES;
}

/////////////////////////////////////////////////////////////////////////////
std::string DBDateTimeToISO8601(DBTYPE wType, DBTIMESTAMP& dt)
{
	char buffer[80];
	
	switch(wType)
	{
	case DBTYPEENUM::DBTYPE_DATE:			
	case DBTYPEENUM::DBTYPE_DBDATE:			
		sprintf(buffer, "%04d-%02d-%02d\0", dt.year, dt.month, dt.day);
		break;					

	case DBTYPEENUM::DBTYPE_DBTIME:
		sprintf(buffer, "%02d:%02d:%02d\0", dt.hour, dt.minute, dt.second);
		break;					

	case DBTYPEENUM::DBTYPE_DBTIMESTAMP:
		sprintf(buffer, "%04d-%02d-%02d %02d:%02d:%02d\0", dt.year, dt.month, dt.day, dt.hour, dt.minute, dt.second);
		break;
	}

	return std::string(buffer);
}

/////////////////////////////////////////////////////////////////////////////
template <class T, class DeAllocator = CRunTimeFree< T > > 
class CAutoMemRelease 
{ 
public: 
	CAutoMemRelease() 
	{ 
		m_pData = NULL; 
	} 
	CAutoMemRelease(T* pData) 
	{ 
		m_pData = pData; 
	} 
	~CAutoMemRelease() 
	{ 
		Attach(NULL); 
	} 
	void Attach(T* pData) 
	{ 
		DeAllocator::Free(m_pData); 
		m_pData = pData; 
	} 
	T* Detach() 
	{ 
		T* pTemp = m_pData; 
		m_pData = NULL; 
		return pTemp; 
	} 
	T* m_pData; 
};

/////////////////////////////////////////////////////////////////////////////
template <class T> 
class CComFree 
{ 
public: 
	static void Free(T* pData) { ::CoTaskMemFree(pData); } 
};

///////////////////////////////////////////////////////////////////////////////////////////////////
class CRegExp
{
public:

	CRegExp()
	{
		HRESULT hr = CoCreateInstance(CLSID_RegExp, NULL, CLSCTX_INPROC_SERVER, IID_IRegExp2, (void**)&mReg);
		hr = mReg->put_Multiline(-1);
		hr = mReg->put_Global(-1);
		hr = mReg->put_IgnoreCase(-1);
		ms = NULL;
	}

	~CRegExp()
	{
		mReg->Release();
		if(ms!=NULL) 
			ms->Release();
	}

	std::string Replace(std::string Buffer, std::string Find, std::string Replace)
	{
		BSTR bs_D = NULL;
		BSTR bs_r = str2bstr(Replace);
		BSTR bs_b = str2bstr(Buffer);
		BSTR bs_f = str2bstr(Find);
		VARIANT v;
		v.vt = VT_BSTR;
		v.bstrVal = bs_r;
		mReg->put_Pattern(bs_f);
		mReg->Replace(bs_b, v, &bs_D);		
		std::string ret = BSTR_to_UTF8(bs_D);
		SysFreeString(bs_D);
		SysFreeString(bs_r);
		SysFreeString(bs_b);
		SysFreeString(bs_f);
		return ret;
	}

	ULONG Parse(std::string sBuffer, std::string sPattern)
	{
		if(ms!=NULL) ms->Release(); 
		BSTR bs_p = str2bstr(sPattern);
		BSTR bs_b = str2bstr(sBuffer);
		mReg->put_Pattern(bs_p);
		mReg->Execute(bs_b, (IDispatch**)&ms);
		SysFreeString(bs_p);
		SysFreeString(bs_b);
		return MatchesCount();
	}

	BOOL TestBool(std::string str, std::string sPattern)
	{
		VARIANT_BOOL matched;
		BSTR bs_s = str2bstr(str);
		BSTR bs_p = str2bstr(sPattern);
		mReg->put_Pattern(bs_p);
		mReg->Test(bs_s, &matched);
		SysFreeString(bs_s);
		SysFreeString(bs_p);
		return matched;
	}

	long MatchesCount()
	{
		long c = 0;
		ms->get_Count(&c);
		return c;
	}

	long SubMatchesCount(long MatchIndex)
	{
		long c;
		IMatch2* pMatch = NULL;
		ISubMatches* pSubMatches = NULL;
		ms->get_Item(MatchIndex, (IDispatch**)&pMatch);
		pMatch->get_SubMatches((IDispatch**)&pSubMatches);
		pSubMatches->get_Count(&c);
		pSubMatches->Release();
		pMatch->Release();
		return c;
	}

	std::string Match(long Index)
	{
		BSTR v;
		IMatch2* pMatch = NULL;
		ms->get_Item(Index, (IDispatch**)&pMatch);
		pMatch->get_Value(&v);
		pMatch->Release();	
		return BSTR_to_UTF8(v);
	}

	std::string SubMatch(long MatchIndex, long SubMatchIndex)
	{
		std::string t;
		IMatch2* pMatch = NULL;
		ISubMatches* pSubMatches = NULL;
		VARIANT subMatch;
		ms->get_Item(MatchIndex, (IDispatch**)&pMatch);
		if(!pMatch) return std::string("");
		pMatch->get_SubMatches((IDispatch**)&pSubMatches);
		if(!pSubMatches) 
		{ 
			pMatch->Release(); 
			return std::string("");
		}
		pSubMatches->get_Item(SubMatchIndex, &subMatch);
		pSubMatches->Release();
		pMatch->Release();
		std::string result = subMatch.vt==VT_EMPTY ? "" : BSTR_to_UTF8(subMatch.bstrVal);
		return result;
	}

	long FirstIndex(long MatchIndex)
	{
		long v;
		std::string t;
		IMatch2* pMatch = NULL;
		ms->get_Item(MatchIndex, (IDispatch**)&pMatch);
		pMatch->get_FirstIndex(&v);
		pMatch->Release();	
		return v;
	}

private:
	IRegExp2* mReg;
	IMatchCollection2* ms;
};

extern CRegExp m_RegExp;

///////////////////////////////////////////////////////////////////////////////////////////////////
static void sqlite3_unbase64_fn(sqlite3_context *f, int nargs, sqlite3_value ** args)
{
	if (nargs != 1)
	{
		sqlite3_result_error(f, "BASE64_DECODE required exactly 1 argument", -1);
		return;
	}
	char* t = (char*) sqlite3_value_text(args[0]);	
	int len=0;
	Base64Decode(t, strlen(t), NULL, &len);	
	BYTE* blob = (BYTE*) malloc(len);
	Base64Decode(t, strlen(t), blob, &len);	
	sqlite3_result_blob(f, blob, len, SQLITE_TRANSIENT);
	free(blob);
}
