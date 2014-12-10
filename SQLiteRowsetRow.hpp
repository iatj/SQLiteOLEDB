#pragma once

#include "SQLiteOLEDB.h"

// ==================================================================================================================================
//	   ___________ ____    __    _ __       ____                          __  ____               
//	  / ____/ ___// __ \  / /   (_) /____  / __ \____ _      __________  / /_/ __ \____ _      __
//	 / /    \__ \/ / / / / /   / / __/ _ \/ /_/ / __ \ | /| / / ___/ _ \/ __/ /_/ / __ \ | /| / /
//	/ /___ ___/ / /_/ / / /___/ / /_/  __/ _, _/ /_/ / |/ |/ (__  )  __/ /_/ _, _/ /_/ / |/ |/ / 
//	\____//____/\___\_\/_____/_/\__/\___/_/ |_|\____/|__/|__/____/\___/\__/_/ |_|\____/|__/|__/  
//	                                                                                             
// ==================================================================================================================================

class CSQLiteRowsetRowHandle;
			  
/////////////////////////////////////////////////////////////////////////////
template <class T>
class SQLiteRowsetRow	 // Holds the Rowset Data in a string vector
{
public:

	std::vector<std::string> Data;			// The Live Record Data
	std::vector<std::string> OrigData;		// A copy of original Record Data

	/////////////////////////////////////////////////////////////////////////////
	// Purpose of this function is to load columns from an SQLite Statement being 
	// executed. We store (for now) all the values to a string vector so basically
	// we need to convert all our data to strings. For floats we use %g and for 
	// blobs we encode them to Base64 in order to safely keep them as a string.
	// In the future we should either use a Union or something a bit more clever.
	bool load(T* pRowset, sqlite3_stmt* stmt)
	{
		ULONG record_count = pRowset->RECORDS_COUNT();
		
		for(ULONG i=0; i<pRowset->fieldCount; i++) 
		{			
			// Get byte-length of column
			int length = sqlite3_column_bytes(stmt, i);

			// Check for NULL
			if(length==0 && sqlite3_column_type(stmt, i)==SQLITE_NULL)
			{
				Data.push_back(NULL_DATA_VALUE);
				continue;
			}			
			
			// Get the column metadata
			CSQLiteColumnsRowsetRow* DATA_COLUMN = pRowset->METADATA[i+(pRowset->m_HasBookmarks?1:0)];

			// Read column data depending on SQLite underlying datatype
			switch(DATA_COLUMN->m_DBCOLUMN_SQLITE_DATATYPE)
			{
			case SQLITE_INTEGER:
				{
					int result = sqlite3_column_int(stmt, i);
					char str[100];
					int len = sprintf(str, "%d", result);
					std::string value(str);
					Data.push_back(value); 
				}								
				break;

			case SQLITE_FLOAT:
				{
					double result = sqlite3_column_double(stmt, i);
					char str[100];
					int len = sprintf(str, "%g", result);
					std::string value(str);
					Data.push_back(value); 
				}								
				break;

			case SQLITE_BLOB:							
				{					
					char* result = (char*) sqlite3_column_blob(stmt, i);
					std::string value(result,0,length);
					Data.push_back(value);					
				}								
				break;

			case SQLITE_TEXT:
				{			
					// Get UTF8 string
					char* utf8 = (char*) sqlite3_column_text(stmt, i);

					// Convert UTF8 to Unicode					 
					 int len = strlen(utf8);
					 if(len==0)
					 {
						 Data.push_back("");
					 }
					 else
					 {
						 std::string value = UTF8_to_STR(utf8);
						 Data.push_back(value);
					 }
				}								
				break;

			default:				
				return false;				
			}
		}

		// ---------------------------------------------------
		// Create new Handle Row
		// ---------------------------------------------------
		DBCOUNTITEM Rows = 0;
		CSQLiteRowsetRowHandle::KeyType key = record_count+1;
		CSQLiteRowsetRowHandle* pRow = new CSQLiteRowsetRowHandle(record_count);
		pRow->m_status = DBPENDINGSTATUS_UNCHANGED;
		pRow->m_Bookmark = record_count+3;
		pRowset->m_rgRowHandles.SetAt(key, pRow);		
		pRow->AddRefRow();

		return true;
	}

	/////////////////////////////////////////////////////////////////////////////
	// Purpose of this function is to generate the update SQL statements.
	// A string stream is passed in order to collect all the SQL statements 
	// for all affected records. We generate update SQL statements only for 
	// records that have changed, they have a Primary Key column and the 
	// base table of the key column is the same with the column being updated.
	DBPENDINGSTATUS update(T* pRowset, CSQLiteRowsetRowHandle* pRow, std::stringstream& SQL)
	{		
		ULONG KF = pRowset->KEY_COLUMN();
		CSQLiteColumnsRowsetRow* KEY_COLUMN = KF==-1 ? NULL : pRowset->METADATA[KF];
		std::string ID = KF!=-1 ? Data[KF-(pRowset->m_HasBookmarks?1:0)] : "";

		switch (pRow->m_status)
		{
			//------------------------------------------------------------
		case 0:
		case DBPENDINGSTATUS_UNCHANGED:
			return DBROWSTATUS_S_OK;			

			//------------------------------------------------------------
		case DBPENDINGSTATUS_INVALIDROW:
			return DBROWSTATUS_E_DELETED;			

			//------------------------------------------------------------
		case DBPENDINGSTATUS_NEW:
			{
				SQL << "INSERT INTO " + KEY_COLUMN->BASETABLENAME + " (";
				bool first = false;
				for(ULONG i=0; i<pRowset->fieldCount; i++)
				{
					ULONG Index = i + (pRowset->m_HasBookmarks ? 1 : 0);
					CSQLiteColumnsRowsetRow* col = pRowset->METADATA[Index];
					if(col->m_DBCOLUMN_KEYCOLUMN==VARIANT_TRUE && col->m_DBCOLUMN_ISAUTOINCREMENT==VARIANT_TRUE) continue;
					if(col->BASETABLENAME==KEY_COLUMN->BASETABLENAME)
					{
						if(first) SQL << ",";
						SQL << col->BASECOLUMNNAME;
						first=true;
					}
				}
				SQL << ") VALUES (";
				first = false;
				for(ULONG i=0; i<pRowset->fieldCount; i++)
				{				
					ULONG Index = i + (pRowset->m_HasBookmarks ? 1 : 0);
					CSQLiteColumnsRowsetRow* col = pRowset->METADATA[Index];
					if(col->m_DBCOLUMN_KEYCOLUMN==VARIANT_TRUE && col->m_DBCOLUMN_ISAUTOINCREMENT==VARIANT_TRUE) continue;
					if(col->BASETABLENAME==KEY_COLUMN->BASETABLENAME)
					{
						if(first) SQL << ",";
						SQL << SQL_VALUE(i, col->m_DBCOLUMN_SQLITE_DATATYPE);
						first=true;
					}
				}
				SQL << ");\n";
			}
			return DBPENDINGSTATUS_UNCHANGED;			

			//------------------------------------------------------------
		case DBPENDINGSTATUS_CHANGED:

			// To change a row it must have a key column
			if(KF==-1) return DB_E_BADCOLUMNID;			

			for(ULONG i=0; i<pRowset->fieldCount; i++)
			{				
				ULONG Index = i + (pRowset->m_HasBookmarks ? 1 : 0);
				CSQLiteColumnsRowsetRow* col = pRowset->METADATA[Index];
				if(col->BASETABLENAME==KEY_COLUMN->BASETABLENAME && Data[i]!=OrigData[i])
				{
					if(col->m_DBCOLUMN_KEYCOLUMN==VARIANT_TRUE && col->m_DBCOLUMN_ISAUTOINCREMENT==VARIANT_TRUE) continue;
					SQL << "UPDATE " + col->BASETABLENAME + " SET " + col->BASECOLUMNNAME + " = " + SQL_VALUE(i, col->m_DBCOLUMN_SQLITE_DATATYPE) + " WHERE " + KEY_COLUMN->FILED_NAME + " = " + ID + ";\n";
				}
			}				

			return DBPENDINGSTATUS_UNCHANGED;

			//------------------------------------------------------------
		case DBPENDINGSTATUS_DELETED:				

			// To delete a row it must have a key column
			if(KF==-1) return DB_E_BADCOLUMNID;

			// Delete the row and invalidate bookmark
			SQL << "DELETE FROM " + KEY_COLUMN->BASETABLENAME + " WHERE " + KEY_COLUMN->FILED_NAME + " = " + ID + ";\n";
			DBROWCOUNT dwBookmark = pRow->m_iRowset+3;
			pRowset->m_rgBookmarks[dwBookmark] = -1;

			return DBPENDINGSTATUS_INVALIDROW;
		}

		return S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	std::string SQL_VALUE(ULONG index, int SQLiteType)
	{
		if(Data[index]==NULL_DATA_VALUE)
		{
			return "NULL";
		}
		else
		{
			switch(SQLiteType)
			{
			case SQLITE_INTEGER:
				return Data[index];

			case SQLITE_FLOAT:
				return Data[index];

			case SQLITE_BLOB:						
				{
					int sz = Data[index].size();
					int bz = Base64EncodeGetRequiredLength(sz, ATL_BASE64_FLAG_NOCRLF);
					std::string s;
					s.assign(bz, 0);
					Base64Encode((BYTE*)Data[index].c_str(), sz, (char*)s.c_str(), &bz, ATL_BASE64_FLAG_NOCRLF);
					return "BASE64_DECODE('" + s + "')";
				}
			
			case SQLITE_TEXT:
			default:				
				return "'" + Data[index] + "'";
			}
		}
	}

	/////////////////////////////////////////////////////////////////////////////
	// Purpose of this function is to copy data from this record to a consumer buffer.
	// pData is the client buffer we need to write to, and the exact position where we
	// should write is defined by pBinding fields. We need to pass the actual data to 
	// m_spConvert->DataConvert() function. This is important because the function will
	// return the dbStat that is needed if the data is NULL. In principle we must convert
	// our data to the format that binder needs. All supported data types are converted to
	// a BSTR taking care UTF8 encoding and then we call DataConvert() to let OLEDB services
	// transcode the data. If the function fails we default to NULL value. 		
	HRESULT read(T* pRowset, void* pData, DBBINDING *pBinding, DBROWCOUNT dwBookmark, bool OriginalData)
	{
		USES_CONVERSION;

		HRESULT hr = S_OK;
		DBLENGTH cbDst = 4;		
		DBSTATUS dbStat = DBSTATUS_S_OK;
		ULONG cbBytesLen = pBinding->cbMaxLen;
		DBLENGTH cbMaxLen = pBinding->cbMaxLen;
		ULONG ColIndex = pBinding->iOrdinal + (pRowset->m_HasBookmarks ? 0 : -1);			
		ATLCOLUMNINFO *pColCur = &(pRowset->m_ColumnsMetadata[ColIndex]);		
		CSQLiteColumnsRowsetRow* DATA_COLUMN = pRowset->METADATA[ColIndex];
		
		// Consumer needs Bookmark?
		if(pBinding->iOrdinal==0)
		{
			if(pBinding->dwPart & DBPART_VALUE)
			{
				if(pBinding->dwPart & DBPART_STATUS)
				{
					DBLENGTH cbLen = pBinding->obStatus - pBinding->obValue;
					if(cbLen<cbMaxLen) cbMaxLen = cbLen;
				}
				if(pBinding->dwPart & DBPART_LENGTH)
				{
					DBLENGTH cbLen = pBinding->obLength - pBinding->obValue;
					if(cbLen<cbMaxLen) cbMaxLen = cbLen;
				}

				BYTE *pDstTemp = (BYTE *) pData + pBinding->obValue;
				hr = pRowset->m_spConvert->DataConvert(DBTYPE_I4, pBinding->wType, 4, &cbBytesLen, &dwBookmark, pDstTemp, pBinding->cbMaxLen, dbStat, &dbStat, pBinding->bPrecision, pBinding->bScale, 0);				
			}

			if(pBinding->dwPart & DBPART_LENGTH)
				*(DBLENGTH *)((BYTE *)pData + pBinding->obLength) = ( dbStat==DBSTATUS_S_ISNULL ? 0 : cbBytesLen );

			if(pBinding->dwPart & DBPART_STATUS)
				*(DBSTATUS *)((BYTE *)pData + pBinding->obStatus) = dbStat;				
		}

		// Consumer needs Non-Bookmark Data Columns
		else
		{											
			// Value
			if (pBinding->dwPart & DBPART_VALUE)							
			{				
				// Read the data from the row (we keep all data as string values)
				ULONG RowIndex = ColIndex-(pRowset->m_HasBookmarks?1:0);
				std::string strData = OriginalData ? OrigData[RowIndex] : Data[RowIndex];				
				cbBytesLen = 0;

				// Check if data are null
				if(strData == NULL_DATA_VALUE)
				{
					dbStat = DBSTATUS_S_ISNULL;
				}
				else
				{	
					// Calculate the pointer we need to write to.
					BYTE* pDataValue = (BYTE*)pData + pBinding->obValue;
					
					#ifdef _DEBUG
					{
						char error[256];			
						sprintf(error, "READ: field %s, column datatype: %s (%d), binding datatype: %s (%d), column size: %d, octet length: %d\n", 
						OLE2A(pColCur->columnid.uName.pwszName),
						OLEDB_TYPENAME(pColCur->wType).c_str(), pColCur->wType,							   
						OLEDB_TYPENAME(pBinding->wType).c_str(), pBinding->wType,
						pColCur->ulColumnSize, DATA_COLUMN->m_DBCOLUMN_OCTETLENGTH);
						OutputDebugStringA(error);					
					}
					#endif

					switch(pColCur->wType)
					{
					// ====================================================
					// BOOL
					// ====================================================
					case DBTYPEENUM::DBTYPE_BOOL:
						{
							cbBytesLen = sizeof(VARIANT_BOOL);
							VARIANT_BOOL v = atoi(strData.c_str())==0 ? VARIANT_FALSE : VARIANT_TRUE;
							hr = pRowset->m_spConvert->DataConvert(DBTYPEENUM::DBTYPE_BOOL, pBinding->wType, cbBytesLen, &cbBytesLen, (void*)&v, pDataValue, pBinding->cbMaxLen, dbStat, &dbStat, pBinding->bPrecision, pBinding->bScale, 0); 
						} 
						break;

					// ====================================================
					// BLOB	treated as VARIANT(ARRAY|UI1)
					// ====================================================
					case DBTYPE_BLOB:					
						{								
							switch(pBinding->wType)
							{
								case DBTYPEENUM::DBTYPE_VARIANT:
								{
									cbBytesLen = strData.size();
									SAFEARRAYBOUND sabounds[1];							
									sabounds[0].lLbound = 0;							
									sabounds[0].cElements = cbBytesLen;									
									VARIANT var;
									VariantInit(&var);
									var.vt = VT_ARRAY|VT_UI1;							
									var.parray = SafeArrayCreate(VT_UI1, 1, sabounds);							
									SAFEARRAY* psa = V_ARRAY(&var);							
									char *blob;
									SafeArrayAccessData(psa, (void**)&blob);
									memset(blob, 0, cbBytesLen);
									memcpy(blob, strData.c_str(), cbBytesLen);
									SafeArrayUnaccessData(psa);
									hr = pRowset->m_spConvert->DataConvert(DBTYPEENUM::DBTYPE_VARIANT, pBinding->wType, sizeof(VARIANT), &cbBytesLen, (void*)&var, pDataValue, pBinding->cbMaxLen, dbStat, &dbStat, pBinding->bPrecision, pBinding->bScale, 0); 
									VariantClear(&var);								
								}
								break;

								case DBTYPEENUM::DBTYPE_WSTR:
									{
										cbBytesLen = strData.size();
										hr = pRowset->m_spConvert->DataConvert(DBTYPEENUM::DBTYPE_STR, pBinding->wType, cbBytesLen, &cbBytesLen, (void*)(strData.c_str()), pDataValue, pBinding->cbMaxLen, dbStat, &dbStat, pBinding->bPrecision, pBinding->bScale, 0);
									}
									break;

								default:
									hr = DB_E_UNSUPPORTEDCONVERSION;
									break;
							}
						}
						break;

					// ====================================================
					// All other types are converted to string
					// ====================================================
					default:
						{
							BSTR pSrcData = STR_to_BSTR(strData);
							cbBytesLen = SysStringByteLen(pSrcData);
							hr = pRowset->m_spConvert->DataConvert(DBTYPEENUM::DBTYPE_BSTR, pBinding->wType, cbBytesLen, &cbBytesLen, (void*)&pSrcData, pDataValue, pBinding->cbMaxLen, dbStat, &dbStat, pBinding->bPrecision, pBinding->bScale, 0);
							SysFreeString(pSrcData);
						}
						break;
					}
				}
			}

			if(hr==S_OK)
			{
				// Length
				if (pBinding->dwPart & DBPART_LENGTH)
				{
					*((ULONG*)((BYTE*)(pData) + pBinding->obLength)) = (dbStat==DBSTATUS_S_ISNULL ? 0 : cbBytesLen );		
				}

				// Status
				if (pBinding->dwPart & DBPART_STATUS)
				{
					*((DBSTATUS*)((BYTE*)(pData) + pBinding->obStatus)) = dbStat;
				}
			}
		}

		if(hr!=S_OK)
		{			
			#ifdef _DEBUG
			{
				char error[256];			
				sprintf(error, "READ-ERROR: field %s, column datatype: %s (%d), binding datatype: %s (%d), column size: %d, octet length: %d\n", 
							   OLE2A(pColCur->columnid.uName.pwszName),
							   OLEDB_TYPENAME(pColCur->wType).c_str(), pColCur->wType,							   
							   OLEDB_TYPENAME(pBinding->wType).c_str(), pBinding->wType,
							   pColCur->ulColumnSize, DATA_COLUMN->m_DBCOLUMN_OCTETLENGTH);
				OutputDebugStringA(error);				
			}
			#endif
		}

		return hr;
	}

	/////////////////////////////////////////////////////////////////////////////
	// Purpose of this function is to read data from a consumer buffer and store them in this record.
	HRESULT write(T* pRowset, void* pData, DBBINDING *pBinding, DBROWCOUNT dwBookmark)
	{
		USES_CONVERSION;

		HRESULT hr = S_OK;
		DBLENGTH cbDst = 4;				
		DBSTATUS dbStat = DBSTATUS_S_OK;
		DBLENGTH cbMaxLen = pBinding->cbMaxLen;
		ULONG ColIndex = pBinding->iOrdinal + (pRowset->m_HasBookmarks ? 0 : -1);			
		ATLCOLUMNINFO *pColCur = &(pRowset->m_ColumnsMetadata[ColIndex]);
		CSQLiteColumnsRowsetRow* DATA_COLUMN = pRowset->METADATA[ColIndex];
		ColIndex -= (pRowset->m_HasBookmarks?1:0);		

		#ifdef _DEBUG
		{
			char error[256];			
			sprintf(error, "WRITE: field %s, column datatype: %s (%d), binding datatype: %s (%d), column size: %d, octet length: %d\n", 
			OLE2A(pColCur->columnid.uName.pwszName),
			OLEDB_TYPENAME(pColCur->wType).c_str(), pColCur->wType,							   
			OLEDB_TYPENAME(pBinding->wType).c_str(), pBinding->wType,
			pColCur->ulColumnSize, DATA_COLUMN->m_DBCOLUMN_OCTETLENGTH);
			OutputDebugStringA(error);					
		}
		#endif

		// Status
		if(pBinding->dwPart & DBPART_STATUS)
		{
			dbStat = *((DBSTATUS*)((BYTE*)(pData) + pBinding->obStatus));

			// Consumer requests this value set to NULL.
			if(dbStat==DBSTATUS_S_ISNULL)
			{
				Data[ColIndex]=NULL_DATA_VALUE;
				return S_OK;
			}
		}

		// Data Length
		ULONG cbBytesLen = 0;
		if(pBinding->dwPart & DBPART_LENGTH)
		{
			cbBytesLen = (dbStat==DBSTATUS_S_ISNULL ? 0 : *((ULONG*)((BYTE*)(pData) + pBinding->obLength)) );		
		}

		// Data Value
		if (pBinding->dwPart & DBPART_VALUE)							
		{
			// Calculate the pointer we need to read from
			BYTE* pDataValue = (BYTE*)pData + pBinding->obValue;
			
			switch(pColCur->wType)
			{
			// ====================================================
			// Date and Time
			// ====================================================
			case DBTYPEENUM::DBTYPE_DBDATE:			
			case DBTYPEENUM::DBTYPE_DBTIME:
			case DBTYPEENUM::DBTYPE_DBTIMESTAMP:
				{					
					DBTIMESTAMP dt;
					cbBytesLen = pBinding->cbMaxLen;
					hr = pRowset->m_spConvert->DataConvert(pBinding->wType, DBTYPEENUM::DBTYPE_DBTIMESTAMP, cbBytesLen, &cbBytesLen, pDataValue, (void*)&dt, pColCur->ulColumnSize, dbStat, &dbStat, pColCur->bPrecision, pColCur->bScale, 0);											
					Data[ColIndex] = DBDateTimeToISO8601(pColCur->wType, dt);
				}
				break;	

			// ====================================================
			// BOOL
			// ====================================================
			case DBTYPEENUM::DBTYPE_BOOL:
				{
					VARIANT_BOOL v = *(VARIANT_BOOL*)((BYTE*)(pDataValue));
					Data[ColIndex] = v==VARIANT_TRUE ? "1" : "0";
					hr = S_OK;
				}
				break;

			// ====================================================
			// BLOB	treated as VARIANT(ARRAY|UI1)
			// ====================================================
			case DBTYPE_BLOB:
				{
					VARIANT* var = (VARIANT *)pDataValue;					
					switch(var->vt)
					{
					case VT_NULL:						
						Data[ColIndex]=NULL_DATA_VALUE;
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
							Data[ColIndex] = std::string(blob, 0, cbBytesLen);
							SafeArrayUnaccessData(psa);
							hr = S_OK;
						}
						break;

					default:
						hr = DB_E_UNSUPPORTEDCONVERSION;
					}
				}
				break;

			// ====================================================
			// WSTR
			// ====================================================
			case DBTYPEENUM::DBTYPE_WSTR:
				{
					// The consumer more likely will send invalid string length
					// so we need to cast to BSTR and get actual byte length.
					// This is the most secure method since BSTR keeps length.
					// Make sure we pass byte length to BSTR_TO_UTF8 because
					// Microsoft Rowset Viewer also basses BLOBs as strings,
					// so there might be other clients that do that as well.

					BSTR pBSTR = *(BSTR*)pDataValue;
					Data[ColIndex] = BSTR_to_STR(pBSTR);
					hr = S_OK;
				}
				break;

			// ====================================================
			// All other types are converted to string
			// ====================================================
			default:
				{										
					BSTR pBSTR = SysAllocString(L"");
					if(cbBytesLen==0) cbBytesLen = pBinding->cbMaxLen;
					hr = pRowset->m_spConvert->DataConvert(pBinding->wType, DBTYPEENUM::DBTYPE_BSTR, cbBytesLen, &cbBytesLen, pDataValue, (void*)&pBSTR, pColCur->ulColumnSize, dbStat, &dbStat, NULL, NULL, 0);
					if(hr==S_OK) Data[ColIndex] = BSTR_to_STR(pBSTR);
					SysFreeString(pBSTR);
				}
				break;
			}
		}

		if(hr!=S_OK)
		{			
			#ifdef _DEBUG
			{
				char error[256];			
				sprintf(error, "WRITE-ERROR: field %s, column datatype: %s (%d), binding datatype: %s (%d), column size: %d, octet length: %d\n", 
					OLE2A(pColCur->columnid.uName.pwszName),
					OLEDB_TYPENAME(pColCur->wType).c_str(), pColCur->wType,							   
					OLEDB_TYPENAME(pBinding->wType).c_str(), pBinding->wType,
					pColCur->ulColumnSize, DATA_COLUMN->m_DBCOLUMN_OCTETLENGTH);
				OutputDebugStringA(error);				
			}
			#endif
		}

		return hr;
	}
};

