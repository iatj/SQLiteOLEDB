Author:

	Elias Politakis
	mobileFX CTO/Partner
	52 Electras str, Kallithea 17673, Greece.	
	www.mobilefx.com


Hi SQLite users,

I got so restless with the lack of performance of OLEDB Providers for SQLite out there and their missing 
features, not to mention the expensive prices!! so I decided to write one and offer it to the community.

The code compiles with Visual C++ 2010 SP1.

I knew nothing about OLE DB Providers 12 hours ago so I did some reasearch and hacked some samples I found
on the internet. It took me few hours to get the code together so please credit me and my company for the effort.

If you are newbe to OLEDB Providers programming like myself, the important bit of information you need is this:

An OLEDB Provider is a COM CoClass (CSQLiteDataSource) that is instantiated when you create an ADODB.Coonection
with the following connection string:

    Dim Conn As ADODB.Connection
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "Provider=SQLiteOLEDBProvider.SQLiteOLEDB.1;Data Source=D:\mobileFX\Projects\Software\SQLiteOLEDB\VBTest\team.db"

The provider creates a session (CSQLiteSession) when you call Connection.Open()

    Conn.Open

================================
PartA: Getting the Schema
================================

In most programming cases you will need to query the database and get its schema.

	Set rs = Conn.OpenSchema(adSchemaColumns)

When you try to get Shema Information the Session object needs to return some "Special" recordsets that we need to
implement according to ATL templates available:

		SCHEMA_ENTRY(DBSCHEMA_TABLES, CRSchema_DBSCHEMA_TABLES)					// Database Tables
		SCHEMA_ENTRY(DBSCHEMA_COLUMNS, CRSchema_DBSCHEMA_COLUMNS)				// Database Table Columns
		SCHEMA_ENTRY(DBSCHEMA_PROVIDER_TYPES, CRSchema_DBSCHEMA_PROVIDER_TYPES)	// Supported Datatypes

You will find that a particularly nessesary "Special" schema recordset has no template, but I created one for you:

		SCHEMA_ENTRY(DBSCHEMA_FOREIGN_KEYS, CRSchema_DBSCHEMA_FOREIGN_KEYS)		

All the above schema recordsets are offered by the CSQLiteSession class. The ATL template has a bug that fails 
to run without crashing. You need edit "atldb.h" and either delete or comment out "ATLASSUME(rgRestrictions != NULL);"
I have no idea what those restrictions are, I just copy-pasted some code that makes sence in "CheckRestrictions" and
"SetRestrictions" and it works.

#define SCHEMA_ENTRY(guid, rowsetClass) \
	if (ppGuid != NULL && SUCCEEDED(hr)) \
	{ \
		cGuids++; \
		*ppGuid = ATL::AtlSafeRealloc<GUID, ATL::CComAllocator>(*ppGuid, cGuids); \
		hr = (*ppGuid == NULL) ? E_OUTOFMEMORY : S_OK; \
		if (SUCCEEDED(hr)) \
			(*ppGuid)[cGuids - 1] = guid; \
		else \
			return hr; \
	} \
	else \
	{ \
		if (InlineIsEqualGUID(guid, rguidSchema)) \
		{ \
			/*ATLASSUME(rgRestrictions != NULL); */ \	  <----------------------(DELETE)----------------------------
			rowsetClass* pRowset = NULL; \
			hr = CheckRestrictions(rguidSchema, cRestrictions, rgRestrictions); \
			if (FAILED(hr)) \
				return E_INVALIDARG; \
			hr =  CreateSchemaRowset(pUnkOuter, cRestrictions, \
							   rgRestrictions, riid, cPropertySets, \
							   rgPropertySets, ppRowset, pRowset); \
			return hr; \
		} \
	}


Now as you can imagine by now the CSQLiteSession class does most of the work. The pattern is this:

1. When the session is created by CSQLiteDataSource::CreateSession() we open the SQLite database and keep it open.

2. We need a RecordClass that describes a single Record for every table we access in order to store field information.
   Luckyly for schema classes there are templates that do that. For real-data classes we need an storage class.

3. We need a RecordSetClass that holds the records. This class is fetching data from SQLite Database and has the 
   following pattern:

class CRecordSetClassXXX : public CRowsetImpl<CRecordXXX, CTABLESRow, CSQLiteSession>
{
public:

	LONG record_count;
	CAtlArray<CRecordXXX,CElementTraits<CRecordXXX>>* m_Data;

	/////////////////////////////////////////////////////////////////////////////
	DBSTATUS GetDBStatus(CSimpleRow*, ATLCOLUMNINFO* pInfo)
	{		
		return DBSTATUS_S_OK;
	}

	/////////////////////////////////////////////////////////////////////////////
	HRESULT Execute(LONG* pcRowsAffected, ULONG, const VARIANT* params)
	{		
		// Get the Session Object
		HRESULT hr;
		CComPtr<IGetDataSource> ipGDS;
		if (FAILED(hr = GetSpecification(IID_IGetDataSource, (IUnknown **)&ipGDS))) return hr;		
		CSQLiteSession *pSess = static_cast<CSQLiteSession *>((IGetDataSource*)ipGDS);

		// Execute the SQL
		char *zErrMsg = 0;	
		record_count=0;		
		m_Data = &m_rgRowData;
		int rc = sqlite3_exec(pSess->db, SQL, sqlite_callback, this, &zErrMsg);
		if(rc) return E_FAIL;
		*pcRowsAffected = record_count;
		return S_OK;
	}

	//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////	
	static int sqlite_callback(void *NotUsed, int argc, char **argv, char **azColName)	
	{
		CRecordSetClassXXX* objPtr = (CRecordSetClassXXX*) NotUsed;

		// Create new record of special data		
		CRecordXXX trData;

		// Save value		
		trData.xxxxx = argv[0];						

		// Add new Record to storage
		objPtr->m_Data->Add(trData);			
		objPtr->record_count++;
		return 0;
	}
};

You will notice that data retreival from SQLite Database is done with static callback "sqlite_callback"
and we use the 1st parameter of the callback to pass the recordset object instance.

Most of recordset data are strings and must be stored as BSTR. Recall to SysAllocFree() your BSTR or you
will suffer memory leaks:

		BSTR name = UTF8_to_BSTR(argv[1]);		
		::wcscpy(trData.m_szColumnName, name);
		SysFreeString(name);		

The more "Special" schema Recordsets you implement the more close you get to being compatible with ADO
Recordsets and the more your client-code will need less changes, if any to work.

However there is a tricky part with ADO Properties but I haven't done anything about them yet.


================================
PartB - SELECTing data
================================

	Rowsets are the internal structure of ADO Recordsets. In fact, ADO Recordsets are just wrappers
	that callback the functions of numerous interfaces implemented by a Rowset. A Rowset has many 
	internal banks, often MFC Arrays that maintain operational and functional data such as Rows,
	Bookmarks and Binders.

	A Row is two things: A "Data Row" (also called User Row) that represents a Record and holds 
	the actual data obtained by executing the command's SQL query, and a "Handles Row" that holds
	the status and the cursor of the Row (aka. Row Index). Thus for every Row in our Rowset we need
	two (2) objects: 

	- A CSQLiteRowsetRowHandle object that holds Row Handle, Status and other control variables.

	- A CSQLiteRowsetRowData object that holds that actual data for this Row.

	Handle Rows are stored in m_rgRowHandles[] and Data Rows are stored in m_rgRowData[]. Other internal
	storages are m_rgBookmarks[] that holds the bookmarks and m_rgBindings[] that holds consumer bindings.
	All those banks are defined in base classes in atldb.h.

	The objects that compose a Rowset and you need to implement are:

		CSQLiteRowset				- The actual Rowset implantation
		CSQLiteRowsetDataStorage	- The Storage of our Rowset holding all the Rows (Records)
		CSQLiteRowsetRowHandle		- A Data Row of our Rowset
		CSQLiteRowsetRowData		- A Handles Row of our Rowset
		CSQLiteColumnsRowset		- A special Metadata Rowset that holds information about columns
		CSQLiteColumnsRowsetRow		- The Row (record) of our special Metadata Rowset.

	CSQLiteRowset opens the SQLite database and reads data. It implements the minimum interfaces
	for fetching and updating data but most of the work is done by the ATL Templates in oledb.h and the
	code we need to implement is very straight forward and it has to do with SQLite access and the
	Columns Recordset.

	CSQLiteRowsetRowData which represents a single record holds the data of the columns in a 
	std::vector<string>	and exposes two important functions: read() and write() that perform writting
	and readding to the consumer buffer (a void* buffer passed from the consumer to our rowset for
	exchanging data). Both functions perform data conversion from the consumer format to the storage
	format (std::string).		  

	In general, the consumer (eg. ADO Recordset) reads and writes data to and from a Data Row, and in
	particular to one column at a time. This means that GetData() and SetData() methods both need to
	know the Row Index and Column Index in order to access the internal arrays.

	GetData(HROW hRow, HACCESSOR hAccessor, void *pData)
	SetData(HROW hRow, HACCESSOR hAccessor, void *pData)

	As you can see both GetData and SetData pass a Row Handle (hRow), an Accessor (hAccessor) and 
	a buffer pointer that is used for reading / writing data. With the passed hRow we lookup 
	m_rgRowHandles[] and retrieve the CSQLiteRowsetRowHandle object that holds the m_iRowset cursor
	which is the RowIndex. With the passed hAccessor we lookup the m_rgBindings[] and retrieve 
	hAccessor's bindings collection. Normally it contains 1 binding where pBindCur->iOrdinal-1 is
	the ColIndex but since we might have bookmarks and many bindings in hAccessor we need to loop
	and also care for bookmarks.

	You will need to hack oledb.h: find DB_E_NULLACCESSORNOTSUPPORTED and comment out the assertion
	to get writable Rowsets that support inserting new records.

	In both GetData and SetData void *pData is the consumer buffer; that is where we need to copy
	or read our records data. Data are defined by 1) a status, 2) their length and 3) their value,
	so when reading data we need to write to that buffer 3 times and when saving data we need to 
	read from that buffer 3 times. Status is simply used for handling NULL values where length and
	value should be obvious by now.

	An important aspect for reading/writing data is transforming them to and from one datatype
	to another. For example, I am using std::string vector for keeping all data so if a column
	is declared as integer I need to properly transcode the string to an integer. Doing this by
	hand is not good enough, you will eventually have to use IDataConvert::DataConvert() function.	

	These options are used to specify the characteristics of the static, keyset, and dynamic cursors defined in ODBC as follows:

	Static cursor	
	
		In a static cursor, the membership, ordering, and values of the rowset is fixed after the rowset is opened. Rows updated, deleted, or inserted after the rowset is opened are not visible to the rowset until the command is re-executed.
					
		To obtain a static cursor, the application sets the properties:	
		DBPROP_CANSCROLLBACKWARDS to VARIANT_TRUE
		DBPROP_OTHERINSERT to VARIANT_FALSE
		DBPROP_OTHERUPDATEDELETE to VARIANT_FALSE

		In ODBC, this is equivalent to specifying SQL_CURSOR_STATIC for the SQL_ATTR_CURSOR_TYPE attribute in a call to SQLSetStmtAttr.

	Keyset-driven cursor	
	
		In a keyset-driven cursor, the membership and ordering of rows in the rowset are fixed after the rowset is opened. However, values within the rows can change after the rowset is opened, including the entire row that is being deleted. Updates to a row are visible the next time the row is fetched, but rows inserted after the rowset is opened are not visible to the rowset until the command is reexecuted.
							
		To obtain a keyset-driven cursor, the application sets the properties:
							
		DBPROP_CANSCROLLBACKWARDS to VARIANT_TRUE
		DBPROP_OTHERINSERT to VARIANT_FALSE
		DBPROP_OTHERUPDATEDELETE to VARIANT_TRUE
							
		In ODBC, this is equivalent to specifying SQL_CURSOR_KEYSET_DRIVEN for the SQL_ATTR_CURSOR_TYPE attribute in a call to SQLSetStmtAttr.

	Dynamic cursor
		
		In a dynamic cursor, the membership, ordering, and values of the rowset can change after the rowset is opened. The row updated, deleted, or inserted after the rowset is opened is visible to the rowset the next time the row is fetched.
	
		To obtain a dynamic cursor, the application sets the properties:
							
		DBPROP_CANSCROLLBACKWARDS to VARIANT_TRUE
		DBPROP_OTHERINSERT to VARIANT_TRUE
		DBPROP_OTHERUPDATEDELETE to VARIANT_TRUE

		In ODBC, this is equivalent to specifying SQL_CURSOR_DYNAMIC for the SQL_ATTR_CURSOR_TYPE attribute in the call to SQLSetStmtAttr.
	
	Cursor sensitivity
	If the rowset property DBPROP_OWNINSERT is set to VARIANT_TRUE, the rowset can see its own inserts; if the rowset property DBPROP_OWNUPDATEDELETE is set to VARIANT_TRUE, the rowset can see its own updates and deletes. These are equivalent to the presence of the SQL_CASE_SENSITIVITY_ADDITIONS bit and a combination of the SQL_CASE_SENSITIVITY_UPDATES and SQL_CASE_SENSITIVITY_DELETIONS bits that are returned in the ODBC SQL_STATIC_CURSOR_ATTRIBUTES2 SQLGetInfo request.

http://msdn.microsoft.com/en-us/library/windows/desktop/ms713643(v=vs.85).aspx
http://msdn.microsoft.com/en-us/library/ms811710.aspx
http://devzone.advantagedatabase.com/dz/webhelp/Advantage7.1/mergedProjects/adsoledb/adsoledb/rowset_properties.htm
http://interested.googlecode.com/svn/trunk/SQLCEHelper/Source/DbValue.cpp
http://msdn.microsoft.com/en-us/library/windows/desktop/ms723069(v=vs.85).aspx
http://msdn.microsoft.com/en-us/library/windows/desktop/ms715968(v=vs.85).aspx
http://msdn.microsoft.com/en-us/library/windows/desktop/ms714373(v=vs.85).aspx
http://msdn.microsoft.com/en-us/library/windows/desktop/ms712925(v=vs.85).aspx

DROP TABLE IF EXISTS TEST;

CREATE TABLE TEST ( 
    ID               INTEGER           PRIMARY KEY AUTOINCREMENT
                                       NOT NULL
                                       UNIQUE,
    [BOOLEAN]        BOOLEAN,
    DATE             DATE,
    DATETIME         DATETIME,
    TIME             TIME,
    TIMESTAMP        TIMESTAMP,
    INT              INT,
    INT2             INT2,
    INT4             INT4,
    INT8             INT8,
    UINT2            UINT2,
    UINT4            UINT4,
    UINT8            UINT8,
    REAL             REAL,
    DOUBLE           DOUBLE,
    FLOAT            FLOAT,
    [DECIMAL(19,2)]  DECIMAL( 19, 2 ),
    [NUMERIC(19,2)]  NUMERIC( 19, 2 ),
    MONEY            MONEY,
    [MONEY(19,2)]    MONEY( 19, 2 ),
    TEXT             TEXT,
    CHAR             CHAR,
    VARCHAR          VARCHAR,
    [VARCHAR(2)]     VARCHAR( 2 ),
    [VARCHAR(250)]   VARCHAR( 250 ),
    [CHAR(250)]      CHAR( 250 ),
    [TEXT(250)]      TEXT( 250 ),
    Ελληνικά         VARCHAR,
    BLOB             BLOB,
    [VARBYTES(1024)] VARBYTES( 1024 ),
    IMAGE            IMAGE 
);

INSERT INTO TEST ([BOOLEAN],[Ελληνικά]) VALUES(1,'Τραγωδία');

