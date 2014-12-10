VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5B44EC52-B95B-45CF-98FF-A49DFEED5A92}#15.3#0"; "Codejock.PropertyGrid.v15.3.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.Controls.v15.3.1.ocx"
Begin VB.Form frm_SQLiteTest 
   Caption         =   "SQLite OLEDB Provider 1.0"
   ClientHeight    =   16515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   29940
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_SQLiteTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1101
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1996
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin SQLiteOLEDBTests.SimpleDataSource SimpleDataSource1 
      Left            =   20040
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   14280
      ScaleHeight     =   360
      ScaleWidth      =   14175
      TabIndex        =   9
      Top             =   165
      Width           =   14175
      Begin VB.OptionButton optCnServerCursor 
         Caption         =   "Server Cursor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Client Cursor"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         Top             =   0
         Width           =   1440
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9690
      Left            =   135
      TabIndex        =   3
      Top             =   630
      Width           =   18735
      _Version        =   983043
      _ExtentX        =   33046
      _ExtentY        =   17092
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DrawFocusRect   =   0   'False
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.ButtonMargin=   "2,2,2,2"
      ItemCount       =   4
      Item(0).Caption =   "SQL"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "cmdExecute"
      Item(0).Control(1)=   "txtSQL"
      Item(0).Control(2)=   "Label2"
      Item(0).Control(3)=   "rsProps"
      Item(0).Control(4)=   "fldProps"
      Item(0).Control(5)=   "cmdUpdate"
      Item(0).Control(6)=   "gDataClient"
      Item(0).Control(7)=   "gDataServer"
      Item(0).Control(8)=   "BlobFields"
      Item(0).Control(9)=   "optClientCursor"
      Item(0).Control(10)=   "optServerCursor"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gTables"
      Item(2).Caption =   "Columns"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gColumns"
      Item(3).Caption =   "Foreign Keys"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gFKs"
      Begin XtremeSuiteControls.TabControl BlobFields 
         Height          =   3945
         Left            =   5055
         TabIndex        =   20
         Top             =   1845
         Width           =   4110
         _Version        =   983043
         _ExtentX        =   7250
         _ExtentY        =   6959
         _StockProps     =   68
         DrawFocusRect   =   0   'False
         Appearance      =   8
         Color           =   4
         PaintManager.BoldSelected=   -1  'True
         PaintManager.HotTracking=   -1  'True
         PaintManager.ButtonMargin=   "2,2,2,2"
         Begin SQLiteOLEDBTests.BoundHexEditCtl HexEdit1 
            DataSource      =   "SimpleDataSource1"
            Height          =   3195
            Left            =   45
            TabIndex        =   21
            Top             =   465
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   5636
         End
      End
      Begin VB.OptionButton optClientCursor 
         Caption         =   "Client Cursor  Recordset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   16395
         TabIndex        =   19
         Top             =   585
         Width           =   1440
      End
      Begin VB.OptionButton optServerCursor 
         Caption         =   "Server Cursor Recordset"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   14730
         TabIndex        =   18
         Top             =   585
         Value           =   -1  'True
         Width           =   1440
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13305
         TabIndex        =   16
         Top             =   585
         Width           =   1215
      End
      Begin XtremePropertyGrid.PropertyGrid rsProps 
         Height          =   3615
         Left            =   9615
         TabIndex        =   14
         Top             =   1215
         Width           =   7125
         _Version        =   983043
         _ExtentX        =   12568
         _ExtentY        =   6376
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ViewBackColor   =   -2147483633
         LineColor       =   12632256
         ToolBarVisible  =   -1  'True
         HelpVisible     =   0   'False
         PropertySort    =   1
         HighlightChangedButtonItems=   0   'False
         VisualTheme     =   6
         HideSelection   =   -1  'True
      End
      Begin VB.TextBox txtSQL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   735
         TabIndex        =   6
         Text            =   "SELECT * FROM TEST"
         Top             =   600
         Width           =   11265
      End
      Begin VB.CommandButton cmdExecute 
         Caption         =   "Execute"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12030
         TabIndex        =   5
         Top             =   585
         Width           =   1215
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDataServer 
         Height          =   2100
         Left            =   90
         OleObjectBlob   =   "frm_SQLiteTest.frx":038A
         TabIndex        =   4
         Top             =   1095
         Width           =   4125
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gColumns 
         Height          =   2700
         Left            =   -69865
         OleObjectBlob   =   "frm_SQLiteTest.frx":1032
         TabIndex        =   8
         Top             =   540
         Visible         =   0   'False
         Width           =   4905
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gTables 
         Height          =   2925
         Left            =   -69790
         OleObjectBlob   =   "frm_SQLiteTest.frx":1CDA
         TabIndex        =   12
         Top             =   615
         Visible         =   0   'False
         Width           =   5910
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gFKs 
         Height          =   2700
         Left            =   -69565
         OleObjectBlob   =   "frm_SQLiteTest.frx":2982
         TabIndex        =   13
         Top             =   825
         Visible         =   0   'False
         Width           =   4365
      End
      Begin XtremePropertyGrid.PropertyGrid fldProps 
         Height          =   3615
         Left            =   9615
         TabIndex        =   15
         Top             =   5070
         Width           =   7125
         _Version        =   983043
         _ExtentX        =   12568
         _ExtentY        =   6376
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ViewBackColor   =   -2147483633
         LineColor       =   12632256
         ViewCategoryForeColor=   8388608
         ToolBarVisible  =   -1  'True
         HelpVisible     =   0   'False
         PropertySort    =   0
         HighlightChangedButtonItems=   0   'False
         VisualTheme     =   6
         HideSelection   =   -1  'True
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gDataClient 
         Height          =   2100
         Left            =   120
         OleObjectBlob   =   "frm_SQLiteTest.frx":362A
         TabIndex        =   17
         Top             =   3510
         Width           =   4125
      End
      Begin VB.Label Label2 
         Caption         =   "SQL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   161
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   7
         Top             =   660
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdOpenDatabase 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12885
      TabIndex        =   2
      Top             =   105
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1590
      TabIndex        =   1
      Text            =   "Provider=SQLiteOLEDBProvider.SQLiteOLEDB.1;Data Source=D:\mobileFX\Projects\Software\SQLiteOLEDB\VBTest\team.db"
      Top             =   120
      Width           =   11265
   End
   Begin VB.Label Label1 
      Caption         =   "Connection String:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   1815
   End
End
Attribute VB_Name = "frm_SQLiteTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Conn As ADODB.Connection
Private rsTables As Recordset
Private rsColumns As Recordset
Private rsFKs As Recordset
Private rsDataServer As Recordset
Private rsDataClient As Recordset
Private dic As New Dictionary

Public Enum DBCOLUMNFLAGS
    DBCOLUMNFLAGS_CACHEDEFERRED = &H1000&
    DBCOLUMNFLAGS_ISBOOKMARK = &H1&
    DBCOLUMNFLAGS_ISCHAPTER = &H2000&
    DBCOLUMNFLAGS_ISCOLLECTION = &H40000
    DBCOLUMNFLAGS_ISDEFAULTSTREAM = &H20000
    DBCOLUMNFLAGS_ISFIXEDLENGTH = &H10&
    DBCOLUMNFLAGS_ISLONG = &H80&
    DBCOLUMNFLAGS_ISNULLABLE = &H20&
    DBCOLUMNFLAGS_ISROW = &H200000
    DBCOLUMNFLAGS_ISROWID = &H100&
    DBCOLUMNFLAGS_ISROWSET = &H100000
    DBCOLUMNFLAGS_ISROWURL = &H10000
    DBCOLUMNFLAGS_ISROWVER = &H200&
    DBCOLUMNFLAGS_ISSTREAM = &H80000
    DBCOLUMNFLAGS_KEYCOLUMN = &H8000&
    DBCOLUMNFLAGS_MAYBENULL = &H40&
    DBCOLUMNFLAGS_MAYDEFER = &H2&
    DBCOLUMNFLAGS_RESERVED = &H8000&
    DBCOLUMNFLAGS_ROWSPECIFICCOLUMN = &H400000
    DBCOLUMNFLAGS_SCALEISNEGATIVE = &H4000&
    DBCOLUMNFLAGS_WRITE = &H4&
    DBCOLUMNFLAGS_WRITEUNKNOWN = &H8&
End Enum

' ==================================================================================================================================
'       ______
'      / ____/___  _________ ___
'     / /_  / __ \/ ___/ __ `__ \
'    / __/ / /_/ / /  / / / / / /
'   /_/    \____/_/  /_/ /_/ /_/
'
' ==================================================================================================================================

Private Sub Form_Load()
    On Error Resume Next
    
    Dim ctl As Control
    Dim dx As dxDBGrid
    
    For Each ctl In Controls
        If TypeOf ctl Is dxDBGrid Then
            Set dx = ctl
            Set dx.Font = Me.Font
            Set dx.Columns.HeaderFont = Me.Font
        End If
    Next
    
    dic.Add "adArray", adArray&
    dic.Add "adBigInt", adBigInt&
    dic.Add "adBinary", adBinary&
    dic.Add "adBoolean", adBoolean&
    dic.Add "adBSTR", adBSTR&
    dic.Add "adChapter", adChapter&
    dic.Add "adChar", adChar&
    dic.Add "adCurrency", adCurrency&
    dic.Add "adDate", adDate&
    dic.Add "adDBDate", adDBDate&
    dic.Add "adDBTime", adDBTime&
    dic.Add "adDBTimeStamp", adDBTimeStamp&
    dic.Add "adDecimal", adDecimal&
    dic.Add "adDouble", adDouble&
    dic.Add "adEmpty", adEmpty&
    dic.Add "adError", adError&
    dic.Add "adFileTime", adFileTime&
    dic.Add "adGUID", adGUID&
    dic.Add "adIDispatch", adIDispatch&
    dic.Add "adInteger", adInteger&
    dic.Add "adIUnknown", adIUnknown&
    dic.Add "adLongVarBinary", adLongVarBinary&
    dic.Add "adLongVarChar", adLongVarChar&
    dic.Add "adLongVarWChar", adLongVarWChar&
    dic.Add "adNumeric", adNumeric&
    dic.Add "adPropVariant", adPropVariant&
    dic.Add "adSingle", adSingle&
    dic.Add "adSmallInt", adSmallInt&
    dic.Add "adTinyInt", adTinyInt&
    dic.Add "adUnsignedBigInt", adUnsignedBigInt&
    dic.Add "adUnsignedInt", adUnsignedInt&
    dic.Add "adUnsignedSmallInt", adUnsignedSmallInt&
    dic.Add "adUnsignedTinyInt", adUnsignedTinyInt&
    dic.Add "adUserDefined", adUserDefined&
    dic.Add "adVarBinary", adVarBinary&
    dic.Add "adVarChar", adVarChar&
    dic.Add "adVariant", adVariant&
    dic.Add "adVarNumeric", adVarNumeric&
    dic.Add "adVarWChar", adVarWChar&
    dic.Add "adWChar", adWChar&
    
    OpenSQLiteDB
    cmdExecute_Click
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim x As Long
    Dim y As Long
    Dim w As Long
    Dim h As Long
    
    TabControl1.Move 4, 40, ScaleWidth - 8, ScaleHeight - 48
    
    x = TabControl1.ClientLeft + 4 * Screen.TwipsPerPixelX
    y = (30) * Screen.TwipsPerPixelY
    w = (TabControl1.ClientWidth - 8) * Screen.TwipsPerPixelX
    h = (TabControl1.ClientHeight * Screen.TwipsPerPixelY) - y
    
    gTables.Move x, y, w, h
    gColumns.Move x, y, w, h
    gFKs.Move x, y, w, h
    
    w = w - fldProps.Width - HexEdit1.Width
    
    y = y + 40 * Screen.TwipsPerPixelY
    h = h - 20 * Screen.TwipsPerPixelY
    
    fldProps.Left = w + 8 * Screen.TwipsPerPixelX + HexEdit1.Width
    rsProps.Left = w + 8 * Screen.TwipsPerPixelX + HexEdit1.Width
    rsProps.Top = y
    rsProps.Height = h / 2 - 4 * Screen.TwipsPerPixelY
    fldProps.Height = h / 2
    fldProps.Top = y + h / 2
    gDataServer.Move x, y, w, h / 2
    gDataClient.Move x, y + h / 2, w, h / 2
    
    BlobFields.Move w + 6 * Screen.TwipsPerPixelX, y, HexEdit1.Width, h
    HexEdit1.Move 8, 380, BlobFields.ClientWidth, BlobFields.ClientHeight
    
    gColumns.GridLineColor = vbButtonFace

End Sub

' ==================================================================================================================================
'      ____                      _____ ____    __    _ __
'     / __ \____  ___  ____     / ___// __ \  / /   (_) /____
'    / / / / __ \/ _ \/ __ \    \__ \/ / / / / /   / / __/ _ \
'   / /_/ / /_/ /  __/ / / /   ___/ / /_/ / / /___/ / /_/  __/
'   \____/ .___/\___/_/ /_/   /____/\___\_\/_____/_/\__/\___/
'       /_/
' ==================================================================================================================================

Private Sub cmdOpenDatabase_Click()
    OpenSQLiteDB
    MsgBox "Database Openned!", vbInformation
End Sub

Private Sub OpenSQLiteDB()
    
    On Error Resume Next
    
    CloseSQLiteDB
        
    ' Open the SQLite Database using ADO Connection
    Set Conn = New ADODB.Connection
    Conn.CursorLocation = IIf(optCnServerCursor.Value, adUseServer, adUseClient)
    Conn.ConnectionString = "Provider=SQLiteOLEDBProvider.SQLiteOLEDB.1;Data Source=D:\mobileFX\Projects\Software\SQLiteOLEDB\VBTest\team.db"
    Conn.Open
    
    ' Get Schema
    Set rsTables = ClientRecordset(Conn.OpenSchema(adSchemaTables))
    Set rsColumns = ClientRecordset(Conn.OpenSchema(adSchemaColumns))
    Set rsFKs = ClientRecordset(Conn.OpenSchema(adSchemaForeignKeys))
        
'    Set rsTables = Conn.OpenSchema(adSchemaTables)
'    Set rsColumns = Conn.OpenSchema(adSchemaColumns)
'    Set rsFKs = Conn.OpenSchema(adSchemaForeignKeys)
        
    Set gTables.DataSource = rsTables
    gTables.KeyField = "TABLE_GUID"
    gTables.Columns.RetrieveFields
    gTables.Columns.ApplyBestFit Nothing
    gTables.ReadOnly = True

    Set gColumns.DataSource = rsColumns
    gColumns.KeyField = "COLUMN_GUID"
    gColumns.Columns.RetrieveFields
    gColumns.Columns.ApplyBestFit Nothing
    gColumns.ReadOnly = True
    gColumns.M.AddGroupColumn gColumns.Columns.ColumnByFieldName("TABLE_NAME")
    gColumns.M.FullExpand
        
    Set gFKs.DataSource = rsFKs
    gFKs.KeyField = "PK_COLUMN_GUID"
    gFKs.Columns.RetrieveFields
    gFKs.Columns.ApplyBestFit Nothing
    gFKs.ReadOnly = True
    
    optCnServerCursor.Value = (Conn.CursorLocation = adUseServer)
    
End Sub

Private Function ClientRecordset(ByVal RS As ADODB.Recordset) As ADODB.Recordset
    Dim Stream As ADODB.Stream
    Set Stream = New ADODB.Stream
    Stream.Type = adTypeBinary
    RS.Save Stream, adPersistXML
    Set RS = New ADODB.Recordset
    RS.Open Stream
    Set ClientRecordset = RS
End Function

Private Sub CloseSQLiteDB()
   
   If Not Conn Is Nothing Then
        
        Set gTables.DataSource = Nothing
        Set gColumns.DataSource = Nothing
        Set gFKs.DataSource = Nothing
        Set gDataClient.DataSource = Nothing
        Set gDataServer.DataSource = Nothing
        
        gTables.Columns.DestroyColumns
        gColumns.Columns.DestroyColumns
        gFKs.Columns.DestroyColumns
        gDataClient.Columns.DestroyColumns
        gDataServer.Columns.DestroyColumns
                
        Set rsTables = Nothing
        Set rsColumns = Nothing
        Set rsFKs = Nothing
        Set rsDataServer = Nothing
        Set rsDataClient = Nothing
        
        Conn.Close
        Set Conn = Nothing
        
    End If

End Sub

' ==================================================================================================================================
'       ______                      __          ____
'      / ____/  _____  ____ ___  __/ /____     / __ \__  _____  _______  __
'     / __/ | |/_/ _ \/ __ `/ / / / __/ _ \   / / / / / / / _ \/ ___/ / / /
'    / /____>  </  __/ /_/ / /_/ / /_/  __/  / /_/ / /_/ /  __/ /  / /_/ /
'   /_____/_/|_|\___/\__, /\__,_/\__/\___/   \___\_\__,_/\___/_/   \__, /
'                      /_/                                        /____/
' ==================================================================================================================================

Private Sub cmdExecute_Click()
    On Error Resume Next
    
    Set gDataServer.DataSource = Nothing
    gDataServer.Columns.DestroyColumns
    Set rsDataServer = New Recordset
    rsDataServer.CursorLocation = adUseServer
    rsDataServer.Open txtSQL.Text, Conn, adOpenKeyset, adLockBatchOptimistic
    gDataServer.Dataset.DisableControls
    Set gDataServer.DataSource = rsDataServer
    gDataServer.KeyField = rsDataServer.Fields(0).Name
    gDataServer.Columns.RetrieveFields
    gDataServer.Columns.ApplyBestFit Nothing
    gDataServer.ReadOnly = False
    gDataServer.Dataset.EnableControls
    
    Set rsDataClient = New Recordset
    rsDataClient.CursorLocation = adUseClient
    rsDataClient.Open txtSQL.Text, Conn, adOpenKeyset, adLockBatchOptimistic
    Set gDataClient.DataSource = Nothing
    gDataClient.Columns.DestroyColumns
    Set gDataClient.DataSource = rsDataClient
    gDataClient.KeyField = rsDataClient.Fields(0).Name
    gDataClient.Columns.RetrieveFields
    gDataClient.Columns.ApplyBestFit Nothing
    gDataClient.ReadOnly = False
    
    RefreshUI
       
End Sub

' ==================================================================================================================================
'      __  __          __      __
'     / / / /___  ____/ /___ _/ /____
'    / / / / __ \/ __  / __ `/ __/ _ \
'   / /_/ / /_/ / /_/ / /_/ / /_/  __/
'   \____/ .___/\__,_/\__,_/\__/\___/
'       /_/
' ==================================================================================================================================

Private Sub cmdUpdate_Click()
    Dim RS As Recordset
    On Error Resume Next
    Set RS = IIf(optServerCursor.Value, rsDataServer, rsDataClient)
    If RS.ActiveConnection Is Nothing Then
        Set RS.ActiveConnection = Conn
    End If
    RS.UpdateBatch
End Sub


' ==================================================================================================================================
'       ___    ____  ____     ____                             __
'      /   |  / __ \/ __ \   / __ \_________  ____  ___  _____/ /_(_)__  _____
'     / /| | / / / / / / /  / /_/ / ___/ __ \/ __ \/ _ \/ ___/ __/ / _ \/ ___/
'    / ___ |/ /_/ / /_/ /  / ____/ /  / /_/ / /_/ /  __/ /  / /_/ /  __(__  )
'   /_/  |_/_____/\____/  /_/   /_/   \____/ .___/\___/_/   \__/_/\___/____/
'                                         /_/
' ==================================================================================================================================

Private Sub RefreshUI()
    gDataServer.Enabled = optServerCursor.Value
    gDataClient.Enabled = optClientCursor.Value
    gDataServer.BackColor = IIf(gDataServer.Enabled, vbWhite, vbButtonFace)
    gDataClient.BackColor = IIf(gDataClient.Enabled, vbWhite, vbButtonFace)
    ReadProperties IIf(optServerCursor.Value, rsDataServer, rsDataClient)
    CreateBlobTabs
    HideBlobsInGrid
    BindBlobField
End Sub

Private Sub optServerCursor_Click()
    RefreshUI
End Sub

Private Sub optClientCursor_Click()
    RefreshUI
End Sub

Private Sub ReadProperties(ByVal rsDataServer As Recordset)
    
    Dim p As ADODB.Property
    Dim n As PropertyGridItem
    Dim g As PropertyGridItem
    Dim c As PropertyGridItem
    Dim f As ADODB.Field
    Dim ctx As PropertyGridUpdateContext
    Dim i As Long
    
    Set ctx = rsProps.BeginUpdate
    rsProps.Categories.Clear
    Set g = rsProps.AddCategory("Dataset Properties")
    For Each p In rsDataServer.Properties
        If p.Type = adBoolean Then
            Set c = g.AddChildItem(PropertyItemBool, p.Name, p.Value)
            If p.Value = False Then
                c.ValueMetrics.ForeColor = vbRed
            End If
        Else
            g.AddChildItem PropertyItemString, p.Name, p.Value
        End If
    Next
    g.Expanded = True
    rsProps.SplitterPos = 0.7
    rsProps.EndUpdate ctx
            
    Set ctx = fldProps.BeginUpdate
    fldProps.Categories.Clear
    For Each f In rsDataServer.Fields
    
        Set g = fldProps.AddCategory("Field: " + f.Name)
        
        For i = 0 To dic.Count - 1
            If dic(dic.Keys(i)) = f.Type Then
                g.Caption = g.Caption + " - " + dic.Keys(i) + "(" & Join(Array(f.DefinedSize, f.Precision, f.NumericScale), ", ") & ")"
            End If
        Next
    
        Set n = g.AddChildItem(PropertyItemString, "Name", f.Name)
        n.ValueMetrics.Font.Bold = True
        n.CaptionMetrics.Font.Bold = True
        
        Set c = g.AddChildItem(PropertyItemNumber, "Type", f.Type)
        Set c = g.AddChildItem(PropertyItemEnum, "Type", f.Type)
        For i = 0 To dic.Count - 1
            c.Constraints.Add dic.Keys(i), dic(dic.Keys(i))
        Next
        c.ValueMetrics.Font.Bold = True
        c.CaptionMetrics.Font.Bold = True
        
        Set c = g.AddChildItem(PropertyItemString, "Attributes", f.Attributes)
        Set c = g.AddChildItem(PropertyItemEnumFlags, "Attributes", f.Attributes)
        c.Constraints.Add "CACHEDEFERRED", DBCOLUMNFLAGS_CACHEDEFERRED
        c.Constraints.Add "ISBOOKMARK", DBCOLUMNFLAGS_ISBOOKMARK
        c.Constraints.Add "ISCHAPTER", DBCOLUMNFLAGS_ISCHAPTER
        c.Constraints.Add "ISCOLLECTION", DBCOLUMNFLAGS_ISCOLLECTION
        c.Constraints.Add "ISDEFAULTSTREAM", DBCOLUMNFLAGS_ISDEFAULTSTREAM
        c.Constraints.Add "ISFIXEDLENGTH", DBCOLUMNFLAGS_ISFIXEDLENGTH
        c.Constraints.Add "ISLONG", DBCOLUMNFLAGS_ISLONG
        c.Constraints.Add "ISNULLABLE", DBCOLUMNFLAGS_ISNULLABLE
        c.Constraints.Add "ISROW", DBCOLUMNFLAGS_ISROW
        c.Constraints.Add "ISROWID", DBCOLUMNFLAGS_ISROWID
        c.Constraints.Add "ISROWSET", DBCOLUMNFLAGS_ISROWSET
        c.Constraints.Add "ISROWURL", DBCOLUMNFLAGS_ISROWURL
        c.Constraints.Add "ISROWVER", DBCOLUMNFLAGS_ISROWVER
        c.Constraints.Add "ISSTREAM", DBCOLUMNFLAGS_ISSTREAM
        c.Constraints.Add "KEYCOLUMN", DBCOLUMNFLAGS_KEYCOLUMN
        c.Constraints.Add "MAYBENULL", DBCOLUMNFLAGS_MAYBENULL
        c.Constraints.Add "MAYDEFER", DBCOLUMNFLAGS_MAYDEFER
        c.Constraints.Add "RESERVED", DBCOLUMNFLAGS_RESERVED
        c.Constraints.Add "ROWSPECIFICCOLUMN", DBCOLUMNFLAGS_ROWSPECIFICCOLUMN
        c.Constraints.Add "SCALEISNEGATIVE", DBCOLUMNFLAGS_SCALEISNEGATIVE
        c.Constraints.Add "WRITE", DBCOLUMNFLAGS_WRITE
        c.Constraints.Add "WRITEUNKNOWN", DBCOLUMNFLAGS_WRITEUNKNOWN
        
        Set c = g.AddChildItem(PropertyItemString, "DataFormat", f.DataFormat)
        Set c = g.AddChildItem(PropertyItemString, "DefinedSize", f.DefinedSize)
        Set c = g.AddChildItem(PropertyItemString, "NumericScale", f.NumericScale)
        Set c = g.AddChildItem(PropertyItemString, "Precision", f.Precision)
        Set c = g.AddChildItem(PropertyItemString, "Status", f.Status)
                        
        For Each p In f.Properties
            If p.Type = adBoolean Then
                Set c = g.AddChildItem(PropertyItemBool, p.Name, p.Value)
                If p.Value = False Then
                    c.ValueMetrics.ForeColor = vbRed
                End If
            Else
                g.AddChildItem PropertyItemString, p.Name, p.Value
            End If
        Next
                
    Next
    fldProps.EndUpdate ctx
                
End Sub

' ==================================================================================================================================
'       ____  __      __
'      / __ )/ /___  / /_  _____
'     / __  / / __ \/ __ \/ ___/
'    / /_/ / / /_/ / /_/ (__  )
'   /_____/_/\____/_.___/____/
'
' ==================================================================================================================================

Private Sub CreateBlobTabs()
    Dim i As Long
    Dim f As Field
    BlobFields.RemoveAll
    For Each f In rsDataServer.Fields
        If f.Type = adVariant Then
            i = i + 1
            If i = 1 Then
                BlobFields.InsertItem(i, f.Name, 0, 0).Selected = True
            Else
                BlobFields.InsertItem i, f.Name, 0, 0
            End If
        End If
    Next
End Sub

Private Sub HideBlobsInGrid()
    Dim f As Field
    For Each f In rsDataServer.Fields
       If f.Type = adVariant Then
            If Not gDataServer.Columns.ColumnByFieldName(f.Name) Is Nothing Then
                gDataServer.Columns.ColumnByFieldName(f.Name).Visible = False
            End If
            If Not gDataClient.Columns.ColumnByFieldName(f.Name) Is Nothing Then
                gDataClient.Columns.ColumnByFieldName(f.Name).Visible = False
            End If
        End If
    Next
End Sub

Private Sub BindBlobField()
    If BlobFields.Selected Is Nothing Then Exit Sub
    HexEdit1.DataField = BlobFields.Selected.Caption
    Set SimpleDataSource1.Recordset = IIf(optServerCursor.Value, rsDataServer, rsDataClient)
    BlobFields.AutoResizeClient = True
    HexEdit1.Move 4, 345 + BlobFields.Selected.Index * 22 * Screen.TwipsPerPixelY, BlobFields.ClientWidth, BlobFields.ClientHeight - 45
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    On Error Resume Next
    Form_Resize
End Sub

Private Sub BlobFields_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    BindBlobField
End Sub


