VERSION 5.00
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#15.3#0"; "Codejock.Controls.v15.3.1.ocx"
Begin VB.Form frm_SQLiteTest 
   Caption         =   "SQLite OLEDB Provider 1.0"
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17160
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   697
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1144
   StartUpPosition =   2  'CenterScreen
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
      ScaleWidth      =   3510
      TabIndex        =   11
      Top             =   165
      Width           =   3510
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
         TabIndex        =   13
         Top             =   0
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
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   1440
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   9690
      Left            =   150
      TabIndex        =   3
      Top             =   615
      Width           =   16845
      _Version        =   983043
      _ExtentX        =   29713
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
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.ButtonMargin=   "2,2,2,2"
      ItemCount       =   5
      Item(0).Caption =   "SQL"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "gData"
      Item(0).Control(1)=   "cmdExecute"
      Item(0).Control(2)=   "txtSQL"
      Item(0).Control(3)=   "optClientCursor"
      Item(0).Control(4)=   "optServerCursor"
      Item(0).Control(5)=   "Label2"
      Item(1).Caption =   "Tables"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gTables"
      Item(2).Caption =   "Columns"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "gColumns"
      Item(3).Caption =   "Primary Keys"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gPKs"
      Item(4).Caption =   "Foreign Keys"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "gFKs"
      Begin VB.OptionButton optServerCursor 
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
         Left            =   13365
         TabIndex        =   8
         Top             =   615
         Width           =   1440
      End
      Begin VB.OptionButton optClientCursor 
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
         Left            =   14835
         TabIndex        =   7
         Top             =   615
         Value           =   -1  'True
         Width           =   1440
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
         Text            =   "select * from CocoProjects"
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
      Begin DXDBGRIDLibCtl.dxDBGrid gData 
         Height          =   2100
         Left            =   120
         OleObjectBlob   =   "Form1.frx":038A
         TabIndex        =   4
         Top             =   1110
         Width           =   4125
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gColumns 
         Height          =   2700
         Left            =   -69865
         OleObjectBlob   =   "Form1.frx":1032
         TabIndex        =   10
         Top             =   540
         Visible         =   0   'False
         Width           =   4905
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gTables 
         Height          =   2925
         Left            =   -69790
         OleObjectBlob   =   "Form1.frx":1CDA
         TabIndex        =   14
         Top             =   615
         Visible         =   0   'False
         Width           =   5910
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gFKs 
         Height          =   2700
         Left            =   -69565
         OleObjectBlob   =   "Form1.frx":2982
         TabIndex        =   15
         Top             =   825
         Visible         =   0   'False
         Width           =   4365
      End
      Begin DXDBGRIDLibCtl.dxDBGrid gPKs 
         Height          =   2700
         Left            =   -69655
         OleObjectBlob   =   "Form1.frx":362A
         TabIndex        =   16
         Top             =   675
         Visible         =   0   'False
         Width           =   4365
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
         TabIndex        =   9
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
Private rsData As Recordset

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
    
    cmdOpenDatabase_Click
    cmdExecute_Click
End Sub

Private Sub cmdOpenDatabase_Click()

    On Error Resume Next
    
    ' Open the SQLite Database using ADO Connection
    Set Conn = New ADODB.Connection
    Conn.CursorLocation = IIf(optCnServerCursor.Value, adUseServer, adUseClient)
    Conn.ConnectionString = "Provider=SQLiteOLEDBProvider.SQLiteOLEDB.1;Data Source=D:\mobileFX\Projects\Software\SQLiteOLEDB\VBTest\team.db"
    Conn.Open
    
    ' Get Schema
    Set rsTables = Conn.OpenSchema(adSchemaTables)
    Set rsColumns = Conn.OpenSchema(adSchemaColumns)
    Set rsFKs = Conn.OpenSchema(adSchemaForeignKeys)
    
    MsgBox "Tables: " & rsTables.RecordCount & vbCrLf + _
           "Columns: " & rsColumns.RecordCount & vbCrLf + _
           "Foreign Keys: " & rsFKs.RecordCount & vbCrLf

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
        
    'Set gFKs.DataSource = rsFKs
    'gFKs.Columns.RetrieveFields
    'gFKs.Columns.ApplyBestFit Nothing
    'gFKs.ReadOnly = True
        
End Sub

Private Sub cmdExecute_Click()
    On Error Resume Next
    Set rsData = New Recordset
    rsData.CursorLocation = IIf(optServerCursor.Value, adUseServer, adUseClient)
    rsData.Open txtSQL.Text, Conn, adOpenKeyset, adLockPessimistic
    MsgBox "Records: " & rsData.RecordCount
    gData.Columns.DestroyColumns
    Set gData.DataSource = rsData
    gData.KeyField = rsData.Fields(0).Name
    gData.Columns.RetrieveFields
    gData.Columns.ApplyBestFit Nothing
    gData.ReadOnly = False
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
    gPKs.Move x, y, w, h
    
    gData.Move x, y + 40 * Screen.TwipsPerPixelY, w, h - 40 * Screen.TwipsPerPixelY
    
End Sub

Private Sub TabControl1_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next
Form_Resize
End Sub
