VERSION 5.00
Begin VB.UserControl SimpleDataSource 
   AutoRedraw      =   -1  'True
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   DataSourceBehavior=   1  'vbDataSource
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.Image Image1 
      Height          =   240
      Left            =   60
      Picture         =   "SimpleDataSource.ctx":0000
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "SimpleDataSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRS As ADODB.Recordset

Private Sub UserControl_Resize()
    On Error Resume Next
    Width = 32 * Screen.TwipsPerPixelX
    Height = 32 * Screen.TwipsPerPixelY
    Image1.Move 8, 8
    Line (0, 0)-(31, 31), vbWhite, B
    Line (31, 0)-(31, 31), vb3DDKShadow
    Line (0, 31)-(32, 31), vb3DDKShadow
End Sub

Private Sub UserControl_GetDataMember(DataMember As String, Data As Object)
    On Error Resume Next
    If Not Ambient.UserMode Then
        Set Data = Nothing
    Else
        Set Data = mRS
    End If
    Err.Clear
End Sub

Public Property Get Recordset() As ADODB.Recordset
    Set Recordset = mRS
End Property

Public Property Set Recordset(RS As ADODB.Recordset)
    On Error Resume Next
    Set mRS = RS
    PropertyChanged ""
    DataMemberChanged ""
    Err.Clear
End Property

Private Sub UserControl_Terminate()
    Set mRS = Nothing
End Sub

