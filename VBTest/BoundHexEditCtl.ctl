VERSION 5.00
Object = "{93966622-AF6F-495B-87D7-88C922C9988B}#1.0#0"; "HexEdit.ocx"
Begin VB.UserControl BoundHexEditCtl 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   ClipBehavior    =   0  'None
   DataBindingBehavior=   1  'vbSimpleBound
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   Begin HEXEDITLib.HexEdit HexEdit1 
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   3930
      _Version        =   65536
      _ExtentX        =   6932
      _ExtentY        =   4339
      _StockProps     =   137
      ForeColor       =   8388608
      BackColor       =   16777215
      ShowAddress     =   0   'False
      AllowChangeSize =   -1  'True
   End
End
Attribute VB_Name = "BoundHexEditCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
    On Error Resume Next
    HexEdit1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Public Property Get Value() As Variant
Attribute Value.VB_MemberFlags = "123c"
    On Error Resume Next
    If Not Ambient.UserMode Then Exit Property
    Value = HexEdit1.GetData
End Property

Public Property Let Value(ByVal v As Variant)
    On Error Resume Next
    HexEdit1.SetData v
    PropertyChanged "Value"
End Property

Private Sub HexEdit1_LostFocus()
    PropertyChanged "Value"
End Sub

Private Sub HexEdit1_Changed()
    PropertyChanged "Value"
End Sub


