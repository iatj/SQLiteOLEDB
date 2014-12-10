VERSION 5.00
Object = "{666458E7-21D8-457E-9F17-630684D7BB4C}#1.1#0"; "DataFX.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   29640
   LinkTopic       =   "Form1"
   ScaleHeight     =   11595
   ScaleWidth      =   29640
   StartUpPosition =   3  'Windows Default
   Begin DataFX.DataManager DataManager3 
      Left            =   13560
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      SQL             =   "SEL_COCOTASKS"
      SQLManager      =   "SQLManager1"
      ConnectionString=   ""
      StartupMethod   =   1
      DetectBinders   =   -1  'True
      UpdateOption    =   0
      AutoRun         =   -1  'True
      UpdateCriteria  =   0
      PopupMenus      =   0   'False
      SuppressMessages=   -1  'True
      FieldsInfo      =   $"Form2.frx":0000
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid3 
      Bindings        =   "Form2.frx":2917
      Height          =   5190
      Left            =   12720
      OleObjectBlob   =   "Form2.frx":2932
      TabIndex        =   2
      Top             =   5355
      Width           =   15990
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid2 
      Height          =   4830
      Left            =   750
      OleObjectBlob   =   "Form2.frx":35D2
      TabIndex        =   1
      Top             =   5235
      Width           =   11460
   End
   Begin DataFX.DataManager DataManager2 
      Left            =   12795
      Top             =   1530
      _ExtentX        =   847
      _ExtentY        =   847
      SQL             =   "SEL_COCOPROJECTS"
      SQLManager      =   "SQLManager1"
      ConnectionString=   ""
      StartupMethod   =   1
      DetectBinders   =   -1  'True
      UpdateOption    =   0
      AutoRun         =   -1  'True
      UpdateCriteria  =   0
      PopupMenus      =   0   'False
      SuppressMessages=   -1  'True
      FieldsInfo      =   $"Form2.frx":68CC
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Bindings        =   "Form2.frx":847A
      Height          =   1665
      Left            =   90
      OleObjectBlob   =   "Form2.frx":8495
      TabIndex        =   0
      Top             =   3300
      Width           =   29085
   End
   Begin DataFX.DataManager DataManager1 
      Left            =   11925
      Top             =   1515
      _ExtentX        =   847
      _ExtentY        =   847
      SQL             =   "SEL_TEST"
      SQLManager      =   "SQLManager1"
      ConnectionString=   ""
      StartupMethod   =   1
      DetectBinders   =   -1  'True
      UpdateOption    =   0
      AutoRun         =   -1  'True
      UpdateCriteria  =   0
      SuppressMessages=   -1  'True
      FieldsInfo      =   $"Form2.frx":913D
   End
   Begin DataFX.SQLManager SQLManager1 
      Left            =   11325
      Top             =   1575
      _ExtentX        =   847
      _ExtentY        =   847
      DataLayer       =   0
      SQLXml          =   $"Form2.frx":11F3F
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


