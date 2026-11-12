VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFilterTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WorkbookName As String
Public SheetName As String
Public SourceKind As String
Public TableName As String
Public FieldIndex As Long
Public SheetColumn As Long
Public ColumnLetter As String
Public HeaderText As String
Public FilterRangeAddress As String