VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AF_FilterTarget"
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
Public TargetSheet As Worksheet
Public Remembered As Boolean
Public CanReapply As Boolean
Public FilterOperator As Long
Public HasCriteria1 As Boolean
Public HasCriteria2 As Boolean
Public Criteria1 As Variant
Public Criteria2 As Variant
