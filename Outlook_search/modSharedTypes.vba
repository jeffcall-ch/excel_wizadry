' ============================================
' Module: modSharedTypes
' Purpose: Shared data types for AQS email search
' Version: 2.0.0
' ============================================

Option Explicit

' Search criteria structure for AQS implementation
Public Type SearchCriteria
    FromAddress As String
    ToAddress As String
    CcAddress As String
    Subject As String
    Body As String
    SearchTerms As String
    DateFrom As Date
    DateTo As Date
    UnreadOnly As Boolean
    ImportantOnly As Boolean
    SearchSubFolders As Boolean  ' Include subfolders in search
    AttachmentFilter As String
    SizeFilter As String
    FolderName As String
End Type
