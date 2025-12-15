' ============================================
' Module: modSharedTypes (Standard Module)
' Purpose: Shared type definitions for email search
' Version: 1.1
' ============================================

Option Explicit

' Type Definitions
Public Type SearchCriteria
    FromAddress As String
    SearchTerms As String
    SearchSubject As Boolean
    SearchBody As Boolean
    DateFrom As String
    DateTo As String
    FolderName As String
    CaseSensitive As Boolean
    ToAddress As String
    CcAddress As String
    AttachmentFilter As String
    SizeFilter As String
    UnreadOnly As Boolean
    ImportantOnly As Boolean
End Type
