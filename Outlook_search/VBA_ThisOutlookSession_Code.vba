' ============================================
' Module: ThisOutlookSession
' Purpose: Main search logic using AdvancedSearch API
' ============================================

Option Explicit

' Type definition for search criteria
Public Type SearchCriteria
    FromText As String
    SubjectText As String
    SearchBody As Boolean
    DateFrom As String
    DateTo As String
    FolderName As String
End Type

' Module-level variable to track search object
Private WithEvents m_SearchObject As Outlook.Search
Private m_SearchStartTime As Double

' ============================================
' PUBLIC METHODS
' ============================================

Public Sub ShowEmailSearchForm()
    ' Entry point - shows the search form
    ' Add this macro to Quick Access Toolbar
    frmEmailSearch.Show
End Sub

Public Sub ExecuteEmailSearch(criteria As SearchCriteria)
    ' Execute search using Outlook's AdvancedSearch API
    On Error GoTo ErrorHandler
    
    ' Start timing
    m_SearchStartTime = Timer
    
    ' Build DASL filter
    Dim filter As String
    filter = BuildDASLFilter(criteria)
    
    If filter = "" Then
        MsgBox "No valid search criteria specified.", vbExclamation, "Search Error"
        Exit Sub
    End If
    
    ' Determine scope (folders to search)
    Dim scope As String
    scope = GetSearchScope(criteria.FolderName)
    
    ' Show status
    Application.ActiveExplorer.WindowState = olMaximized
    
    ' Execute AdvancedSearch (uses Windows Search index - FAST!)
    Set m_SearchObject = Application.AdvancedSearch( _
        scope:=scope, _
        filter:=filter, _
        SearchSubFolders:=True, _
        Tag:="EmailSearchResults")
    
    ' Results will appear when search completes (see m_SearchObject_Complete event)
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Search error: " & Err.Description, vbCritical, "Search Failed"
End Sub

' ============================================
' EVENT HANDLERS
' ============================================

Private Sub m_SearchObject_Complete(ByVal SearchObject As Search)
    ' Called when search finishes (usually < 3 seconds!)
    On Error Resume Next
    
    Dim searchTime As Double
    searchTime = Timer - m_SearchStartTime
    
    Dim resultCount As Long
    resultCount = SearchObject.Results.Count
    
    ' Save results to a Search Folder
    SearchObject.Save "Email Search Results"
    
    ' Navigate to the Search Folder
    Dim searchFolder As Outlook.Folder
    Set searchFolder = Application.Session.GetDefaultFolder(olFolderInbox).Parent.Folders("Search Folders")
    
    ' Find and display the results folder
    Dim resultFolder As Outlook.Folder
    On Error Resume Next
    Set resultFolder = searchFolder.Folders("Email Search Results")
    On Error GoTo 0
    
    If Not resultFolder Is Nothing Then
        Application.ActiveExplorer.SelectFolder resultFolder
        Application.ActiveExplorer.CurrentFolder.ShowAsOutlookAB
    End If
    
    ' Show completion message
    MsgBox "Search complete!" & vbCrLf & vbCrLf & _
           "Found: " & resultCount & " emails" & vbCrLf & _
           "Time: " & Format(searchTime, "0.00") & " seconds" & vbCrLf & vbCrLf & _
           "Results displayed in 'Email Search Results' folder.", _
           vbInformation, "Search Complete"
End Sub

' ============================================
' HELPER FUNCTIONS
' ============================================

Private Function BuildDASLFilter(criteria As SearchCriteria) As String
    ' Build DASL query filter from search criteria
    ' DASL = DAV Searching and Locating (uses Windows Search index)
    
    Dim filters() As String
    Dim filterCount As Integer
    filterCount = 0
    ReDim filters(0 To 10)
    
    ' From field
    If criteria.FromText <> "" Then
        ' Search in sender name AND email address
        Dim fromFilter As String
        fromFilter = "(" & _
            "@SQL=""urn:schemas:httpmail:fromname"" LIKE '%" & EscapeSQL(criteria.FromText) & "%' OR " & _
            "@SQL=""urn:schemas:httpmail:fromemail"" LIKE '%" & EscapeSQL(criteria.FromText) & "%'" & _
            ")"
        filters(filterCount) = fromFilter
        filterCount = filterCount + 1
    End If
    
    ' Subject field
    If criteria.SubjectText <> "" Then
        If criteria.SearchBody Then
            ' Search in subject OR body
            Dim textFilter As String
            textFilter = "(" & _
                "@SQL=""urn:schemas:httpmail:subject"" LIKE '%" & EscapeSQL(criteria.SubjectText) & "%' OR " & _
                "@SQL=""urn:schemas:httpmail:textdescription"" LIKE '%" & EscapeSQL(criteria.SubjectText) & "%'" & _
                ")"
            filters(filterCount) = textFilter
        Else
            ' Subject only
            filters(filterCount) = "@SQL=""urn:schemas:httpmail:subject"" LIKE '%" & EscapeSQL(criteria.SubjectText) & "%'"
        End If
        filterCount = filterCount + 1
    End If
    
    ' Date From
    If criteria.DateFrom <> "" Then
        If IsDate(criteria.DateFrom) Then
            Dim dateFrom As Date
            dateFrom = CDate(criteria.DateFrom)
            filters(filterCount) = "@SQL=""urn:schemas:httpmail:datereceived"" >= '" & Format(dateFrom, "mm/dd/yyyy") & " 00:00 AM'"
            filterCount = filterCount + 1
        End If
    End If
    
    ' Date To
    If criteria.DateTo <> "" Then
        If IsDate(criteria.DateTo) Then
            Dim dateTo As Date
            dateTo = CDate(criteria.DateTo)
            filters(filterCount) = "@SQL=""urn:schemas:httpmail:datereceived"" <= '" & Format(dateTo, "mm/dd/yyyy") & " 11:59 PM'"
            filterCount = filterCount + 1
        End If
    End If
    
    ' Combine filters with AND
    If filterCount = 0 Then
        BuildDASLFilter = ""
    ElseIf filterCount = 1 Then
        BuildDASLFilter = filters(0)
    Else
        Dim i As Integer
        Dim result As String
        result = filters(0)
        For i = 1 To filterCount - 1
            result = result & " AND " & filters(i)
        Next i
        BuildDASLFilter = result
    End If
End Function

Private Function GetSearchScope(folderName As String) As String
    ' Determine which folders to search based on selection
    ' Scope format: "'Folder1','Folder2','Folder3'"
    
    Select Case folderName
        Case "All Folders (including subfolders)"
            ' Search all mail folders in mailbox
            GetSearchScope = "'Inbox','Sent Items','Deleted Items','Drafts','Junk Email','Archive'," & _
                            "'01 Projects','02 Office','99 Misc','98 PR','97 Meeting'"
        
        Case "Inbox"
            GetSearchScope = "'Inbox'"
        
        Case "Sent Items"
            GetSearchScope = "'Sent Items'"
        
        Case "01 Projects"
            GetSearchScope = "'01 Projects'"
        
        Case "02 Office"
            GetSearchScope = "'02 Office'"
        
        Case "99 Misc"
            GetSearchScope = "'99 Misc'"
        
        Case "Archive"
            GetSearchScope = "'Archive'"
        
        Case Else
            ' Default to Inbox
            GetSearchScope = "'Inbox'"
    End Select
End Function

Private Function EscapeSQL(text As String) As String
    ' Escape special characters for DASL query
    ' Replace single quotes with two single quotes
    EscapeSQL = Replace(text, "'", "''")
End Function
