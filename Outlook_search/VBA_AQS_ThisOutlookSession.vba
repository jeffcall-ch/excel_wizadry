' ============================================
' Module: ThisOutlookSession (AQS Version 2.0)
' Purpose: Email search engine using Explorer.Search with AQS (Advanced Query Syntax)
' Features: Native Windows Search integration, Result capture, Bulk actions, Exports
' Version: 2.0.0 - Complete refactor from DASL to AQS
' Date: 2024-12-15
' ============================================

Option Explicit

' Constants
Private Const APP_VERSION As String = "2.0.0"
Private Const APP_DATE As String = "2024-12-15"

Private Const KB_IN_BYTES As Long = 1024
Private Const MB_IN_BYTES As Long = 1048576
Private Const GB_IN_BYTES As Long = 1073741824

' Module-level variables (AQS Architecture)
Private m_CapturedResults As Collection    ' Stores EntryIDs from captured search results
Private m_SearchExplorer As Outlook.Explorer ' Reference to Explorer window with search results
Private m_Namespace As Outlook.NameSpace    ' Cached namespace object
Private m_FormInstance As frmEmailSearch    ' Reference to search form

' ============================================
' MAIN ENTRY POINT
' ============================================

Public Sub ShowEmailSearchForm()
    On Error GoTo ErrorHandler
    
    Set m_FormInstance = New frmEmailSearch
    m_FormInstance.Show vbModeless
    Exit Sub
    
ErrorHandler:
    MsgBox "Error opening search form: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Common causes:" & vbCrLf & _
           "1. UserForm not created or not named 'frmEmailSearch'" & vbCrLf & _
           "2. UserForm code has compile errors" & vbCrLf & _
           "3. Missing controls on the form" & vbCrLf & vbCrLf & _
           "Press Alt+F11 to check the VBA editor.", vbCritical, "Cannot Open Search Form"
End Sub

' ============================================
' PUBLIC ACCESS FUNCTIONS
' ============================================

'-----------------------------------------------------------
' Function: GetCachedNamespace
' Purpose: Return cached Outlook namespace for performance
' Returns: Outlook.NameSpace object
' Notes: CACHED - avoids repeated GetNamespace calls
'-----------------------------------------------------------
Private Function GetCachedNamespace() As Outlook.NameSpace
    If m_Namespace Is Nothing Then
        Set m_Namespace = Application.GetNamespace("MAPI")
    End If
    Set GetCachedNamespace = m_Namespace
End Function

'-----------------------------------------------------------
' Function: GetCapturedResults
' Purpose: Return current captured results for form/external access
' Returns: Collection of EntryIDs (String)
' Notes: Used by form to check if results available for export/bulk actions
'-----------------------------------------------------------
Public Function GetCapturedResults() As Collection
    Set GetCapturedResults = m_CapturedResults
End Function

' ============================================
' MAIN SEARCH EXECUTION (AQS/Explorer.Search)
' ============================================

'-----------------------------------------------------------
' Sub: ExecuteEmailSearch
' Purpose: Execute UI search using Explorer.Search with AQS
' Parameters:
'   criteria (SearchCriteria) - User search criteria from form
' Notes: ASYNC - Results appear in Outlook UI, not returned directly
'        User must click "Capture Results" after reviewing in UI
'        NATIVE AQS - Leverages Windows Search index for speed
'-----------------------------------------------------------
Public Sub ExecuteEmailSearch(criteria As SearchCriteria)
    On Error GoTo ErrorHandler
    
    Dim targetFolder As Outlook.Folder
    Dim searchExplorer As Outlook.Explorer
    Dim searchScope As OlSearchScope
    Dim aqsQuery As String
    
    ' Resolve target folder
    Set targetFolder = ResolveFolder(criteria.FolderName)
    If targetFolder Is Nothing Then
        MsgBox "Cannot resolve folder: " & criteria.FolderName & vbCrLf & vbCrLf & _
               "Defaulting to Inbox.", vbExclamation, "Folder Not Found"
        Set targetFolder = GetCachedNamespace().GetDefaultFolder(olFolderInbox)
    End If
    
    ' Open Explorer for search (creates new window to avoid hijacking user's view)
    Set searchExplorer = OpenSearchExplorer(targetFolder)
    
    ' Build AQS query from criteria
    aqsQuery = BuildAQSQuery(criteria)
    
    ' Determine search scope (current folder, subfolders, mailbox, or all)
    searchScope = DetermineSearchScope(criteria)
    
    ' Debug logging (view in Immediate Window - Ctrl+G)
    Debug.Print ""
    Debug.Print "========== SEARCH STARTED (AQS) =========="
    Debug.Print "Timestamp: " & Now
    Debug.Print "Target Folder: " & targetFolder.FolderPath
    Debug.Print "Search Scope: " & GetScopeName(searchScope)
    Debug.Print "AQS Query: " & aqsQuery
    Debug.Print "Query Length: " & Len(aqsQuery) & " characters"
    Debug.Print "=========================================="
    Debug.Print ""
    
    ' Execute search (ASYNC - results appear in Outlook UI)
    ' NATIVE AQS: Uses Windows Search index for fast execution
    ' Note: Search must be non-empty to execute
    If Len(Trim(aqsQuery)) = 0 Then
        MsgBox "No search criteria specified.", vbExclamation, "Empty Search"
        Exit Sub
    End If
    
    searchExplorer.Search aqsQuery, searchScope
    
    ' Store reference for later result capture
    Set m_SearchExplorer = searchExplorer
    
    ' Clear any previous captured results
    Set m_CapturedResults = Nothing
    
    ' Notify user
    MsgBox "Search launched in Outlook." & vbCrLf & vbCrLf & _
           "Review the results in the Outlook window, then click 'Capture Results' to enable exports and bulk actions.", _
           vbInformation, "Search Launched"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print ""
    Debug.Print "!!! ERROR IN SEARCH EXECUTION !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "Folder: " & IIf(targetFolder Is Nothing, "(null)", targetFolder.FolderPath)
    Debug.Print "AQS Query: " & aqsQuery
    Debug.Print "================================"
    Debug.Print ""
    
    MsgBox "Error during search execution: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for details.", vbCritical, "Search Error"
End Sub

' ============================================
' RESULT CAPTURE (User-Initiated)
' ============================================

'-----------------------------------------------------------
' Function: CaptureCurrentSearchResults
' Purpose: Capture EntryIDs from search results using GetTable
' Parameters:
'   None - uses stored m_SearchExplorer
' Returns: Collection of EntryIDs (String)
' Notes: USER-INITIATED - Called when user clicks "Capture Results" button
'        Uses GetTable for efficient capture of large result sets
'-----------------------------------------------------------
Public Function CaptureCurrentSearchResults() As Collection
    On Error GoTo ErrorHandler
    
    Dim table As Outlook.table
    Dim row As Outlook.row
    Dim entryIDs As New Collection
    Dim rowCount As Long
    
    ' Validate Explorer exists
    If m_SearchExplorer Is Nothing Then
        MsgBox "No search Explorer available. Please run a search first.", vbCritical, "No Search Results"
        Set CaptureCurrentSearchResults = entryIDs
        Exit Function
    End If
    
    ' Get table from current view
    Set table = m_SearchExplorer.CurrentView.GetTable
    
    If table Is Nothing Then
        MsgBox "Cannot access search results. The search may still be running." & vbCrLf & vbCrLf & _
               "Wait a moment and try again.", vbExclamation, "Results Not Ready"
        Set CaptureCurrentSearchResults = entryIDs
        Exit Function
    End If
    
    ' Debug output
    Debug.Print ""
    Debug.Print "========== CAPTURING RESULTS =========="
    Debug.Print "Timestamp: " & Now
    
    ' Capture all EntryIDs
    rowCount = 0
    Do Until table.EndOfTable
        Set row = table.GetNextRow
        entryIDs.Add row("EntryID")
        rowCount = rowCount + 1
    Loop
    
    Debug.Print "Captured: " & rowCount & " result(s)"
    Debug.Print "======================================="
    Debug.Print ""
    
    ' Store for later access
    Set m_CapturedResults = entryIDs
    
    Set CaptureCurrentSearchResults = entryIDs
    Exit Function
    
ErrorHandler:
    Debug.Print ""
    Debug.Print "!!! CAPTURE RESULTS ERROR !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.Description
    Debug.Print "==============================="
    Debug.Print ""
    
    MsgBox "Error capturing results: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Common causes:" & vbCrLf & _
           "1. Search still running - wait and try again" & vbCrLf & _
           "2. Search window was closed" & vbCrLf & _
           "3. No results to capture", vbCritical, "Capture Error"
    
    Set CaptureCurrentSearchResults = New Collection
End Function

' ============================================
' FOLDER RESOLUTION & SCOPE
' ============================================

'-----------------------------------------------------------
' Function: ResolveFolder
' Purpose: Resolve folder name to Folder object with error handling
' Parameters:
'   folderName (String) - Folder name or path to resolve
' Returns: Outlook.Folder object (falls back to Inbox if not found)
' Notes: Supports common folder names and recursive custom folder search
'-----------------------------------------------------------
Private Function ResolveFolder(folderName As String) As Outlook.Folder
    On Error GoTo ErrorHandler
    
    Dim ns As Outlook.NameSpace
    Dim folder As Outlook.Folder
    Dim trimmedName As String
    
    Set ns = GetCachedNamespace()
    trimmedName = UCase(Trim(folderName))
    
    ' Handle empty or default
    If trimmedName = "" Or trimmedName = "---" Then
        Set folder = ns.GetDefaultFolder(olFolderInbox)
        Set ResolveFolder = folder
        Exit Function
    End If
    
    ' Check common system folders first
    Select Case trimmedName
        Case "INBOX"
            Set folder = ns.GetDefaultFolder(olFolderInbox)
        Case "SENT ITEMS", "SENT"
            Set folder = ns.GetDefaultFolder(olFolderSentMail)
        Case "DELETED ITEMS", "TRASH", "DELETED"
            Set folder = ns.GetDefaultFolder(olFolderDeletedItems)
        Case "DRAFTS"
            Set folder = ns.GetDefaultFolder(olFolderDrafts)
        Case "JUNK", "JUNK EMAIL", "JUNK E-MAIL"
            Set folder = ns.GetDefaultFolder(olFolderJunk)
        Case "OUTBOX"
            Set folder = ns.GetDefaultFolder(olFolderOutbox)
        Case Else
            ' Try to find custom folder recursively
            Set folder = FindFolderByName(ns.GetDefaultFolder(olFolderInbox).Parent, folderName)
    End Select
    
    ' Fallback to Inbox if not found
    If folder Is Nothing Then
        Debug.Print "WARNING: Folder '" & folderName & "' not found, using Inbox"
        Set folder = ns.GetDefaultFolder(olFolderInbox)
    End If
    
    Set ResolveFolder = folder
    Exit Function
    
ErrorHandler:
    Debug.Print "ResolveFolder Error: " & Err.Description & " (Folder: " & folderName & ")"
    On Error Resume Next
    Set ResolveFolder = ns.GetDefaultFolder(olFolderInbox)
End Function

'-----------------------------------------------------------
' Function: FindFolderByName
' Purpose: Recursively search for folder by name
' Parameters:
'   parentFolder (Folder) - Starting folder for search
'   searchName (String) - Folder name to find
' Returns: Outlook.Folder object or Nothing if not found
' Notes: Case-insensitive search
'-----------------------------------------------------------
Private Function FindFolderByName(parentFolder As Outlook.Folder, searchName As String) As Outlook.Folder
    On Error Resume Next
    
    Dim folder As Outlook.Folder
    Dim subFolder As Outlook.Folder
    
    ' Search direct children
    For Each folder In parentFolder.Folders
        If UCase(Trim(folder.Name)) = UCase(Trim(searchName)) Then
            Set FindFolderByName = folder
            Exit Function
        End If
    Next
    
    ' Search subfolders recursively
    For Each folder In parentFolder.Folders
        Set subFolder = FindFolderByName(folder, searchName)
        If Not subFolder Is Nothing Then
            Set FindFolderByName = subFolder
            Exit Function
        End If
    Next
    
    Set FindFolderByName = Nothing
End Function

'-----------------------------------------------------------
' Function: OpenSearchExplorer
' Purpose: Create new Explorer window for search results
' Parameters:
'   folder (Folder) - Folder to open in new Explorer
' Returns: Outlook.Explorer object
' Notes: Creates separate window to avoid hijacking user's current view
'-----------------------------------------------------------
Private Function OpenSearchExplorer(folder As Outlook.Folder) As Outlook.Explorer
    On Error GoTo ErrorHandler
    
    Dim newExplorer As Outlook.Explorer
    
    ' Create new Explorer for this folder
    Set newExplorer = Application.Explorers.Add(folder, olFolderDisplayNormal)
    newExplorer.Display
    
    ' Give Explorer time to initialize
    DoEvents
    
    Set OpenSearchExplorer = newExplorer
    Exit Function
    
ErrorHandler:
    Debug.Print "OpenSearchExplorer Error: " & Err.Description
    ' Try to use active Explorer as fallback
    Set OpenSearchExplorer = Application.ActiveExplorer
End Function

'-----------------------------------------------------------
' Function: DetermineSearchScope
' Purpose: Determine OlSearchScope from criteria
' Parameters:
'   criteria (SearchCriteria) - User search criteria
' Returns: OlSearchScope enum value
' Notes: Priority order: All Folders > Current Store > Subfolders > Current Folder
'-----------------------------------------------------------
Private Function DetermineSearchScope(criteria As SearchCriteria) As OlSearchScope
    Dim folderName As String
    folderName = UCase(Trim(criteria.FolderName))
    
    ' Priority 1: All Folders (entire Outlook)
    If folderName = "ALL FOLDERS" Or folderName = "EVERYWHERE" Or folderName = "ALL FOLDERS (INCLUDING SUBFOLDERS)" Then
        DetermineSearchScope = olSearchScopeAllFolders  ' = 3
        Exit Function
    End If
    
    ' Priority 2: Current Store (entire mailbox)
    If folderName = "CURRENT STORE" Or folderName = "THIS MAILBOX" Or folderName = "MAILBOX" Then
        DetermineSearchScope = olSearchScopeCurrentStore  ' = 2
        Exit Function
    End If
    
    ' Priority 3: Specific folder WITH subfolders
    If criteria.SearchSubFolders Then
        DetermineSearchScope = olSearchScopeSubfolders  ' = 1
        Exit Function
    End If
    
    ' Priority 4: Specific folder ONLY
    DetermineSearchScope = olSearchScopeCurrentFolder  ' = 0
End Function

'-----------------------------------------------------------
' Function: GetScopeName
' Purpose: Convert OlSearchScope enum to readable string for debug
' Parameters:
'   scope (OlSearchScope) - Scope enum value
' Returns: String description
' Notes: Helper for debug output
'-----------------------------------------------------------
Private Function GetScopeName(scope As OlSearchScope) As String
    Select Case scope
        Case olSearchScopeCurrentFolder
            GetScopeName = "Current Folder Only"
        Case olSearchScopeSubfolders
            GetScopeName = "Current Folder + Subfolders"
        Case olSearchScopeCurrentStore
            GetScopeName = "Current Mailbox (All Folders)"
        Case olSearchScopeAllFolders
            GetScopeName = "All Folders (All Accounts)"
        Case Else
            GetScopeName = "Unknown (" & scope & ")"
    End Select
End Function

' ============================================
' AQS QUERY BUILDER
' ============================================

'-----------------------------------------------------------
' Function: BuildAQSQuery
' Purpose: Convert SearchCriteria to AQS (Advanced Query Syntax) string
' Parameters:
'   criteria (SearchCriteria) - User search criteria from form
' Returns: AQS query string
' Notes: NATIVE AQS - Uses Windows Search index tokens
'        All field mappings: from, to, cc, subject, body, received, hasattachments, etc.
'-----------------------------------------------------------
Private Function BuildAQSQuery(criteria As SearchCriteria) As String
    On Error GoTo ErrorHandler
    
    Dim queryParts As Collection
    Set queryParts = New Collection
    
    ' FROM ADDRESS
    ' NATIVE AQS: from:value searches both display name and email address
    If Len(Trim(criteria.FromAddress)) > 0 Then
        queryParts.Add "from:" & EscapeAQSValue(criteria.FromAddress)
    End If
    
    ' TO ADDRESS
    ' NATIVE AQS: to:value searches recipient names/addresses
    If Len(Trim(criteria.ToAddress)) > 0 Then
        queryParts.Add "to:" & EscapeAQSValue(criteria.ToAddress)
    End If
    
    ' CC ADDRESS
    ' NATIVE AQS: cc:value searches CC recipients
    If Len(Trim(criteria.CcAddress)) > 0 Then
        queryParts.Add "cc:" & EscapeAQSValue(criteria.CcAddress)
    End If
    
    ' SUBJECT
    ' NATIVE AQS: subject:value searches subject line
    If Len(Trim(criteria.Subject)) > 0 Then
        queryParts.Add "subject:" & EscapeAQSValue(criteria.Subject)
    End If
    
    ' BODY
    ' NATIVE AQS: body:value searches message body
    If Len(Trim(criteria.Body)) > 0 Then
        queryParts.Add "body:" & EscapeAQSValue(criteria.Body)
    End If
    
    ' DATE RANGE
    ' NATIVE AQS: received:date1..date2 for ranges, received:>=date for comparisons
    ' Note: Check for date > 0 to avoid default 12/30/1899 (Date = 0)
    If criteria.DateFrom > 0 And criteria.DateTo > 0 Then
        ' Range syntax preferred for performance
        queryParts.Add "received:" & FormatAQSDate(criteria.DateFrom) & ".." & FormatAQSDate(criteria.DateTo)
    ElseIf criteria.DateFrom > 0 Then
        queryParts.Add "received:>=" & FormatAQSDate(criteria.DateFrom)
    ElseIf criteria.DateTo > 0 Then
        queryParts.Add "received:<=" & FormatAQSDate(criteria.DateTo)
    End If
    
    ' UNREAD STATUS
    ' NATIVE AQS: isread:no for unread, isread:yes for read
    If criteria.UnreadOnly Then
        queryParts.Add "isread:no"
    End If
    
    ' IMPORTANCE
    ' NATIVE AQS: importance:high/normal/low
    If criteria.ImportantOnly Then
        queryParts.Add "importance:high"
    End If
    
    ' ATTACHMENT FILTERS
    ' NATIVE AQS: hasattachments:yes/no, attachmentfilenames:pattern
    If Len(Trim(criteria.AttachmentFilter)) > 0 Then
        Select Case UCase(Trim(criteria.AttachmentFilter))
            Case "WITH ATTACHMENTS", "WITHATTACHMENTS", "YES"
                queryParts.Add "hasattachments:yes"
                
            Case "WITHOUT ATTACHMENTS", "WITHOUTATTACHMENTS", "NO", "NONE"
                queryParts.Add "hasattachments:no"
                
            Case "PDF", "PDF FILES"
                queryParts.Add "attachmentfilenames:*.pdf"
                
            Case "EXCEL", "EXCEL FILES (.XLS, .XLSX)", "XLS", "XLSX"
                queryParts.Add "(attachmentfilenames:*.xlsx OR attachmentfilenames:*.xls OR attachmentfilenames:*.xlsm)"
                
            Case "WORD", "WORD FILES (.DOC, .DOCX)", "DOC", "DOCX"
                queryParts.Add "(attachmentfilenames:*.docx OR attachmentfilenames:*.doc)"
                
            Case "IMAGES", "IMAGES (.JPG, .PNG, .GIF)", "JPG", "PNG", "GIF"
                queryParts.Add "(attachmentfilenames:*.jpg OR attachmentfilenames:*.jpeg OR attachmentfilenames:*.png OR attachmentfilenames:*.gif)"
                
            Case "POWERPOINT", "POWERPOINT FILES (.PPT, .PPTX)", "PPT", "PPTX"
                queryParts.Add "(attachmentfilenames:*.ppt OR attachmentfilenames:*.pptx)"
                
            Case "ZIP", "ZIP/ARCHIVES", "ARCHIVES"
                queryParts.Add "(attachmentfilenames:*.zip OR attachmentfilenames:*.rar OR attachmentfilenames:*.7z)"
        End Select
    End If
    
    ' SIZE FILTERS
    ' NATIVE AQS: size:>=value, size:value1..value2 (supports KB, MB, GB units)
    If Not IsEmpty(criteria.SizeFilter) Then
        Dim sizeFilterStr As String
        sizeFilterStr = UCase(Trim(criteria.SizeFilter))
        
        Select Case sizeFilterStr
            Case "< 100 KB"
                queryParts.Add "size:<100KB"
            Case "100 KB - 1 MB"
                queryParts.Add "size:100KB..1MB"
            Case "1 MB - 10 MB"
                queryParts.Add "size:1MB..10MB"
            Case "> 10 MB"
                queryParts.Add "size:>10MB"
        End Select
    End If
    
    ' GENERAL SEARCH TERMS
    ' NATIVE AQS: Space-separated keywords treated as implicit AND
    If Len(Trim(criteria.SearchTerms)) > 0 Then
        queryParts.Add Trim(criteria.SearchTerms)
    End If
    
    ' COMBINE ALL PARTS WITH AND
    If queryParts.Count = 0 Then
        BuildAQSQuery = ""
    Else
        BuildAQSQuery = JoinCollection(queryParts, " AND ")
    End If
    
    ' Validate query
    ValidateAQSQuery BuildAQSQuery
    
    Exit Function
    
ErrorHandler:
    Debug.Print "BuildAQSQuery Error: " & Err.Description
    BuildAQSQuery = ""
End Function

'-----------------------------------------------------------
' Function: EscapeAQSValue
' Purpose: Escape special characters in AQS values
' Parameters:
'   value (String) - Value to escape
' Returns: Escaped string
' Notes: Wraps in quotes if contains spaces or special chars
'        Escapes internal quotes by doubling them
'-----------------------------------------------------------
Private Function EscapeAQSValue(value As String) As String
    Dim needsQuotes As Boolean
    
    ' Check if value needs quoting
    needsQuotes = (InStr(value, " ") > 0) Or _
                  (InStr(value, ":") > 0) Or _
                  (InStr(value, "(") > 0) Or _
                  (InStr(value, ")") > 0) Or _
                  (InStr(value, """") > 0)
    
    If needsQuotes Then
        ' Escape internal quotes
        value = Replace(value, """", """""")
        EscapeAQSValue = """" & value & """"
    Else
        EscapeAQSValue = value
    End If
End Function

'-----------------------------------------------------------
' Function: FormatAQSDate
' Purpose: Format date for AQS syntax
' Parameters:
'   dateValue (Date) - Date to format
' Returns: String in MM/DD/YYYY format (AQS standard)
' Notes: AQS accepts MM/DD/YYYY format
'-----------------------------------------------------------
Private Function FormatAQSDate(dateValue As Date) As String
    FormatAQSDate = Format(dateValue, "mm/dd/yyyy")
End Function

'-----------------------------------------------------------
' Sub: ValidateAQSQuery
' Purpose: Check AQS query for common syntax errors
' Parameters:
'   query (String) - AQS query to validate
' Notes: Logs warnings to Debug.Print if issues found
'-----------------------------------------------------------
Private Sub ValidateAQSQuery(query As String)
    ' Check for double colons
    If InStr(query, "::") > 0 Then
        Debug.Print "WARNING: Double colon detected in AQS query - may fail"
    End If
    
    ' Check for unbalanced parentheses
    Dim openCount As Integer
    Dim closeCount As Integer
    openCount = Len(query) - Len(Replace(query, "(", ""))
    closeCount = Len(query) - Len(Replace(query, ")", ""))
    
    If openCount <> closeCount Then
        Debug.Print "WARNING: Unbalanced parentheses in AQS query (Open: " & openCount & ", Close: " & closeCount & ")"
    End If
    
    ' Check for unbalanced quotes
    Dim quoteCount As Integer
    quoteCount = Len(query) - Len(Replace(query, """", ""))
    
    If quoteCount Mod 2 <> 0 Then
        Debug.Print "WARNING: Unbalanced quotes in AQS query (Count: " & quoteCount & ")"
    End If
End Sub

'-----------------------------------------------------------
' Function: JoinCollection
' Purpose: Join Collection items into delimited string
' Parameters:
'   col (Collection) - Collection to join
'   delimiter (String) - Delimiter between items
' Returns: Joined string
' Notes: Helper function for building query strings
'-----------------------------------------------------------
Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim result As String
    Dim item As Variant
    
    For Each item In col
        If Len(result) > 0 Then result = result & delimiter
        result = result & CStr(item)
    Next
    
    JoinCollection = result
End Function

' ============================================
' BULK ACTIONS (EntryID-based)
' ============================================

'-----------------------------------------------------------
' Sub: BulkMarkAsRead
' Purpose: Mark all captured results as read
' Notes: Uses m_CapturedResults Collection
'        Continues on individual item failures
'-----------------------------------------------------------
Public Sub BulkMarkAsRead()
    On Error Resume Next
    
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim successCount As Long
    Dim failCount As Long
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Mark all " & m_CapturedResults.Count & " captured emails as read?", _
                      vbYesNo + vbQuestion, "Bulk Mark as Read")
    
    If response <> vbYes Then Exit Sub
    
    Set ns = GetCachedNamespace()
    
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            mailItem.UnRead = False
            mailItem.Save
            successCount = successCount + 1
        Else
            failCount = failCount + 1
            Debug.Print "Failed to retrieve item: " & entryID
        End If
        Set mailItem = Nothing
    Next
    
    Debug.Print "Bulk Mark as Read: " & successCount & " succeeded, " & failCount & " failed"
    
    MsgBox "Marked " & successCount & " item(s) as read." & vbCrLf & _
           IIf(failCount > 0, failCount & " item(s) could not be processed.", ""), _
           vbInformation, "Bulk Action Complete"
End Sub

'-----------------------------------------------------------
' Sub: BulkFlagItems
' Purpose: Flag all captured results
' Notes: Uses m_CapturedResults Collection
'-----------------------------------------------------------
Public Sub BulkFlagItems()
    On Error Resume Next
    
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim successCount As Long
    Dim failCount As Long
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Flag all " & m_CapturedResults.Count & " captured emails?", _
                      vbYesNo + vbQuestion, "Bulk Flag Items")
    
    If response <> vbYes Then Exit Sub
    
    Set ns = GetCachedNamespace()
    
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            mailItem.FlagStatus = olFlagMarked
            mailItem.Save
            successCount = successCount + 1
        Else
            failCount = failCount + 1
        End If
        Set mailItem = Nothing
    Next
    
    Debug.Print "Bulk Flag: " & successCount & " succeeded, " & failCount & " failed"
    
    MsgBox "Flagged " & successCount & " item(s)." & vbCrLf & _
           IIf(failCount > 0, failCount & " item(s) failed.", ""), _
           vbInformation, "Bulk Action Complete"
End Sub

'-----------------------------------------------------------
' Sub: BulkMoveItems
' Purpose: Move all captured results to target folder
' Notes: Uses m_CapturedResults Collection
'        Uses folder picker dialog
'-----------------------------------------------------------
Public Sub BulkMoveItems()
    On Error Resume Next
    
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim targetFolder As Outlook.Folder
    Dim successCount As Long
    Dim failCount As Long
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    Set ns = GetCachedNamespace()
    Set targetFolder = ns.PickFolder
    
    If targetFolder Is Nothing Then Exit Sub
    
    Dim response As VbMsgBoxResult
    response = MsgBox("Move all " & m_CapturedResults.Count & " emails to '" & targetFolder.Name & "'?", _
                      vbYesNo + vbQuestion, "Bulk Move Items")
    
    If response <> vbYes Then Exit Sub
    
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            mailItem.Move targetFolder
            successCount = successCount + 1
        Else
            failCount = failCount + 1
        End If
        Set mailItem = Nothing
    Next
    
    Debug.Print "Bulk Move: " & successCount & " succeeded, " & failCount & " failed"
    
    MsgBox "Moved " & successCount & " item(s) to '" & targetFolder.Name & "'." & vbCrLf & _
           IIf(failCount > 0, failCount & " item(s) failed.", ""), _
           vbInformation, "Bulk Action Complete"
End Sub

' ============================================
' EXPORT FUNCTIONS (EntryID-based)
' ============================================

'-----------------------------------------------------------
' Function: ExportSearchResultsToCSV
' Purpose: Export captured results to CSV file
' Returns: Boolean (success/failure)
' Notes: Exports to Desktop by default
'        Uses m_CapturedResults Collection
'-----------------------------------------------------------
Public Function ExportSearchResultsToCSV() As Boolean
    On Error GoTo ErrorHandler
    
    Dim fso As Object
    Dim txtFile As Object
    Dim filePath As String
    Dim desktopPath As String
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim line As String
    Dim rowCount As Long
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results to export. Click 'Capture Results' first.", vbExclamation, "No Results"
        ExportSearchResultsToCSV = False
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ns = GetCachedNamespace()
    
    desktopPath = Environ("USERPROFILE") & "\Desktop"
    If Not fso.FolderExists(desktopPath) Then
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    End If
    
    filePath = desktopPath & "\EmailSearchResults_" & Format(Now, "yyyymmdd_hhnnss") & ".csv"
    
    Set txtFile = fso.CreateTextFile(filePath, True, True)
    
    ' Write header
    txtFile.WriteLine "Date,From,To,Subject,Size,Attachments,Unread,Flagged"
    
    ' Write data
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            line = Format(mailItem.ReceivedTime, "yyyy-mm-dd hh:nn") & "," & _
                   EscapeCSV(mailItem.SenderName) & "," & _
                   EscapeCSV(mailItem.To) & "," & _
                   EscapeCSV(mailItem.Subject) & "," & _
                   mailItem.Size & "," & _
                   mailItem.Attachments.Count & "," & _
                   IIf(mailItem.UnRead, "Yes", "No") & "," & _
                   IIf(mailItem.FlagStatus = olFlagMarked, "Yes", "No")
            
            txtFile.WriteLine line
            rowCount = rowCount + 1
        End If
        Set mailItem = Nothing
    Next
    
    txtFile.Close
    
    MsgBox "Exported " & rowCount & " email(s) to:" & vbCrLf & vbCrLf & filePath, vbInformation, "Export Complete"
    
    ' Open file location
    On Error Resume Next
    CreateObject("WScript.Shell").Run "explorer.exe /select," & Chr(34) & filePath & Chr(34)
    
    ExportSearchResultsToCSV = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not txtFile Is Nothing Then txtFile.Close
    MsgBox "Error during export: " & Err.Description, vbCritical, "Export Error"
    ExportSearchResultsToCSV = False
End Function

'-----------------------------------------------------------
' Function: ExportSearchResultsToExcel
' Purpose: Export captured results to Excel file
' Returns: Boolean (success/failure)
' Notes: Creates formatted Excel workbook with AutoFilter
'        Uses m_CapturedResults Collection
'-----------------------------------------------------------
Public Function ExportSearchResultsToExcel() As Boolean
    On Error GoTo ErrorHandler
    
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim row As Long
    Dim filePath As String
    Dim desktopPath As String
    Dim fso As Object
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results to export. Click 'Capture Results' first.", vbExclamation, "No Results"
        ExportSearchResultsToExcel = False
        Exit Function
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ns = GetCachedNamespace()
    
    desktopPath = Environ("USERPROFILE") & "\Desktop"
    If Not fso.FolderExists(desktopPath) Then
        desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    End If
    
    ' Create Excel application
    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    
    ' Set headers
    With xlSheet
        .Cells(1, 1).Value = "Date"
        .Cells(1, 2).Value = "From"
        .Cells(1, 3).Value = "To"
        .Cells(1, 4).Value = "Subject"
        .Cells(1, 5).Value = "Size (KB)"
        .Cells(1, 6).Value = "Attachments"
        .Cells(1, 7).Value = "Unread"
        .Cells(1, 8).Value = "Flagged"
        
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.Color = RGB(200, 200, 200)
    End With
    
    ' Populate data
    row = 2
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            With xlSheet
                .Cells(row, 1).Value = mailItem.ReceivedTime
                .Cells(row, 2).Value = SanitizeForExcel(mailItem.SenderName)
                .Cells(row, 3).Value = SanitizeForExcel(mailItem.To)
                .Cells(row, 4).Value = SanitizeForExcel(mailItem.Subject)
                .Cells(row, 5).Value = Round(mailItem.Size / 1024, 2)
                .Cells(row, 6).Value = mailItem.Attachments.Count
                .Cells(row, 7).Value = IIf(mailItem.UnRead, "Yes", "No")
                .Cells(row, 8).Value = IIf(mailItem.FlagStatus = olFlagMarked, "Yes", "No")
            End With
            row = row + 1
        End If
        Set mailItem = Nothing
    Next
    
    ' Auto-fit columns
    xlSheet.Columns.AutoFit
    
    ' Add filter
    xlSheet.Range("A1:H" & row - 1).AutoFilter
    
    ' Save file
    filePath = desktopPath & "\EmailSearchResults_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
    xlBook.SaveAs filePath
    
    ' Show Excel
    xlApp.Visible = True
    
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    
    MsgBox "Exported " & (row - 2) & " email(s) to Excel:" & vbCrLf & vbCrLf & filePath, vbInformation, "Export Complete"
    
    ExportSearchResultsToExcel = True
    Exit Function
    
ErrorHandler:
    On Error Resume Next
    If Not xlApp Is Nothing Then
        xlApp.DisplayAlerts = False
        xlApp.Quit
    End If
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
    MsgBox "Error during Excel export: " & Err.Description, vbCritical, "Export Error"
    ExportSearchResultsToExcel = False
End Function

'-----------------------------------------------------------
' Function: EscapeCSV
' Purpose: Escape commas and quotes for CSV format
' Parameters:
'   text (String) - Text to escape
' Returns: Escaped string
'-----------------------------------------------------------
Private Function EscapeCSV(text As String) As String
    Dim result As String
    result = Replace(text, Chr(34), Chr(34) & Chr(34))
    If InStr(result, ",") > 0 Or InStr(result, Chr(34)) > 0 Or InStr(result, vbCrLf) > 0 Then
        result = Chr(34) & result & Chr(34)
    End If
    EscapeCSV = result
End Function

'-----------------------------------------------------------
' Function: SanitizeForExcel
' Purpose: Prevent CSV/Excel formula injection
' Parameters:
'   text (String) - Text to sanitize
' Returns: Sanitized string
' Notes: Prefixes dangerous characters (= + - @) with single quote
'-----------------------------------------------------------
Private Function SanitizeForExcel(text As String) As String
    If Len(text) > 0 Then
        Dim firstChar As String
        firstChar = Left(text, 1)
        If firstChar = "=" Or firstChar = "+" Or firstChar = "-" Or firstChar = "@" Then
            SanitizeForExcel = "'" & text
        Else
            SanitizeForExcel = text
        End If
    Else
        SanitizeForExcel = text
    End If
End Function

' ============================================
' STATISTICS (EntryID-based)
' ============================================

'-----------------------------------------------------------
' Sub: ShowSearchStatistics
' Purpose: Display statistics for captured results
' Notes: Uses m_CapturedResults Collection
'-----------------------------------------------------------
Public Sub ShowSearchStatistics()
    On Error Resume Next
    
    Dim entryID As Variant
    Dim mailItem As Outlook.mailItem
    Dim totalCount As Long
    Dim unreadCount As Long
    Dim flaggedCount As Long
    Dim withAttachments As Long
    Dim totalSize As Long
    Dim senders As Object
    Dim senderName As String
    Dim msg As String
    Dim topSender As String
    Dim topSenderCount As Long
    Dim key As Variant
    Dim ns As Outlook.NameSpace
    
    If m_CapturedResults Is Nothing Or m_CapturedResults.Count = 0 Then
        MsgBox "No results captured. Click 'Capture Results' first.", vbExclamation, "No Results"
        Exit Sub
    End If
    
    Set ns = GetCachedNamespace()
    Set senders = CreateObject("Scripting.Dictionary")
    totalCount = m_CapturedResults.Count
    
    ' Analyze results
    For Each entryID In m_CapturedResults
        Set mailItem = ns.GetItemFromID(CStr(entryID))
        If Not mailItem Is Nothing Then
            If mailItem.UnRead Then unreadCount = unreadCount + 1
            If mailItem.FlagStatus = olFlagMarked Then flaggedCount = flaggedCount + 1
            If mailItem.Attachments.Count > 0 Then withAttachments = withAttachments + 1
            
            totalSize = totalSize + mailItem.Size
            
            senderName = mailItem.SenderName
            If senders.Exists(senderName) Then
                senders(senderName) = senders(senderName) + 1
            Else
                senders.Add senderName, 1
            End If
        End If
        Set mailItem = Nothing
    Next
    
    ' Find top sender
    For Each key In senders.Keys
        If senders(key) > topSenderCount Then
            topSenderCount = senders(key)
            topSender = key
        End If
    Next
    
    ' Build statistics message
    msg = "SEARCH STATISTICS" & vbCrLf & vbCrLf
    msg = msg & "Total Emails: " & totalCount & vbCrLf
    
    If totalCount > 0 Then
        msg = msg & "Unread: " & unreadCount & " (" & Format(unreadCount / totalCount * 100, "0.0") & "%)" & vbCrLf
        msg = msg & "Flagged: " & flaggedCount & " (" & Format(flaggedCount / totalCount * 100, "0.0") & "%)" & vbCrLf
        msg = msg & "With Attachments: " & withAttachments & " (" & Format(withAttachments / totalCount * 100, "0.0") & "%)" & vbCrLf & vbCrLf
        msg = msg & "Total Size: " & FormatSize(totalSize) & vbCrLf
        msg = msg & "Average Size: " & FormatSize(totalSize / totalCount) & vbCrLf & vbCrLf
        msg = msg & "Unique Senders: " & senders.Count & vbCrLf
        msg = msg & "Most Common Sender: " & topSender & " (" & topSenderCount & " emails)"
    Else
        msg = msg & "No results to analyze."
    End If
    
    MsgBox msg, vbInformation, "Search Statistics"
End Sub

'-----------------------------------------------------------
' Function: FormatSize
' Purpose: Convert bytes to human-readable format
' Parameters:
'   bytes (Long) - Size in bytes
' Returns: Formatted string (KB, MB, GB)
'-----------------------------------------------------------
Private Function FormatSize(bytes As Long) As String
    If bytes >= GB_IN_BYTES Then
        FormatSize = Format(bytes / GB_IN_BYTES, "0.00") & " GB"
    ElseIf bytes >= MB_IN_BYTES Then
        FormatSize = Format(bytes / MB_IN_BYTES, "0.00") & " MB"
    ElseIf bytes >= KB_IN_BYTES Then
        FormatSize = Format(bytes / KB_IN_BYTES, "0.00") & " KB"
    Else
        FormatSize = bytes & " bytes"
    End If
End Function

' ============================================
' PRESET MANAGEMENT (Placeholder for future)
' ============================================

'-----------------------------------------------------------
' NOTE: Preset save/load requires JSON file implementation
'       or registry access. This is a placeholder for future enhancement.
'-----------------------------------------------------------

Public Sub SaveSearchPreset(presetName As String, criteria As SearchCriteria)
    MsgBox "Preset save functionality requires JSON file implementation. Feature not yet available.", vbInformation, "Feature Not Implemented"
End Sub

Public Function LoadSearchPreset(presetName As String, ByRef criteria As SearchCriteria) As Boolean
    LoadSearchPreset = False
    MsgBox "Preset load functionality requires JSON file implementation. Feature not yet available.", vbInformation, "Feature Not Implemented"
End Function
